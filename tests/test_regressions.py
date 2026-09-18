"""Local regression cases: no network, credentials, or generated-image calls."""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
import artifact_contracts as ac
import local_runtime as rt
import offline_prepare as op
import page_reuse as reuse
import preflight as pf
import project_state as ps
import speech_timing as timing
import validate_delivery as vd


def result():
    return {"errors": [], "warnings": [], "metrics": {}}


class LocalCase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="ppt-gen-test-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def write(self, name, value):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(value if isinstance(value, str) else json.dumps(value, ensure_ascii=False), encoding="utf-8")
        return path

    def state_command(self, *args, success=True):
        completed = subprocess.run([sys.executable, "-B", str(SCRIPTS / "project_state.py"), *map(str, args)], capture_output=True, text=True)
        if success:
            self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        else:
            self.assertNotEqual(completed.returncode, 0)
        return completed


class RuntimeTests(LocalCase):
    def test_app_fallback_is_shared_with_preflight(self):
        app = self.write("app-soffice", "test executable")
        app.chmod(0o700)
        with patch.dict(os.environ, {"PPT_GEN_LIBREOFFICE": ""}), patch.object(rt.shutil, "which", return_value=None), patch.dict(rt.TOOL_PATHS, {"libreoffice": (str(app),)}), patch.object(pf, "probe_command", side_effect=lambda path, args, timeout: {"path": path}):
            self.assertEqual(pf.detect_libreoffice(1)["path"], rt.locate_tool("libreoffice"))
            self.assertEqual(rt.locate_tool("libreoffice"), str(app.resolve()))

    def test_invalid_override_does_not_fall_back(self):
        with patch.dict(os.environ, {"PPT_GEN_LIBREOFFICE": str(self.root / "missing")}):
            self.assertIsNone(rt.locate_tool("libreoffice"))

    def test_cache_hit_and_corruption_rebuild(self):
        cache = rt.LocalCache(self.root / "cache")
        calls = []
        def build(stage):
            calls.append(1)
            path = stage / "page.png"
            path.write_bytes(b"test pixels")
            return [path]
        first = cache.get("render", {"input": "a"}, build)
        self.assertEqual(first, cache.get("render", {"input": "a"}, build))
        self.assertEqual(len(calls), 1)
        first[0].write_bytes(b"corrupt")
        cache.get("render", {"input": "a"}, build)
        self.assertEqual(len(calls), 2)
        cache.get("render", {"input": "b"}, build)
        self.assertEqual(len(calls), 3)

    def test_failed_cache_builder_is_not_cached(self):
        cache = rt.LocalCache(self.root / "cache")
        with self.assertRaises(RuntimeError):
            cache.get("broken", {}, lambda stage: [])
        self.assertFalse(list(cache.root.glob("*/cache.json")))

    def test_malformed_cache_metadata_rebuilds(self):
        cache = rt.LocalCache(self.root / "cache")
        def build(stage):
            path = stage / "output.txt"
            path.write_text("valid")
            return [path]
        first = cache.get("metadata", {}, build)
        (first[0].parent / "cache.json").write_text("[]")
        self.assertEqual(cache.get("metadata", {}, build)[0].read_text(), "valid")

    def test_default_cache_is_cleaned_on_process_exit(self):
        code = "import local_runtime; print(local_runtime.cache().root)"
        done = subprocess.run([sys.executable, "-B", "-c", code], cwd=SCRIPTS, capture_output=True, text=True, check=True)
        self.assertFalse(Path(done.stdout.strip()).exists())

    def test_ocr_never_falls_back_on_non_mac(self):
        with patch.object(rt.platform, "system", return_value="Windows"):
            with self.assertRaisesRegex(RuntimeError, "no online fallback"):
                rt.recognize_image(self.root / "image.png")


class OfflineTests(LocalCase):
    def handoff(self):
        images = [op.binding(self.write("01.png", "synthetic image"))]
        data = {"schema_version": 1, "ocr_policy": "offline", "ask_for_ocr_token": False,
                "data_handling": "standard", "image_backend": "builtin-allowed", "source_images": images,
                "source_policy": "faithful", "hashes": {}}
        for key in ("source_manifest", "slide_copy_ledger", "slide_manifest"):
            path = self.write(key + ".json", {"schema_version": 1})
            data[key] = str(path)
            data["hashes"][key] = rt.sha256(path)
        return data

    def test_offline_flag_even_when_synthetic_token_is_set(self):
        with patch.dict(os.environ, {"PADDLE_OCR_TOKEN": "synthetic-not-a-credential"}), patch.object(op, "locate_tool", return_value="/test/editppt"):
            command = op.prepare_command(self.handoff(), self.root / "run")
        self.assertIn("--no-text-hints", command)
        self.assertNotIn("synthetic-not-a-credential", " ".join(command))
        self.assertNotIn("run hints", " ".join(command))

    def test_changed_source_blocks_prepare(self):
        data = self.handoff()
        Path(data["source_images"][0]["path"]).write_text("changed")
        with self.assertRaisesRegex(ValueError, "hash-mismatched"):
            op.prepare_command(data, self.root / "run")

    def test_online_or_unknown_policy_blocks_prepare(self):
        data = self.handoff()
        data["ocr_policy"] = "online"
        with self.assertRaises(ValueError):
            op.prepare_command(data, self.root / "run")

    def test_local_only_does_not_configure_remote_backend(self):
        data = self.handoff()
        data.update(data_handling="local-only", image_backend="disabled")
        with self.assertRaisesRegex(ValueError, "no disabled"):
            op.prepare_command(data, self.root / "run")

    def test_missing_execution_report_rejected(self):
        self.assertTrue(op.validate_report(self.root / "missing.json", None))

    def test_final_editable_gate_requires_execution_record(self):
        args = vd.build_parser().parse_args(["--editable-pptx", "synthetic.pptx"])
        output = result()
        vd.enforce_final_contract(args, output)
        self.assertTrue(any("--offline-preparation" in message for message in output["errors"]))


class StateTests(LocalCase):
    def qa_state(self, typed=True):
        project = self.root / "project"
        self.state_command("init", project, "--start-at", "qa-package", "--stop-after", "qa-package", "--deliverable", "qa-report")
        document = self.write("existing.docx", "route-only fixture")
        flags = ["--artifact-role", "report"] if typed else []
        self.state_command("add-input", project, "existing=" + str(document), "--stage", "qa-package", *flags)
        return project, document, json.loads((project / ps.MANIFEST_NAME).read_text())

    def test_qa_only_accepts_typed_existing_artifact(self):
        _, doc, state = self.qa_state()
        output = result()
        vd.validate_state_artifact_linkage(state, [{"role": "report", **op.binding(doc)}], output)
        self.assertEqual(output["errors"], [])
        self.assertEqual(state["stages"]["report"]["status"], "skipped")
        self.assertEqual(state["deliverables"], ["qa-report"])

    def test_untyped_input_does_not_bypass_production_gate(self):
        _, doc, state = self.qa_state(typed=False)
        output = result()
        vd.validate_state_artifact_linkage(state, [{"role": "report", **op.binding(doc)}], output)
        self.assertTrue(output["errors"])

    def test_import_type_cannot_be_changed_at_validation(self):
        _, doc, state = self.qa_state()
        self.assertFalse(ac.matching_import(state, {"role": "editable-pptx", **op.binding(doc)}))

    def test_import_drift_requires_revalidation(self):
        project, doc, state = self.qa_state()
        doc.write_text("changed authority")
        ps.detect_drift(state)
        self.assertEqual(state["stages"]["qa-package"]["status"], "stale")
        self.assertFalse(ac.matching_import(state, {"role": "report", **op.binding(doc)}))

    def test_imported_artifact_is_not_new_output(self):
        self.state_command("init", self.root / "bad", "--start-at", "qa-package", "--stop-after", "qa-package", "--deliverable", "report-docx", success=False)

    def test_normal_routes_remain_independent(self):
        for index, (start, stop, role) in enumerate((("report", "report", "report-docx"), ("outline", "qa-package", "editable-pptx"), ("speaker-notes", "speaker-notes", "speaker-notes-docx"))):
            project = self.root / str(index)
            self.state_command("init", project, "--start-at", start, "--stop-after", stop, "--deliverable", role)
            state = json.loads((project / ps.MANIFEST_NAME).read_text())
            self.assertEqual(state["current_stage"], start)
            self.assertEqual(state["deliverables"], [role])

    def test_pause_and_resume(self):
        project, _, _ = self.qa_state()
        self.state_command("pause", project)
        self.state_command("resume", project)
        state = json.loads((project / ps.MANIFEST_NAME).read_text())
        self.assertEqual(state["current_stage"], "qa-package")
        self.assertEqual(state["run_status"], "active")

    def test_duration_intake_does_not_mark_unstarted_work_stale(self):
        project = self.root / "notes"
        self.state_command("init", project, "--start-at", "speaker-notes", "--stop-after", "speaker-notes", "--deliverable", "speaker-notes-docx")
        self.state_command("record", project, "--preference", "requested_duration_seconds=600")
        state = json.loads((project / ps.MANIFEST_NAME).read_text())
        self.assertEqual(state["preferences"]["requested_duration_seconds"], 600)
        self.assertEqual(state["stages"]["speaker-notes"]["status"], "pending")

    def test_adopt_keeps_imported_qa_role(self):
        project, doc, _ = self.qa_state()
        self.state_command("adopt", project, "qa-package", "--input", "existing=" + str(doc))
        state = json.loads((project / ps.MANIFEST_NAME).read_text())
        self.assertEqual(state["inputs"]["existing"]["artifact_role"], "report")

    def test_malformed_import_role_is_reported_not_crashed(self):
        _, _, state = self.qa_state()
        state["inputs"]["existing"]["artifact_role"] = []
        self.assertTrue(ps.manifest_errors(state))

    def test_historical_qa_contract_is_readable_without_weakening_new_gate(self):
        self.assertNotIn("offline-preparation", ps.required_qa_roles("editable-pptx", 1))
        self.assertIn("offline-preparation", ps.required_qa_roles("editable-pptx", 2))
        self.assertIn("editable-validation", ps.required_qa_roles("editable-pptx", 1))


class TimingTests(unittest.TestCase):
    def check(self, script, target, settings=None, requested=None, **extra):
        entry = {"page": 1, "script": script, "target_seconds": target, **extra}
        output = result()
        timing.validate_timing({"timing": settings or {}}, [entry], output, requested)
        return output

    def test_2400_characters_cannot_be_labeled_sixty_seconds(self):
        self.assertTrue(self.check("测" * 2400, 60)["errors"])

    def test_natural_chinese_duration(self):
        self.assertEqual(self.check("测" * 260, 60)["errors"], [])

    def test_natural_english_duration(self):
        self.assertEqual(self.check("word " * 150, 60)["errors"], [])

    def test_explicit_pauses_are_accounted_for(self):
        output = self.check("测" * 130, 60, pause_seconds=30)
        self.assertEqual(output["errors"], [])
        self.assertEqual(output["metrics"]["speaker_timing"]["estimated_total_seconds"], 60)

    def test_user_request_is_not_replaced_by_matching_labels(self):
        self.assertTrue(self.check("测" * 260, 60, requested=600)["errors"])

    def test_conflicting_or_invalid_settings_fail(self):
        for settings in ({"requested_seconds": 30}, {"cjk_chars_per_minute": float("nan")}, {"words_per_minute": True}):
            self.assertTrue(self.check("测" * 260, 60, settings, requested=60)["errors"])

    def test_legacy_duration_alias_is_supported(self):
        output = result()
        timing.validate_timing({}, [{"script": "测" * 260, "duration": "1分钟"}], output)
        self.assertEqual(output["errors"], [])

    def test_nonfinite_time_cannot_pass(self):
        self.assertIsNone(timing.parse_duration_seconds(float("inf")))
        self.assertIsNone(timing.parse_duration_seconds(float("nan")))


class ReuseTests(LocalCase):
    def fixture(self):
        ledger = {"slides": [{"page": i, "id": f"slide-{i}", "body_copy": [f"page {i}"]} for i in (1, 2)]}
        manifest = {"slides": [{"page": i, "id": f"slide-{i}", **op.binding(self.write(f"{i}.png", str(i)))} for i in (1, 2)]}
        visual = self.write("visual.json", {"palette": ["black"]})
        qa = self.write("qa.json", {"mode": "final-delivery", "passed": True, "errors": []})
        return ledger, visual, reuse.make_index(ledger, visual, manifest, qa)

    def test_only_changed_page_is_regenerated(self):
        ledger, visual, index = self.fixture()
        ledger["slides"][1]["body_copy"] = ["changed"]
        actions = reuse.plan(ledger, visual, index)
        self.assertEqual([p["action"] for p in actions["pages"]], ["reuse-candidate", "regenerate"])
        self.assertTrue(actions["requires_current_final_qa"])

    def test_global_style_change_invalidates_all_pages(self):
        ledger, visual, index = self.fixture()
        visual.write_text("changed")
        self.assertTrue(all(p["action"] == "regenerate" for p in reuse.plan(ledger, visual, index)["pages"]))

    def test_corrupted_image_is_not_reused(self):
        ledger, visual, index = self.fixture()
        Path(index["pages"][0]["image"]["path"]).write_text("changed")
        self.assertEqual(reuse.plan(ledger, visual, index)["pages"][0]["action"], "regenerate")

    def test_diagnostic_qa_cannot_seed_reuse(self):
        ledger, visual, _ = self.fixture()
        qa = self.write("diagnostic.json", {"mode": "structural-only", "passed": True})
        with self.assertRaises(ValueError):
            reuse.make_index(ledger, visual, {"slides": []}, qa)


if __name__ == "__main__":
    unittest.main()
