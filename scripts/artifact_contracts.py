"""Shared artifact role semantics for production and imported QA targets."""
from pathlib import Path


CLI_DELIVERABLE_ROLES = {
    "report": ("report-docx", "report"),
    "outline": ("outline", "outline"),
    "style-options": ("style-options", "style-options"),
    "template-overview": ("template-overview", "template"),
    "slide-images": ("slide-images", "image-deck"),
    "image-pptx": ("image-pptx", "image-deck"),
    "editable-pptx": ("editable-pptx", "editable-deck"),
    "revised-pptx": ("revised-pptx", "deck-revision"),
    "speaker-notes": ("speaker-notes-docx", "speaker-notes"),
    "ppt-notes-pptx": ("ppt-notes", "speaker-notes"),
}


def matching_import(state: dict, artifact: dict) -> bool:
    """Only an explicitly typed QA input may stand in for a producing stage."""
    if not isinstance(state.get("targets"), list) or "qa-package" not in state["targets"]:
        return False
    inputs = state.get("inputs")
    if not isinstance(inputs, dict):
        return False
    for record in inputs.values():
        if (isinstance(record, dict) and record.get("artifact_role") == artifact.get("role")
                and record.get("authority_stage") == "qa-package"
                and record.get("status") == "current" and record.get("exists") is True
                and isinstance(record.get("path"), str) and isinstance(record.get("sha256"), str)
                and isinstance(artifact.get("path"), str) and isinstance(artifact.get("sha256"), str)
                and Path(record["path"]).resolve() == Path(artifact["path"]).resolve()
                and record["sha256"].casefold() == artifact["sha256"].casefold()):
            return True
    return False
