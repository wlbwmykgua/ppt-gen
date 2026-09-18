import AppKit
import Foundation
import Vision

guard CommandLine.arguments.count == 2 else {
    fputs("usage: offline_ocr.swift /absolute/image.png\n", stderr)
    exit(2)
}

let path = CommandLine.arguments[1]
guard let image = NSImage(contentsOfFile: path),
      let data = image.tiffRepresentation,
      let bitmap = NSBitmapImageRep(data: data),
      let cgImage = bitmap.cgImage else {
    fputs("cannot decode image: \(path)\n", stderr)
    exit(1)
}

let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.usesLanguageCorrection = true
if #available(macOS 13.0, *) {
    request.automaticallyDetectsLanguage = true
}

do {
    try VNImageRequestHandler(cgImage: cgImage, options: [:]).perform([request])
    let observations = request.results ?? []
    let ordered = observations.sorted {
        let yDifference = abs($0.boundingBox.midY - $1.boundingBox.midY)
        if yDifference > 0.02 { return $0.boundingBox.midY > $1.boundingBox.midY }
        return $0.boundingBox.minX < $1.boundingBox.minX
    }
    let lines = ordered.compactMap { $0.topCandidates(1).first?.string }
    let blocks: [[String: Any]] = ordered.compactMap { observation in
        guard let candidate = observation.topCandidates(1).first else { return nil }
        let box = observation.boundingBox
        return ["text": candidate.string, "confidence": candidate.confidence,
                "box_normalized_bottom_left": [box.minX, box.minY, box.width, box.height]]
    }
    let payload: [String: Any] = ["backend": "apple-vision-offline", "lines": lines, "blocks": blocks]
    let encoded = try JSONSerialization.data(withJSONObject: payload, options: [.prettyPrinted])
    FileHandle.standardOutput.write(encoded)
    FileHandle.standardOutput.write(Data("\n".utf8))
} catch {
    fputs("offline OCR failed: \(error)\n", stderr)
    exit(1)
}
