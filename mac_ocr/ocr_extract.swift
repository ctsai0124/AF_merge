#!/usr/bin/swift
import Foundation
import Vision
import AppKit
import PDFKit

struct Token: Codable {
    let text: String
    let x: Double
    let y: Double
    let w: Double
    let conf: Float
    let page: Int
}

func tokens(from image: CGImage, page: Int) -> [Token] {
    var out: [Token] = []
    let req = VNRecognizeTextRequest()
    req.recognitionLevel = .accurate
    req.usesLanguageCorrection = false
    req.recognitionLanguages = ["zh-Hant", "en-US"]

    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    try? handler.perform([req])

    for obs in (req.results ?? []) {
        guard let cand = obs.topCandidates(1).first else { continue }
        let full = cand.string

        var searchStart = full.startIndex
        for piece in full.split(separator: " ") {
            guard let r = full.range(of: String(piece), range: searchStart..<full.endIndex)
            else { continue }
            searchStart = r.upperBound

            var box = obs.boundingBox
            if let bb = try? cand.boundingBox(for: r) {
                box = bb.boundingBox
            }
            out.append(Token(text: String(piece),
                             x: Double(box.midX),
                             y: Double(1.0 - box.midY),
                             w: Double(box.width),
                             conf: cand.confidence,
                             page: page))
        }
    }
    return out
}

let args = CommandLine.arguments
guard args.count > 1 else {
    FileHandle.standardError.write("用法：swift ocr_extract.swift 檔案.pdf\n".data(using: .utf8)!)
    exit(1)
}
guard let doc = PDFDocument(url: URL(fileURLWithPath: args[1])) else {
    FileHandle.standardError.write("無法開啟 PDF\n".data(using: .utf8)!)
    exit(1)
}

var all: [Token] = []
for pi in 0..<doc.pageCount {
    guard let page = doc.page(at: pi) else { continue }
    let rect = page.bounds(for: .mediaBox)
    let scale: CGFloat = 3.0
    // PDFPage.draw 會套用頁面的 /Rotate，但 bounds(for:) 回傳的是尚未旋轉的
    // MediaBox。90°／270°頁面若仍建立橫向畫布，旋轉後的直向內容會被裁掉。
    let rotation = ((page.rotation % 360) + 360) % 360
    let quarterTurn = rotation == 90 || rotation == 270
    let displayWidth = quarterTurn ? rect.height : rect.width
    let displayHeight = quarterTurn ? rect.width : rect.height
    let size = NSSize(width: displayWidth * scale, height: displayHeight * scale)

    let img = NSImage(size: size)
    img.lockFocus()
    NSColor.white.setFill()
    NSRect(origin: .zero, size: size).fill()
    if let ctx = NSGraphicsContext.current?.cgContext {
        ctx.scaleBy(x: scale, y: scale)
        page.draw(with: .mediaBox, to: ctx)
    }
    img.unlockFocus()

    guard let tiff = img.tiffRepresentation,
          let rep = NSBitmapImageRep(data: tiff),
          let cg = rep.cgImage else { continue }

    all.append(contentsOf: tokens(from: cg, page: pi))
}

let enc = JSONEncoder()
enc.outputFormatting = [.prettyPrinted, .sortedKeys]
if let data = try? enc.encode(all) {
    FileHandle.standardOutput.write(data)
}
