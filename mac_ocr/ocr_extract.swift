#!/usr/bin/swift
import Foundation
import Vision
import AppKit
import PDFKit

struct Token: Codable {
    var text: String
    var x: Double
    var y: Double
    var w: Double
    var conf: Float
    var page: Int
    // 以下兩欄只在程式內部使用，不輸出：文字方向（0/90/180/270，影像座標
    // y 朝上量測）與沿文字方向的像素長度，供整頁方向校正。
    var dir: Int = 0
    var lenPx: Double = 0

    enum CodingKeys: String, CodingKey { case text, x, y, w, conf, page }
}

/// 掃描時紙張放反或橫放：Vision 仍讀得出字，但座標是影像座標，會讓整頁
/// 上下左右顛倒（例：姓名列跑到本俸列下方）。以整頁多數文字的方向判定，
/// 把座標轉回「文字正立」的版面；正立頁面完全不變動。
func normalizeOrientation(_ toks: [Token], imageW: Double, imageH: Double) -> [Token] {
    var bins = [0, 0, 0, 0]
    for t in toks where t.text.count >= 2 { bins[t.dir / 90] += 1 }
    let total = bins.reduce(0, +)
    guard total >= 5, let best = bins.indices.max(by: { bins[$0] < bins[$1] }),
          best != 0, Double(bins[best]) >= 0.6 * Double(total) else { return toks }
    return toks.map { t in
        var o = t
        switch best {
        case 2:   // 上下顛倒
            o.x = 1 - t.x; o.y = 1 - t.y
            o.w = t.lenPx / imageW
        case 1:   // 文字由下往上：頁面逆時針轉了 90°
            o.x = 1 - t.y; o.y = t.x
            o.w = t.lenPx / imageH
        default:  // 3：文字由上往下：頁面順時針轉了 90°
            o.x = t.y; o.y = 1 - t.x
            o.w = t.lenPx / imageH
        }
        return o
    }
}

/// 在整頁影像的一個矩形區塊上執行 Vision，座標換回整頁的 0–1 比例。
/// region 以整頁比例表示（左上為原點）。dropEdges 為 true 時，捨棄碰到「非頁緣」
/// 裁切邊界的文字，避免被切半的數字混入；完整版本會出現在相鄰區塊的重疊帶。
func tokens(from image: CGImage, page: Int,
            region: CGRect = CGRect(x: 0, y: 0, width: 1, height: 1),
            dropEdges: Bool = false) -> [Token] {
    let W = Double(image.width), H = Double(image.height)
    let rx = Double(region.minX), ry = Double(region.minY)
    let rw = Double(region.width), rh = Double(region.height)
    var source = image
    if rx > 0 || ry > 0 || rw < 1 || rh < 1 {
        guard let crop = image.cropping(to: CGRect(x: rx * W, y: ry * H,
                                                    width: rw * W, height: rh * H))
        else { return [] }
        source = crop
    }
    let edge = 0.004
    return recognize(source, page: page).compactMap { t in
        // t 的座標是區塊內比例；換回整頁比例。
        let left = t.x - t.w / 2, right = t.x + t.w / 2
        if dropEdges {
            if (rx > 0 && left < edge) || (rx + rw < 1 && right > 1 - edge) {
                return nil
            }
        }
        return Token(text: t.text, x: rx + t.x * rw, y: ry + t.y * rh,
                     w: t.w * rw, conf: t.conf, page: page,
                     dir: t.dir, lenPx: t.lenPx)
    }
}

func recognize(_ image: CGImage, page: Int) -> [Token] {
    var out: [Token] = []
    let req = VNRecognizeTextRequest()
    req.recognitionLevel = .accurate
    req.usesLanguageCorrection = true
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
            var dir = 0, lenPx = Double(box.width) * Double(image.width)
            if let bb = try? cand.boundingBox(for: r) {
                box = bb.boundingBox
                let dx = Double(bb.bottomRight.x - bb.bottomLeft.x) * Double(image.width)
                let dy = Double(bb.bottomRight.y - bb.bottomLeft.y) * Double(image.height)
                let ang = atan2(dy, dx) * 180 / Double.pi
                dir = (Int(((ang + 45 + 360).truncatingRemainder(dividingBy: 360)) / 90) % 4) * 90
                lenPx = (dx * dx + dy * dy).squareRoot()
            }
            out.append(Token(text: String(piece),
                             x: Double(box.midX),
                             y: Double(1.0 - box.midY),
                             w: Double(box.width),
                             conf: cand.confidence,
                             page: page, dir: dir, lenPx: lenPx))
        }
    }
    return out
}

let args = CommandLine.arguments
guard args.count > 1 else {
    FileHandle.standardError.write("用法：swift ocr_extract.swift 檔案.pdf [--scale 3]\n".data(using: .utf8)!)
    exit(1)
}
let pdfPath = args[1]
// --tiles CxR：把每頁切成 C 欄 × R 列、彼此重疊的區塊分別辨識。整頁一次辨識時，
// Vision 偶爾會整欄漏掉字級較大或底紋較重的金額（例：A4 橫式「每人兩行」清冊的
// 月支俸額／學術研究／應發合計欄）；分塊後每塊文字相對變大即可讀到。預設 1x1
// 與舊版完全相同，只有 worker 判斷需要時才會另外呼叫。
var tileCols = 1, tileRows = 1
if let ti = args.firstIndex(of: "--tiles"), ti + 1 < args.count {
    let parts = args[ti + 1].lowercased().split(separator: "x").compactMap { Int($0) }
    if parts.count == 2, (1...6).contains(parts[0]), (1...6).contains(parts[1]) {
        tileCols = parts[0]; tileRows = parts[1]
    }
}
var renderScale: CGFloat = 3.0
if let scaleIndex = args.firstIndex(of: "--scale"), scaleIndex + 1 < args.count,
   let value = Double(args[scaleIndex + 1]), value >= 2.0, value <= 6.0 {
    renderScale = CGFloat(value)
}
guard let doc = PDFDocument(url: URL(fileURLWithPath: pdfPath)) else {
    FileHandle.standardError.write("無法開啟 PDF\n".data(using: .utf8)!)
    exit(1)
}

var all: [Token] = []
for pi in 0..<doc.pageCount {
    guard let page = doc.page(at: pi) else { continue }
    let rect = page.bounds(for: .mediaBox)
    // 保留既有清冊驗證過的 3 倍預設；特殊 A3 雙行版面另由 worker 以
    // 4 倍補辨一次，利用兩次辨識信心互補，不影響其他學校的既有結果。
    let scale = renderScale
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

    var pageTokens: [Token] = []
    if tileCols == 1 && tileRows == 1 {
        pageTokens = tokens(from: cg, page: pi)
    } else {
        let overlap = 0.04
        for ci in 0..<tileCols {
            for ri in 0..<tileRows {
                let x0 = max(0, Double(ci) / Double(tileCols) - overlap)
                let x1 = min(1, Double(ci + 1) / Double(tileCols) + overlap)
                let y0 = max(0, Double(ri) / Double(tileRows) - overlap)
                let y1 = min(1, Double(ri + 1) / Double(tileRows) + overlap)
                pageTokens.append(contentsOf: tokens(
                    from: cg, page: pi,
                    region: CGRect(x: x0, y: y0, width: x1 - x0, height: y1 - y0),
                    dropEdges: true))
            }
        }
    }
    all.append(contentsOf: normalizeOrientation(
        pageTokens, imageW: Double(cg.width), imageH: Double(cg.height)))
}

let enc = JSONEncoder()
enc.outputFormatting = [.prettyPrinted, .sortedKeys]
if let data = try? enc.encode(all) {
    FileHandle.standardOutput.write(data)
}
