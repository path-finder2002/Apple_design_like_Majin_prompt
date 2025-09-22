import Foundation
import ImageIO
import CoreGraphics

let path = CommandLine.arguments[1]
let url = URL(fileURLWithPath: path)
guard let source = CGImageSourceCreateWithURL(url as CFURL, nil),
      let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
    fatalError("Failed to load image")
}
print("width", image.width, "height", image.height, "bitsPerPixel", image.bitsPerPixel, "bytesPerRow", image.bytesPerRow)
if let data = image.dataProvider?.data {
    let ptr = CFDataGetBytePtr(data)!
    print("firstPixel", ptr[0], ptr[1], ptr[2], ptr[3])
}
