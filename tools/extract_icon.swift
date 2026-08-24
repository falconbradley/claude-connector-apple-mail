// extract_icon.swift — render a macOS app's icon to a PNG at an exact size.
//
// Why not just `iconutil` the app's .icns? Mail's AppIcon.icns tops out at
// 256x256, and on macOS 26 the real icon lives in Assets.car as a layered
// (Liquid Glass) asset that no raster in the .icns represents. Asking
// NSWorkspace for the icon gets whatever the system itself draws, composited
// at whatever size we ask for.
import AppKit

let a = CommandLine.arguments
guard a.count == 4, let px = Int(a[3]) else {
    FileHandle.standardError.write("usage: extract_icon.swift <app-path> <out.png> <px>\n".data(using: .utf8)!)
    exit(2)
}
let (appPath, outPath) = (a[1], a[2])

let icon = NSWorkspace.shared.icon(forFile: appPath)
icon.size = NSSize(width: px, height: px)

guard let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: px, pixelsHigh: px,
                                bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true,
                                isPlanar: false, colorSpaceName: .deviceRGB,
                                bytesPerRow: 0, bitsPerPixel: 0) else {
    FileHandle.standardError.write("could not allocate bitmap\n".data(using: .utf8)!)
    exit(1)
}
rep.size = NSSize(width: px, height: px)

// NSImage.size is in points, so draw through an explicit bitmap context to get
// a PNG that is genuinely px-by-px rather than point-scaled.
NSGraphicsContext.saveGraphicsState()
NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
NSGraphicsContext.current?.imageInterpolation = .high
icon.draw(in: NSRect(x: 0, y: 0, width: px, height: px),
          from: .zero, operation: .sourceOver, fraction: 1.0)
NSGraphicsContext.restoreGraphicsState()

guard let data = rep.representation(using: .png, properties: [:]) else {
    FileHandle.standardError.write("PNG encode failed\n".data(using: .utf8)!)
    exit(1)
}
try data.write(to: URL(fileURLWithPath: outPath))
