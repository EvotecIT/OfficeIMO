// swift-tools-version: 6.0
import Foundation
import PackageDescription

// The engine is built from this checkout, so the apps always use the engine source they were built with:
// run Build/Studio/Apple/Build-StudioEngine.ps1 before building them.
let engine = "OfficeIMOStudioEngine.xcframework"
let packageDirectory = URL(fileURLWithPath: #filePath).deletingLastPathComponent()
precondition(FileManager.default.fileExists(atPath: packageDirectory.appendingPathComponent(engine).path),
             "The Studio engine hasn't been built. Run Build/Studio/Apple/Build-StudioEngine.ps1 from the repository root.")

let package = Package(
    name: "OfficeIMOStudioEngine",
    platforms: [.macOS("26.0"), .iOS("26.0")],
    products: [.library(name: "OfficeIMOStudioEngine", targets: ["OfficeIMOStudioEngine", "OfficeIMOStudioEngineNotices"])],
    targets: [
        .binaryTarget(name: "OfficeIMOStudioEngine", path: engine),
        .target(name: "OfficeIMOStudioEngineNotices", resources: [.copy("Licenses")])
    ]
)
