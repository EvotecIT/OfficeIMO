// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "OfficeIMOStudioEngine",
    platforms: [.macOS("26.0"), .iOS("26.0")],
    products: [.library(name: "OfficeIMOStudioEngine", targets: ["OfficeIMOStudioEngine", "OfficeIMOStudioEngineNotices"])],
    targets: [
        .binaryTarget(name: "OfficeIMOStudioEngine",
                      url: "https://github.com/EvotecIT/OfficeIMO/releases/download/studio-engine-v1.0.0/OfficeIMOStudioEngine-1.0.0.zip",
                      checksum: "97e882243fffcdc04074552ea22cdfa249a8691e7c6b1e9b9ea9b6ff6070cbea"),
        .target(name: "OfficeIMOStudioEngineNotices", resources: [.copy("Licenses")])
    ]
)
