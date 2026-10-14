// swift-tools-version: 6.0
import PackageDescription

// These Foundation-only sources are also compiled directly into the native apps.
let package = Package(
    name: "OfficeIMOStudioStateContracts",
    platforms: [.macOS(.v13)],
    targets: [
        .target(name: "OfficeIMOStudioState", path: "OfficeIMOStudio/Shared/State"),
        .testTarget(name: "OfficeIMOStudioStateTests", dependencies: ["OfficeIMOStudioState"],
                    path: "OfficeIMOStudioStateTests")
    ]
)
