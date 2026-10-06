# Standalone CLI archives

Use the shared PowerForge publisher to build self-contained CLI archives without requiring users to install .NET:

```sh
powerforge dotnet publish --config Build/Tool/powerforge.dotnetpublish.json --validate --rid osx-arm64
powerforge dotnet publish --config Build/Tool/powerforge.dotnetpublish.json --rid osx-arm64
```

Select the target runtime explicitly. The configuration defines Windows, Linux and macOS x64/Arm64 archive candidates under `Artifacts/Tool`. It uses the existing CLI project and command contract. Test the extracted executable on each target platform before distributing it; a successful cross-publish does not establish native runtime compatibility. This configuration does not sign, notarize or publish a release. NativeAOT remains a separately qualified build style through the existing AOT lane.

For provenance verification, users supply their own trusted c2patool executable and trust material. Ordinary structural inspection and Unicode assessment need neither. See the [provenance support matrix](../../Docs/officeimo.provenance-support-matrix.md#configure-a-verifier) for setup and limitations. OfficeIMO source is MIT-licensed; preserve the licenses of packaged dependencies when distributing archives.
