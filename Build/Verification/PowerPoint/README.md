# PowerPoint Desktop development verification

`PowerPointDesktopReferenceRenderer.cs` is linked into `OfficeIMO.PowerPoint.Tests` and the non-packable `OfficeIMO.ChartForgeX.Examples` executable. It is not compiled into a shipped OfficeIMO assembly or package. The helper uses public document and Open XML APIs; no product friend access is required.

On Windows with PowerPoint Desktop installed, the existing tests opt in through `OFFICEIMO_POWERPOINT_DESKTOP_REFERENCE=1` or `OFFICEIMO_RUN_POWERPOINT_DESKTOP_CORPUS=1`. The legacy PPT corpus uses `OFFICEIMO_RUN_LEGACY_PPT_COM_VALIDATION=1`. Ordinary tests do not launch Office. The ChartForgeX example opts in with `--native-powerpoint`; see its README for the command.

The shared verifier disables macros before opening a document, validates the exported image sequence and PNG content, and retains on-slide geometry and omitted-content checks. It reports unavailable or failed native rendering without substituting managed images. Native Office execution is separate from managed rendering proof.

Product callers use `PowerPointSlide.ExportImage(...)` or `ToImage()` for managed output. Record independently produced native artifacts with `PowerPointCompatibilityReport.RecordExternal(...)`.
