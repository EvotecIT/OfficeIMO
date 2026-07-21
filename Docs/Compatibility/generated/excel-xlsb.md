# OfficeIMO.Excel.Xlsb capability contract

Schema version: 1

| Category | Capability | Format | Representation | Legacy import | New legacy | Legacy round-trip | Modern to legacy | Legacy to modern | Fidelity | Note |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Lifecycle | Excel.Xlsb.Lifecycle.File | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Structure | Excel.Xlsb.Structure.Worksheets | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Cells | Excel.Xlsb.Cells.Values | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Formulas | Excel.Xlsb.Formulas.Tokens | Excel.Xlsb | Native | Native | Approximated | PreservedOpaque | Approximated | Native | Editability, Carrier | Existing unsupported formula payloads are retained during allowed cell-value rewrites; new generation supports a bounded token subset. |
| Formatting | Excel.Xlsb.Formatting.Styles | Excel.Xlsb | Native | Native | Approximated | PreservedOpaque | Approximated | Native | Editability, Carrier | A useful style subset projects and can be generated. Complex gradients, extensions, differential styles, and custom style families remain guarded. |
| Structure | Excel.Xlsb.Structure.Geometry | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Navigation | Excel.Xlsb.Navigation.Hyperlinks | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Data | Excel.Xlsb.Data.AutoFilter | Excel.Xlsb | Native | Native | Approximated | PreservedOpaque | Approximated | Native | Editability, Carrier | Unsupported criteria remain preserved on import; native new-package generation supports the documented equality-list subset. |
| Print | Excel.Xlsb.Print.PageSetup | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Security | Excel.Xlsb.Security.Protection | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Names | Excel.Xlsb.Names.DefinedNames | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Calculation | Excel.Xlsb.Calculation.Settings | Excel.Xlsb | Native | Native | Native | PreservedOpaque | Native | Native | None |  |
| Drawing | Excel.Xlsb.Drawing.ImagesShapes | Excel.Xlsb | Approximation | PreservedOpaque | Blocked | PreservedOpaque | Rasterized | EmbeddedSource | Visual, Editability, Carrier | Related package parts survive exact copy. Unsupported generated output either blocks or is rendered into a palette-quantized cell raster; PreservationOnly can also retain the source carrier. |
| Drawing | Excel.Xlsb.Drawing.Charts | Excel.Xlsb | Approximation | PreservedOpaque | Blocked | PreservedOpaque | Rasterized | EmbeddedSource | Visual, Editability, Carrier | Chart package parts survive exact copy. Unsupported cross-format output either blocks or uses the omission-gated worksheet cell-raster fallback. |
| Analytics | Excel.Xlsb.Analytics.Pivots | Excel.Xlsb | Approximation | PreservedOpaque | Blocked | PreservedOpaque | Rasterized | EmbeddedSource | Visual, Editability, Carrier | Existing package parts are preservation-owned. Unsupported cross-format output either blocks or uses the omission-gated worksheet cell-raster fallback. |
| Embedded | Excel.Xlsb.Embedded.Vba | Excel.Xlsb | Opaque | PreservedOpaque | Blocked | PreservedOpaque | Blocked | Dropped | Behavioral, Carrier, Security | Exact package copy and preservation-aware rewrites retain the carrier; conversion never treats unprojected BIFF12 records as editable parity. |
| Embedded | Excel.Xlsb.Embedded.OleActiveX | Excel.Xlsb | Opaque | PreservedOpaque | Blocked | PreservedOpaque | Blocked | Dropped | Behavioral, Carrier, Security | Exact package copy and preservation-aware rewrites retain the carrier; conversion never treats unprojected BIFF12 records as editable parity. |
| Security | Excel.Xlsb.Security.DigitalSignatures | Excel.Xlsb | Opaque | PreservedOpaque | Blocked | PreservedOpaque | Blocked | Dropped | Carrier, Security | Exact package copy and preservation-aware rewrites retain the carrier; conversion never treats unprojected BIFF12 records as editable parity. |
| Preservation | Excel.Xlsb.Preservation.UnknownRecords | Excel.Xlsb | Opaque | PreservedOpaque | Blocked | PreservedOpaque | Blocked | Dropped | Semantic, Editability, Carrier | Exact package copy and preservation-aware rewrites retain the carrier; conversion never treats unprojected BIFF12 records as editable parity. |
| Preservation | Excel.Xlsb.Preservation.SourceCarrier | Excel.Xlsb | Opaque | PreservedOpaque | NotApplicable | PreservedOpaque | EmbeddedSource | EmbeddedSource | Editability, Carrier, Security | Embedding is explicit because original bytes may contain macros, hidden sheets, or embedded payloads. PreservationOnly enables it automatically; callers can recover the verified payload through TryGetCompatibilitySourcePayload. |
