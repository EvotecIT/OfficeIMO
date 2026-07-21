using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

/// <summary>Versioned capability contracts for BIFF8 XLS and BIFF12 XLSB interoperability.</summary>
public static class ExcelCompatibilityCatalog {
    private const string XlsFormatId = "Excel.Xls";
    private const string XlsbFormatId = "Excel.Xlsb";

    /// <summary>Gets the current machine-readable BIFF8 XLS contract.</summary>
    public static OfficeCapabilityCatalog Xls { get; } = new(
        "OfficeIMO.Excel.LegacyXls",
        schemaVersion: 1,
        new[] {
            XlsNative("Lifecycle.File", "Lifecycle", "Path, stream, byte, synchronous, and asynchronous workbook flows."),
            XlsSubset("Lifecycle.Variants", "Lifecycle", "XLS, XLT, XLA, XLM, and XLW legacy extension classification.",
                "All legacy variants are detected and reported. Native output is intentionally limited to XLS until each variant has a proven writer contract."),
            XlsNative("Structure.Worksheets", "Structure", "Worksheets, order, names, visibility, selected tab, and active sheet."),
            XlsSubset("Structure.LegacySheetKinds", "Structure", "Chart, macro, dialog, and other legacy sheet kinds.",
                "Chart sheets project to chart-sheet parts. Macro, dialog, and unsupported sheet kinds remain explicit conversion findings."),
            XlsNative("Cells.Values", "Cells", "Numeric, text, Boolean, error, blank, date, and cached formula values."),
            XlsSubset("Cells.RichText", "Cells", "Cell rich-text runs and phonetic information.",
                "Common rich text is editable; unsupported phonetic and formatting variants are preserved or blocked."),
            XlsSubset("Formulas.Tokens", "Formulas", "BIFF8 formula tokens, shared formulas, arrays, names, and cached results.",
                "A broad tested token subset is native. Unsupported functions, token forms, and multi-cell array cases are diagnosed before output."),
            XlsSubset("Formatting.Styles", "Formatting", "Fonts, number formats, fills, borders, alignment, protection, and XF inheritance.",
                "Common BIFF8 style records are editable; non-representable Open XML styling blocks legacy output."),
            XlsNative("Structure.Geometry", "Structure", "Row and column geometry, hidden state, outlines, merges, panes, views, and selections."),
            XlsNative("Navigation.Hyperlinks", "Navigation", "Internal and external hyperlinks."),
            XlsNative("Review.Comments", "Review", "Classic cell comments and authors."),
            XlsSubset("Data.AutoFilterSort", "Data", "AutoFilter, sort state, and supported filter criteria.",
                "Supported BIFF8 filter forms are native; modern-only filters are preflight-blocked."),
            XlsSubset("Data.Validation", "Data", "Data-validation rules, prompts, and error messages.",
                "Supported rule and formula subsets are native."),
            XlsSubset("Data.ConditionalFormatting", "Data", "Classic conditional-formatting rules and differential formatting.",
                "Classic rules are native; modern extension rules are approximated or blocked."),
            XlsNative("Names.DefinedNames", "Names", "Workbook and worksheet defined names and supported external-sheet references."),
            XlsSubset("Print.PageSetup", "Print", "Page setup, margins, breaks, print titles and areas, and headers and footers.",
                "Supported printer settings and header/footer text are native; unsupported relationships are blocked."),
            XlsNative("Security.Protection", "Security", "Workbook, worksheet, range, sharing, and classic password-hash protection metadata."),
            XlsSubset("Metadata.Properties", "Metadata", "Built-in and custom compound-file document properties.",
                "Supported scalar properties are native; unknown variants remain preservation-owned."),
            XlsVisualFallback("Drawing.ImagesShapes", "Drawing", "Worksheet images, shapes, text boxes, groups, and controls.",
                OfficeCapabilityCoverageState.Native,
                "Supported legacy images project editably. Unsupported modern drawing content either blocks or is rendered into a palette-quantized cell raster; PreservationOnly can also retain the source carrier."),
            XlsVisualFallback("Drawing.Charts", "Drawing", "Embedded charts and chart sheets.",
                OfficeCapabilityCoverageState.Approximated,
                "Chart-sheet structure and supported chart visuals project through the editable model. Unsupported modern chart output either blocks or is rendered into a palette-quantized cell raster."),
            XlsVisualFallback("Analytics.Sparklines", "Analytics", "Modern sparklines and their rendered appearance.",
                OfficeCapabilityCoverageState.Dropped,
                "BIFF8 has no native sparkline model. Visual modes render the complete worksheet into a palette-quantized cell raster; strict modes block."),
            XlsVisualFallback("Analytics.Pivots", "Analytics", "Pivot tables, pivot caches, slicers, and timelines.",
                OfficeCapabilityCoverageState.Dropped,
                "Existing binary metadata is detected or preserved. Unsupported modern output either blocks or uses the omission-gated worksheet cell-raster fallback."),
            XlsOpaque("Embedded.Vba", "Embedded", "VBA project storage and macro-sheet behavior.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            XlsOpaque("Embedded.OleActiveX", "Embedded", "OLE packages, ActiveX controls, and cached previews.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            XlsOpaque("Security.DigitalSignatures", "Security", "Legacy digital-signature streams and storages.",
                OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            new OfficeCapability(
                "Excel.Security.Encryption", XlsFormatId, OfficeDocumentFamily.Excel, "Security",
                "BIFF8 XOR, classic RC4, and RC4 CryptoAPI password-to-open encryption.",
                OfficeCapabilityRepresentability.Native,
                OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.NotImplemented,
                OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
                OfficeCapabilityCoverageState.Dropped,
                OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier,
                "Supported encrypted XLS input can be loaded with a password. Cross-generation conversion reports that password protection is removed and blocks by default; first-party encrypted XLS authoring remains incomplete."),
            XlsOpaque("Preservation.UnknownRecords", "Preservation", "Unknown BIFF records, compound streams, and storages.",
                OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability),
            SourceCarrier(XlsFormatId, "Excel.Xls.Preservation.SourceCarrier")
        });

    /// <summary>Gets the current machine-readable BIFF12 XLSB contract.</summary>
    public static OfficeCapabilityCatalog Xlsb { get; } = new(
        "OfficeIMO.Excel.Xlsb",
        schemaVersion: 1,
        new[] {
            XlsbNative("Lifecycle.File", "Lifecycle", "Path, stream, byte, new-package, exact-copy, and preservation-aware rewrite flows."),
            XlsbNative("Structure.Worksheets", "Structure", "Worksheet order, names, visibility, workbook views, and date system."),
            XlsbNative("Cells.Values", "Cells", "Numeric, text, Boolean, error, blank, date, shared-string, and cached formula values."),
            XlsbSubset("Formulas.Tokens", "Formulas", "BIFF12 formulas, cached values, and supported defined-name tokens.",
                "Existing unsupported formula payloads are retained during allowed cell-value rewrites; new generation supports a bounded token subset."),
            XlsbSubset("Formatting.Styles", "Formatting", "Fonts, number formats, fills, borders, alignment, protection, and cell formats.",
                "A useful style subset projects and can be generated. Complex gradients, extensions, differential styles, and custom style families remain guarded."),
            XlsbNative("Structure.Geometry", "Structure", "Dimensions, rows, columns, merges, panes, views, selections, and outline metadata."),
            XlsbNative("Navigation.Hyperlinks", "Navigation", "Internal and relationship-backed external hyperlinks."),
            XlsbSubset("Data.AutoFilter", "Data", "Worksheet AutoFilter ranges and equality-list criteria.",
                "Unsupported criteria remain preserved on import; native new-package generation supports the documented equality-list subset."),
            XlsbNative("Print.PageSetup", "Print", "Print options, margins, page setup, and text headers and footers."),
            XlsbNative("Security.Protection", "Security", "Workbook and worksheet classic protection metadata."),
            XlsbNative("Names.DefinedNames", "Names", "Supported workbook defined names and self-sheet references."),
            XlsbNative("Calculation.Settings", "Calculation", "Workbook calculation mode, iteration, and related supported properties."),
            XlsbVisualFallback("Drawing.ImagesShapes", "Drawing", "Images, shapes, controls, and drawing relationships.",
                "Related package parts survive exact copy. Unsupported generated output either blocks or is rendered into a palette-quantized cell raster; PreservationOnly can also retain the source carrier."),
            XlsbVisualFallback("Drawing.Charts", "Drawing", "Charts and chart-sheet relationships.",
                "Chart package parts survive exact copy. Unsupported cross-format output either blocks or uses the omission-gated worksheet cell-raster fallback."),
            XlsbVisualFallback("Analytics.Pivots", "Analytics", "Pivot tables, caches, slicers, timelines, and connections.",
                "Existing package parts are preservation-owned. Unsupported cross-format output either blocks or uses the omission-gated worksheet cell-raster fallback."),
            XlsbOpaque("Embedded.Vba", "Embedded", "VBA project and related macro carrier parts.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            XlsbOpaque("Embedded.OleActiveX", "Embedded", "OLE packages, ActiveX controls, and relationship-backed embedded content.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            XlsbOpaque("Security.DigitalSignatures", "Security", "Open Packaging Conventions digital signatures.",
                OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            XlsbOpaque("Preservation.UnknownRecords", "Preservation", "Unknown BIFF12 records and package parts.",
                OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability),
            SourceCarrier(XlsbFormatId, "Excel.Xlsb.Preservation.SourceCarrier")
        });

    /// <summary>Gets both Excel binary-format contracts in stable order.</summary>
    public static IReadOnlyList<OfficeCapabilityCatalog> All { get; } = Array.AsReadOnly(new[] { Xls, Xlsb });

    private static OfficeCapability XlsNative(string id, string category, string description) => Capability(
        XlsFormatId, "Excel.Xls." + id, category, description, OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.Native);

    private static OfficeCapability XlsSubset(string id, string category, string description, string note) => Capability(
        XlsFormatId, "Excel.Xls." + id, category, description, OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Approximated,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Approximated,
        OfficeCapabilityCoverageState.Native,
        OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability, note);

    private static OfficeCapability XlsVisualFallback(
        string id,
        string category,
        string description,
        OfficeCapabilityCoverageState legacyToModern,
        string note) => Capability(
        XlsFormatId, "Excel.Xls." + id, category, description, OfficeCapabilityRepresentability.Approximation,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Rasterized,
        legacyToModern,
        OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Carrier, note);

    private static OfficeCapability XlsOpaque(string id, string category, string description, OfficeCompatibilityImpact impact) => Capability(
        XlsFormatId, "Excel.Xls." + id, category, description, OfficeCapabilityRepresentability.Opaque,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.Dropped, impact,
        "Exact carrier retention is safe for unchanged or explicitly preservation-aware XLS flows; editable cross-format parity is not claimed.");

    private static OfficeCapability XlsbNative(string id, string category, string description) => Capability(
        XlsbFormatId, "Excel.Xlsb." + id, category, description, OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.Native);

    private static OfficeCapability XlsbSubset(string id, string category, string description, string note) => Capability(
        XlsbFormatId, "Excel.Xlsb." + id, category, description, OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Approximated,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Approximated,
        OfficeCapabilityCoverageState.Native,
        OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Carrier, note);

    private static OfficeCapability XlsbVisualFallback(string id, string category, string description, string note) => Capability(
        XlsbFormatId, "Excel.Xlsb." + id, category, description, OfficeCapabilityRepresentability.Approximation,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Rasterized,
        OfficeCapabilityCoverageState.EmbeddedSource,
        OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Carrier, note);

    private static OfficeCapability XlsbOpaque(string id, string category, string description, OfficeCompatibilityImpact impact) => Capability(
        XlsbFormatId, "Excel.Xlsb." + id, category, description, OfficeCapabilityRepresentability.Opaque,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.Dropped, impact,
        "Exact package copy and preservation-aware rewrites retain the carrier; conversion never treats unprojected BIFF12 records as editable parity.");

    private static OfficeCapability SourceCarrier(string formatId, string id) => Capability(
        formatId,
        id,
        "Preservation",
        "Hash-verified original source retained beside an editable projection or visual compatibility fallback.",
        OfficeCapabilityRepresentability.Opaque,
        OfficeCapabilityCoverageState.PreservedOpaque,
        OfficeCapabilityCoverageState.NotApplicable,
        OfficeCapabilityCoverageState.PreservedOpaque,
        OfficeCapabilityCoverageState.EmbeddedSource,
        OfficeCapabilityCoverageState.EmbeddedSource,
        OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Security,
        "Embedding is explicit because original bytes may contain macros, hidden sheets, or embedded payloads. PreservationOnly enables it automatically; callers can recover the verified payload through TryGetCompatibilitySourcePayload.");

    private static OfficeCapability Capability(
        string formatId,
        string id,
        string category,
        string description,
        OfficeCapabilityRepresentability representability,
        OfficeCapabilityCoverageState legacyImport,
        OfficeCapabilityCoverageState newLegacyWrite,
        OfficeCapabilityCoverageState legacyRoundTrip,
        OfficeCapabilityCoverageState modernToLegacy,
        OfficeCapabilityCoverageState legacyToModern,
        OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.None,
        string? note = null) => new(
        id, formatId, OfficeDocumentFamily.Excel, category, description, representability,
        legacyImport, newLegacyWrite, legacyRoundTrip, modernToLegacy, legacyToModern, impact, note);
}
