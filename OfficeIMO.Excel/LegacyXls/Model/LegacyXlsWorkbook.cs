using OfficeIMO.Excel.LegacyXls.Compound;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Projection;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Neutral workbook model produced from a legacy BIFF `.xls` stream.
    /// </summary>
    public sealed class LegacyXlsWorkbook {
        private readonly List<LegacyXlsWorksheet> _worksheets = new();
        private readonly List<LegacyXlsNumberFormat> _numberFormats = new();
        private readonly List<LegacyXlsFont> _fonts = new();
        private readonly List<string> _paletteColors = new();
        private readonly List<LegacyXlsCellFormat> _cellFormats = new();
        private readonly List<LegacyXlsCellStyle> _cellStyles = new();
        private readonly List<LegacyXlsCellStyleExtension> _cellStyleExtensions = new();
        private readonly List<LegacyXlsDefinedName> _definedNames = new();
        private readonly List<LegacyXlsExternalReference> _externalReferences = new();
        private readonly List<LegacyXlsExternalQueryConnection> _externalQueryConnections = new();
        private readonly List<LegacyXlsDataConsolidationReference> _dataConsolidationReferences = new();
        private readonly List<LegacyXlsDataConsolidationName> _dataConsolidationNames = new();
        private readonly List<LegacyXlsPivotTableRecord> _pivotTableRecords = new();
        private readonly List<LegacyXlsChartRecord> _chartRecords = new();
        private readonly List<LegacyXlsChartSheet> _chartSheets = new();
        private readonly List<LegacyXlsDrawingRecord> _drawingRecords = new();
        private readonly List<LegacyXlsThemeRecord> _themeRecords = new();
        private readonly List<LegacyXlsDifferentialFormat> _differentialFormats = new();
        private readonly List<LegacyXlsTableStyleCollection> _tableStyleCollections = new();
        private readonly List<LegacyXlsTableStyle> _tableStyles = new();
        private readonly List<LegacyXlsCompoundFeatureRecord> _compoundFeatureRecords = new();
        private readonly List<LegacyXlsUnsupportedSheet> _unsupportedSheets = new();
        private readonly List<LegacyXlsUnsupportedFeature> _unsupportedFeatures = new();
        private readonly List<LegacyXlsPreservedFeatureRecord> _preservedFeatureRecords = new();
        private readonly List<LegacyXlsWorkbookMetadataRecord> _metadataRecords = new();
        private readonly List<LegacyXlsWorkbookFutureMetadataRecord> _futureMetadataRecords = new();
        private readonly List<LegacyXlsFormulaTokenRecord> _formulaTokenRecords = new();
        private readonly List<LegacyXlsFutureFunctionAlias> _futureFunctionAliases = new();
        private readonly List<LegacyXlsWorkbookWindow> _windows = new();
        private readonly List<LegacyXlsImportDiagnostic> _diagnostics = new();
        private readonly LegacyXlsCalculationSettings _calculationSettings = new();

        /// <summary>Gets the private source container retained for same-format compound rewriting.</summary>
        internal OfficeCompoundFile? SourceCompoundFile { get; private set; }

        /// <summary>Gets whether the workbook stream was successfully decrypted from password-to-open protection.</summary>
        public bool WasEncryptedSource { get; internal set; }

        internal LegacyXlsWorkbook() {
        }

        /// <summary>
        /// Gets worksheets parsed from the legacy workbook stream.
        /// </summary>
        public IReadOnlyList<LegacyXlsWorksheet> Worksheets => _worksheets;

        /// <summary>
        /// Gets custom number formats parsed from FORMAT records.
        /// </summary>
        public IReadOnlyList<LegacyXlsNumberFormat> NumberFormats => _numberFormats;

        /// <summary>
        /// Gets fonts parsed from Font records.
        /// </summary>
        public IReadOnlyList<LegacyXlsFont> Fonts => _fonts;

        /// <summary>
        /// Gets custom palette colors parsed from the Palette record as ARGB hex values.
        /// </summary>
        public IReadOnlyList<string> PaletteColors => _paletteColors;

        /// <summary>
        /// Gets cell formats parsed from XF records.
        /// </summary>
        public IReadOnlyList<LegacyXlsCellFormat> CellFormats => _cellFormats;

        /// <summary>
        /// Gets workbook cell styles parsed from Style records.
        /// </summary>
        public IReadOnlyList<LegacyXlsCellStyle> CellStyles => _cellStyles;

        /// <summary>
        /// Gets preserve-only cell style extension records parsed from XFExt records.
        /// </summary>
        public IReadOnlyList<LegacyXlsCellStyleExtension> CellStyleExtensions => _cellStyleExtensions;

        /// <summary>
        /// Gets defined names parsed from Lbl records.
        /// </summary>
        public IReadOnlyList<LegacyXlsDefinedName> DefinedNames => _definedNames;

        /// <summary>
        /// Gets supporting links discovered from SupBook records.
        /// </summary>
        public IReadOnlyList<LegacyXlsExternalReference> ExternalReferences => _externalReferences;

        /// <summary>
        /// Gets preserve-only DBQueryExt query connection metadata discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsExternalQueryConnection> ExternalQueryConnections => _externalQueryConnections;

        /// <summary>
        /// Gets preserve-only DConRef source ranges discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsDataConsolidationReference> DataConsolidationReferences => _dataConsolidationReferences;

        /// <summary>
        /// Gets DConName named consolidation sources discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsDataConsolidationName> DataConsolidationNames => _dataConsolidationNames;

        /// <summary>
        /// Gets preserve-only PivotTable BIFF records discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsPivotTableRecord> PivotTableRecords => _pivotTableRecords;

        /// <summary>
        /// Gets preserve-only chart BIFF records discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsChartRecord> ChartRecords => _chartRecords;

        /// <summary>
        /// Gets legacy chart sheets decoded from chart-sheet substreams.
        /// </summary>
        public IReadOnlyList<LegacyXlsChartSheet> ChartSheets => _chartSheets;

        /// <summary>
        /// Gets preserve-only drawing and object BIFF records discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsDrawingRecord> DrawingRecords => _drawingRecords;

        /// <summary>
        /// Gets workbook Theme records discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsThemeRecord> ThemeRecords => _themeRecords;

        /// <summary>
        /// Gets parsed differential formats used by conditional formatting extensions.
        /// </summary>
        public IReadOnlyList<LegacyXlsDifferentialFormat> DifferentialFormats => _differentialFormats;

        /// <summary>
        /// Gets preserve-only workbook table style collection records discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsTableStyleCollection> TableStyleCollections => _tableStyleCollections;

        /// <summary>
        /// Gets preserve-only user-defined table styles discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsTableStyle> TableStyles => _tableStyles;

        /// <summary>
        /// Gets preserve-only compound container features discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsCompoundFeatureRecord> CompoundFeatureRecords => _compoundFeatureRecords;

        /// <summary>
        /// Gets calculation settings parsed from BIFF calculation records.
        /// </summary>
        public LegacyXlsCalculationSettings CalculationSettings => _calculationSettings;

        /// <summary>
        /// Gets legacy sheet entries discovered but not imported as worksheets.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedSheet> UnsupportedSheets => _unsupportedSheets;

        /// <summary>
        /// Gets unsupported or preserve-only features discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedFeature> UnsupportedFeatures => _unsupportedFeatures;

        /// <summary>
        /// Gets preserve-only BIFF feature records discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsPreservedFeatureRecord> PreservedFeatureRecords => _preservedFeatureRecords;

        internal LegacyXlsDocumentProperties? DocumentProperties { get; private set; }

        /// <summary>
        /// Gets workbook-level BIFF metadata records decoded during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsWorkbookMetadataRecord> MetadataRecords => _metadataRecords;

        /// <summary>
        /// Gets preserve-only extended workbook metadata records decoded during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsWorkbookFutureMetadataRecord> FutureMetadataRecords => _futureMetadataRecords;

        /// <summary>
        /// Gets BIFF parsed-formula token observations captured during import for corpus diagnostics.
        /// </summary>
        public IReadOnlyList<LegacyXlsFormulaTokenRecord> FormulaTokenRecords => _formulaTokenRecords;

        /// <summary>
        /// Gets Excel future-function aliases discovered from BIFF defined-name records.
        /// </summary>
        public IReadOnlyList<LegacyXlsFutureFunctionAlias> FutureFunctionAliases => _futureFunctionAliases;

        /// <summary>
        /// Gets workbook sheet tab identifiers decoded from a TabId record.
        /// </summary>
        public LegacyXlsSheetTabIdCollection? SheetTabIds { get; private set; }

        /// <summary>
        /// Gets workbook windows decoded from Window1 records.
        /// </summary>
        public IReadOnlyList<LegacyXlsWorkbookWindow> Windows => _windows;

        /// <summary>
        /// Gets diagnostics produced while reading the legacy workbook.
        /// </summary>
        public IReadOnlyList<LegacyXlsImportDiagnostic> Diagnostics => _diagnostics;

        /// <summary>
        /// Gets whether the workbook uses the Excel 1904 date system.
        /// </summary>
        public bool Uses1904DateSystem { get; private set; }

        /// <summary>
        /// Gets the workbook text code page decoded from a CodePage record.
        /// </summary>
        public ushort? CodePage { get; private set; }

        /// <summary>
        /// Gets the workbook object name used by the VBA project, when specified.
        /// </summary>
        public string? CodeName { get; private set; }

        /// <summary>
        /// Gets the user interface code page decoded from an InterfaceHdr record.
        /// </summary>
        public ushort? UserInterfaceCodePage { get; private set; }

        /// <summary>
        /// Gets country and region metadata decoded from a Country record.
        /// </summary>
        public LegacyXlsCountryInfo? Country { get; private set; }

        /// <summary>
        /// Gets whether the workbook requested saving a backup copy.
        /// </summary>
        public bool? SaveBackup { get; private set; }

        /// <summary>
        /// Gets whether external-link values should not be saved with the workbook.
        /// </summary>
        public bool? DoNotSaveExternalLinkValues { get; private set; }

        /// <summary>
        /// Gets whether the workbook has an envelope.
        /// </summary>
        public bool? HasEnvelope { get; private set; }

        /// <summary>
        /// Gets whether the workbook envelope is visible.
        /// </summary>
        public bool? EnvelopeVisible { get; private set; }

        /// <summary>
        /// Gets whether the workbook envelope was initialized.
        /// </summary>
        public bool? EnvelopeInitialized { get; private set; }

        /// <summary>
        /// Gets the raw external-link update mode decoded from BookBool flags.
        /// </summary>
        public byte? ExternalLinkUpdateMode { get; private set; }

        /// <summary>
        /// Gets whether borders are hidden for inactive tables.
        /// </summary>
        public bool? HideBordersForInactiveTables { get; private set; }

        /// <summary>
        /// Gets the number of built-in function categories decoded from a BuiltInFnGroupCount record.
        /// </summary>
        public ushort? BuiltInFunctionGroupCount { get; private set; }

        /// <summary>
        /// Gets the raw hidden-object display mode decoded from a HideObj record.
        /// </summary>
        public ushort? HiddenObjectsMode { get; private set; }

        /// <summary>
        /// Gets whether the workbook supports natural language formulas.
        /// </summary>
        public bool? UsesNaturalLanguageFormulas { get; private set; }

        /// <summary>
        /// Gets whether the workbook stream contains a RefreshAll marker.
        /// </summary>
        public bool HasRefreshAllMarker { get; private set; }

        /// <summary>
        /// Gets whether the workbook stream contains an ObProj marker for a VBA project.
        /// </summary>
        public bool HasVbaProjectMarker { get; private set; }

        /// <summary>
        /// Gets whether the workbook stream declares a VBA project with no forms, modules, or class modules.
        /// </summary>
        public bool HasVbaProjectWithoutMacros { get; private set; }

        /// <summary>
        /// Gets whether workbook windows are locked from moving or resizing.
        /// </summary>
        public bool? WindowsLocked { get; private set; }

        /// <summary>
        /// Gets whether workbook revision tracking is locked.
        /// </summary>
        public bool? RevisionTrackingLocked { get; private set; }

        /// <summary>
        /// Gets the legacy password verifier for workbook revision-tracking protection.
        /// </summary>
        public ushort? RevisionTrackingPasswordHash { get; private set; }

        /// <summary>
        /// Gets the raw print size mode decoded from a PrintSize record.
        /// </summary>
        public ushort? PrintSize { get; private set; }

        /// <summary>
        /// Gets the user name stored by the WriteAccess record, if present.
        /// </summary>
        public string? LastWriteUserName { get; private set; }

        /// <summary>
        /// Gets workbook write-reservation metadata parsed from a FileSharing record, if present.
        /// </summary>
        public LegacyXlsWriteReservation? WriteReservation { get; private set; }

        /// <summary>
        /// Gets parsed workbook protection metadata.
        /// </summary>
        public LegacyXlsWorkbookProtection? Protection { get; private set; }

        internal List<LegacyXlsWorksheet> MutableWorksheets => _worksheets;

        internal List<LegacyXlsNumberFormat> MutableNumberFormats => _numberFormats;

        internal List<LegacyXlsFont> MutableFonts => _fonts;

        internal List<string> MutablePaletteColors => _paletteColors;

        internal List<LegacyXlsCellFormat> MutableCellFormats => _cellFormats;

        internal List<LegacyXlsDefinedName> MutableDefinedNames => _definedNames;

        internal List<LegacyXlsExternalReference> MutableExternalReferences => _externalReferences;

        internal List<LegacyXlsExternalQueryConnection> MutableExternalQueryConnections => _externalQueryConnections;

        internal List<LegacyXlsDataConsolidationReference> MutableDataConsolidationReferences => _dataConsolidationReferences;

        internal List<LegacyXlsDataConsolidationName> MutableDataConsolidationNames => _dataConsolidationNames;

        internal List<LegacyXlsPivotTableRecord> MutablePivotTableRecords => _pivotTableRecords;

        internal List<LegacyXlsChartRecord> MutableChartRecords => _chartRecords;

        internal List<LegacyXlsChartSheet> MutableChartSheets => _chartSheets;

        internal List<LegacyXlsDrawingRecord> MutableDrawingRecords => _drawingRecords;

        internal List<LegacyXlsThemeRecord> MutableThemeRecords => _themeRecords;

        internal List<LegacyXlsDifferentialFormat> MutableDifferentialFormats => _differentialFormats;

        internal List<LegacyXlsTableStyleCollection> MutableTableStyleCollections => _tableStyleCollections;

        internal List<LegacyXlsTableStyle> MutableTableStyles => _tableStyles;

        internal List<LegacyXlsCompoundFeatureRecord> MutableCompoundFeatureRecords => _compoundFeatureRecords;

        internal LegacyXlsCalculationSettings MutableCalculationSettings => _calculationSettings;

        internal List<LegacyXlsUnsupportedSheet> MutableUnsupportedSheets => _unsupportedSheets;

        internal List<LegacyXlsUnsupportedFeature> MutableUnsupportedFeatures => _unsupportedFeatures;

        internal List<LegacyXlsPreservedFeatureRecord> MutablePreservedFeatureRecords => _preservedFeatureRecords;

        internal void SetDocumentProperties(LegacyXlsDocumentProperties properties) {
            DocumentProperties = properties ?? throw new ArgumentNullException(nameof(properties));
        }

        internal List<LegacyXlsFormulaTokenRecord> MutableFormulaTokenRecords => _formulaTokenRecords;

        internal List<LegacyXlsFutureFunctionAlias> MutableFutureFunctionAliases => _futureFunctionAliases;

        internal List<LegacyXlsImportDiagnostic> MutableDiagnostics => _diagnostics;

        internal void AddCellStyle(LegacyXlsCellStyle style) {
            _cellStyles.Add(style);
        }

        internal void AddCellStyleExtension(LegacyXlsCellStyleExtension extension) {
            _cellStyleExtensions.Add(extension);
        }

        internal void AddTableStyleElement(LegacyXlsTableStyleElement element) {
            if (_tableStyles.Count == 0) {
                return;
            }

            _tableStyles[_tableStyles.Count - 1].AddElement(element);
        }

        internal LegacyXlsCellFormat? GetCellFormat(ushort styleIndex) {
            return styleIndex < _cellFormats.Count ? _cellFormats[styleIndex] : null;
        }

        internal LegacyXlsCellFormat? GetEffectiveCellFormat(ushort styleIndex) {
            return styleIndex < _cellFormats.Count
                ? ResolveEffectiveCellFormat(_cellFormats[styleIndex], new HashSet<ushort>())
                : null;
        }

        internal LegacyXlsFont? GetFont(ushort fontIndex) {
            int index = fontIndex < 4 ? fontIndex : fontIndex > 4 ? fontIndex - 1 : -1;
            return index >= 0 && index < _fonts.Count ? _fonts[index] : null;
        }

        internal bool TryResolveColor(ushort colorIndex, out string? argb) {
            return BiffColorPalette.TryResolve(colorIndex, _paletteColors, out argb);
        }

        internal void SetUses1904DateSystem(bool value) {
            Uses1904DateSystem = value;
        }

        internal void SetCodePage(ushort value) {
            CodePage = value;
        }

        internal void SetCodeName(string? value) {
            CodeName = value;
        }

        internal void SetUserInterfaceCodePage(ushort value) {
            UserInterfaceCodePage = value;
        }

        internal void SetCountry(ushort defaultCountryCode, ushort systemCountryCode) {
            Country = new LegacyXlsCountryInfo(defaultCountryCode, systemCountryCode);
        }

        internal void SetSaveBackup(bool value) {
            SaveBackup = value;
        }

        internal void SetBookOptions(ushort flags) {
            DoNotSaveExternalLinkValues = (flags & 0x0001) != 0;
            HasEnvelope = (flags & 0x0004) != 0;
            EnvelopeVisible = (flags & 0x0008) != 0;
            EnvelopeInitialized = (flags & 0x0010) != 0;
            ExternalLinkUpdateMode = checked((byte)((flags >> 5) & 0x0003));
            HideBordersForInactiveTables = (flags & 0x0100) != 0;
        }

        internal void SetBuiltInFunctionGroupCount(ushort value) {
            BuiltInFunctionGroupCount = value;
        }

        internal void SetHiddenObjectsMode(ushort value) {
            HiddenObjectsMode = value;
        }

        internal void SetUsesNaturalLanguageFormulas(bool value) {
            UsesNaturalLanguageFormulas = value;
        }

        internal void SetHasRefreshAllMarker() {
            HasRefreshAllMarker = true;
        }

        internal void SetHasVbaProjectMarker() {
            HasVbaProjectMarker = true;
        }

        internal void SetHasVbaProjectWithoutMacros() {
            HasVbaProjectMarker = true;
            HasVbaProjectWithoutMacros = true;
        }

        internal void SetWindowsLocked(bool value) {
            WindowsLocked = value;
        }

        internal void SetRevisionTrackingLocked(bool value) {
            RevisionTrackingLocked = value;
        }

        internal void SetRevisionTrackingPasswordHash(ushort value) {
            RevisionTrackingPasswordHash = value;
        }

        internal void SetPrintSize(ushort value) {
            PrintSize = value;
        }

        internal void SetLastWriteUserName(string? value) {
            LastWriteUserName = value;
        }

        internal void SetWriteReservation(bool readOnlyRecommended, ushort? passwordHash, string? userName) {
            WriteReservation = new LegacyXlsWriteReservation(
                readOnlyRecommended,
                passwordHash.HasValue ? passwordHash.Value.ToString("X4") : null,
                userName);
        }

        internal void AddWindow(LegacyXlsWorkbookWindow window) {
            _windows.Add(window);
        }

        internal void AddMetadataRecord(LegacyXlsWorkbookMetadataKind kind, int recordOffset, ushort recordType) {
            _metadataRecords.Add(new LegacyXlsWorkbookMetadataRecord(kind, recordOffset, recordType));
        }

        internal void AddFutureMetadataRecord(LegacyXlsWorkbookFutureMetadataRecord record) {
            _futureMetadataRecords.Add(record);
            AddMetadataRecord(record.Kind, record.RecordOffset, record.RecordType);
        }

        internal void SetSheetTabIds(LegacyXlsSheetTabIdCollection sheetTabIds) {
            SheetTabIds = sheetTabIds;
        }

        internal void SetProtection(bool isProtected) {
            Protection = new LegacyXlsWorkbookProtection(isProtected, Protection?.LegacyPasswordHash);
        }

        internal void SetProtectionPasswordHash(ushort passwordHash) {
            Protection = (Protection ?? new LegacyXlsWorkbookProtection(isProtected: false)).WithLegacyPasswordHash(passwordHash);
        }

        private LegacyXlsCellFormat ResolveEffectiveCellFormat(LegacyXlsCellFormat format, HashSet<ushort> seen) {
            if (format.IsStyle || format.ParentStyleIndex >= _cellFormats.Count || !seen.Add(format.StyleIndex)) {
                return format;
            }

            LegacyXlsCellFormat parent = _cellFormats[format.ParentStyleIndex];
            if (!parent.IsStyle) {
                return format;
            }

            LegacyXlsCellFormat effectiveParent = ResolveEffectiveCellFormat(parent, seen);
            return MergeInheritedCellFormat(format, effectiveParent);
        }

        private static LegacyXlsCellFormat MergeInheritedCellFormat(LegacyXlsCellFormat format, LegacyXlsCellFormat parent) {
            bool inheritNumberFormat = !format.ApplyNumberFormat;
            bool inheritFont = !format.ApplyFont;
            bool inheritAlignment = !format.ApplyAlignment && HasNonDefaultAlignment(parent);
            bool inheritBorder = !format.ApplyBorder && parent.Border != null;
            bool inheritFill = !format.ApplyFill && parent.ApplyFill && parent.FillPattern != 0;
            bool inheritProtection = !format.ApplyProtection && HasNonDefaultProtection(parent);

            return new LegacyXlsCellFormat(
                format.StyleIndex,
                inheritFont ? parent.FontIndex : format.FontIndex,
                inheritNumberFormat ? parent.NumberFormatId : format.NumberFormatId,
                format.IsStyle,
                format.ParentStyleIndex,
                format.ApplyNumberFormat || inheritNumberFormat,
                format.ApplyFont || inheritFont,
                format.ApplyFill || inheritFill,
                inheritFill ? parent.FillPattern : format.FillPattern,
                inheritFill ? parent.FillForegroundColorIndex : format.FillForegroundColorIndex,
                inheritFill ? parent.FillBackgroundColorIndex : format.FillBackgroundColorIndex,
                format.ApplyBorder || inheritBorder,
                format.ApplyAlignment || inheritAlignment,
                inheritAlignment ? parent.HorizontalAlignment : format.HorizontalAlignment,
                inheritAlignment ? parent.VerticalAlignment : format.VerticalAlignment,
                inheritAlignment ? parent.WrapText : format.WrapText,
                inheritAlignment ? parent.TextRotation : format.TextRotation,
                inheritAlignment ? parent.Indent : format.Indent,
                inheritAlignment ? parent.ShrinkToFit : format.ShrinkToFit,
                inheritAlignment ? parent.ReadingOrder : format.ReadingOrder,
                format.ApplyProtection || inheritProtection,
                inheritProtection ? parent.Locked : format.Locked,
                inheritProtection ? parent.FormulaHidden : format.FormulaHidden,
                format.QuotePrefix,
                inheritBorder ? parent.Border : format.Border,
                inheritNumberFormat ? parent.NumberFormatCode : format.NumberFormatCode,
                inheritNumberFormat ? parent.IsBuiltInNumberFormat : format.IsBuiltInNumberFormat,
                inheritNumberFormat ? parent.IsDateLike : format.IsDateLike);
        }

        private static bool HasNonDefaultAlignment(LegacyXlsCellFormat format) {
            return format.ApplyAlignment
                || format.HorizontalAlignment != 0
                || format.VerticalAlignment != 2
                || format.WrapText
                || format.TextRotation != 0
                || format.Indent != 0
                || format.ShrinkToFit
                || format.ReadingOrder != 0;
        }

        private static bool HasNonDefaultProtection(LegacyXlsCellFormat format) {
            return format.ApplyProtection
                || !format.Locked
                || format.FormulaHidden;
        }

        /// <summary>
        /// Loads a legacy `.xls` workbook from a file path.
        /// </summary>
        public static LegacyXlsWorkbook Load(string path, LegacyXlsImportOptions? options = null) {
            if (path == null) {
                throw new ArgumentNullException(nameof(path));
            }

            if (!File.Exists(path)) {
                throw new FileNotFoundException($"File '{path}' doesn't exist.", path);
            }

            return Load(File.ReadAllBytes(path), options);
        }

        /// <summary>
        /// Loads a legacy `.xls` workbook from a stream.
        /// </summary>
        public static LegacyXlsWorkbook Load(Stream stream, LegacyXlsImportOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return Load(buffer.ToArray(), options);
        }

        /// <summary>
        /// Loads a legacy `.xls` workbook from a byte array.
        /// </summary>
        public static LegacyXlsWorkbook Load(byte[] bytes, LegacyXlsImportOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            options ??= new LegacyXlsImportOptions();
            if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compoundFile, out string? compoundError)) {
                var workbook = new LegacyXlsWorkbook();
                if (!string.IsNullOrWhiteSpace(compoundError)) {
                    workbook.MutableDiagnostics.Add(CreateCompoundDiagnostic(compoundError!));
                }

                if (workbook.MutableDiagnostics.Count == 0) {
                    workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-COMPOUND-INVALID",
                        "The input is not a supported OLE compound XLS file."));
                }

                return workbook;
            }

            byte[]? workbookStream = LegacyWorkbookStreamLocator.FindWorkbookStream(compoundFile!.Streams);
            if (workbookStream == null) {
                var workbook = new LegacyXlsWorkbook();
                workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-WORKBOOK-STREAM-MISSING",
                    "The OLE compound file does not contain a Workbook or Book stream."));
                return workbook;
            }

            if (workbookStream.Length > options.MaxInputBytes) {
                var workbook = new LegacyXlsWorkbook();
                workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-WORKBOOK-STREAM-TOO-LARGE",
                    $"The BIFF workbook stream is {workbookStream.Length} bytes, which exceeds the configured limit of {options.MaxInputBytes} bytes."));
                return workbook;
            }

            LegacyXlsWorkbook parsedWorkbook = LegacyBiffWorkbookParser.Parse(workbookStream, options);
            parsedWorkbook.SourceCompoundFile = compoundFile;
            LegacyOleDocumentPropertyReader.AddDocumentProperties(compoundFile, parsedWorkbook, options);
            LegacyCompoundFeatureScanner.AddPreserveOnlyFeatures(compoundFile, parsedWorkbook, options);
            return parsedWorkbook;
        }

        private static LegacyXlsImportDiagnostic CreateCompoundDiagnostic(string message) {
            if (message.IndexOf("signature", StringComparison.OrdinalIgnoreCase) >= 0) {
                return new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-COMPOUND-SIGNATURE",
                    message);
            }

            if (message.IndexOf("sector sizes", StringComparison.OrdinalIgnoreCase) >= 0) {
                return new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-COMPOUND-SECTOR-SIZE",
                    message);
            }

            return new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Error,
                "XLS-COMPOUND-CORRUPT",
                message);
        }

        /// <summary>
        /// Projects this legacy workbook into a normal OfficeIMO Excel document.
        /// </summary>
        public ExcelDocument ToExcelDocument() {
            return LegacyXlsWorkbookProjector.ToExcelDocument(this);
        }

        /// <summary>
        /// Creates a compact import report for corpus baselines and preflight checks.
        /// </summary>
        public LegacyXlsImportReport CreateImportReport() {
            return new LegacyXlsImportReport(this);
        }
    }
}
