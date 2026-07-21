using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsWritePreflight {
        internal static void ThrowIfUnsupported(
            ExcelDocument document,
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFontTable fontTable,
            LegacyXlsFormulaNameIndex formulaNameIndex) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            if (fontTable == null) throw new ArgumentNullException(nameof(fontTable));
            if (formulaNameIndex == null) throw new ArgumentNullException(nameof(formulaNameIndex));

            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Workbook workbook = document.WorkbookRoot;
            if (document.InspectSignatures().HasSignatures) {
                ThrowUnsupported("digital signatures");
            }

            if (!SupportsWorkbookPackageParts(workbookPart, out string? workbookPackageReason)) {
                ThrowUnsupported(workbookPackageReason ?? "workbook package parts");
            }

            if (!SupportsWorkbookSingletonElements(workbook, out string? workbookSingletonReason)) {
                ThrowUnsupported(workbookSingletonReason ?? "workbook singleton elements");
            }

            if (!SupportsWorkbookViews(workbook, out string? workbookViewReason)) {
                ThrowUnsupported(workbookViewReason ?? "workbook views");
            }

            if (!SupportsCustomWorkbookViews(workbook, out string? customWorkbookViewReason)) {
                ThrowUnsupported(customWorkbookViewReason ?? "custom workbook views");
            }

            if (!LegacyXlsDefinedNameWriter.SupportsWorkbookDefinedNames(document, sheets, formulaNameIndex, out string? definedNameReason)) {
                ThrowUnsupported(definedNameReason ?? "defined names, named ranges, print areas, or print titles");
            }

            if (!SupportsWorkbookProtection(workbook, out string? workbookProtectionReason)) {
                ThrowUnsupported(workbookProtectionReason ?? "workbook protection");
            }

            if (!SupportsWriteReservation(workbook, out string? writeReservationReason)) {
                ThrowUnsupported(writeReservationReason ?? "write-reservation metadata");
            }

            if (!SupportsWorkbookCalculationProperties(workbook, out string? workbookCalculationReason)) {
                ThrowUnsupported(workbookCalculationReason ?? "workbook calculation properties");
            }

            if (!LegacyXlsWriter.SupportsWorkbookTableStyles(workbookPart.WorkbookStylesPart?.Stylesheet, out string? tableStyleReason)) {
                ThrowUnsupported(tableStyleReason ?? "table styles");
            }

            if (workbook.Elements<WorkbookExtensionList>().Any(extensionList => extensionList.Elements<WorkbookExtension>().Any())) {
                ThrowUnsupported("workbook extension metadata");
            }

            for (int i = 0; i < sheets.Count; i++) {
                ThrowIfUnsupported(sheets[i], i, workbookPart, fontTable, formulaNameIndex);
            }
        }

        private static void ThrowIfUnsupported(
            ExcelSheet sheet,
            int sheetIndex,
            WorkbookPart workbookPart,
            LegacyXlsFontTable fontTable,
            LegacyXlsFormulaNameIndex formulaNameIndex) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            Worksheet? worksheet = worksheetPart.Worksheet;
            if (worksheet == null) {
                return;
            }

            if (!SupportsWorksheetSingletonElements(worksheet, out string? singletonElementReason)) {
                ThrowUnsupported(sheet, singletonElementReason ?? "worksheet singleton elements");
            }

            if (!LegacyXlsAutoFilterWriter.SupportsWorksheetAutoFilter(sheet, out string? autoFilterReason)) {
                ThrowUnsupported(sheet, autoFilterReason ?? "AutoFilter");
            }

            if (!LegacyXlsHyperlinkWriter.SupportsWorksheetHyperlinks(sheet, out string? hyperlinkReason)) {
                ThrowUnsupported(sheet, hyperlinkReason ?? "hyperlinks");
            }

            if (!LegacyXlsDataValidationWriter.SupportsWorksheetDataValidations(sheet, sheetIndex, formulaNameIndex, out string? dataValidationReason)) {
                ThrowUnsupported(sheet, dataValidationReason ?? "data validation");
            }

            if (!LegacyXlsConditionalFormattingWriter.SupportsWorksheetConditionalFormatting(sheet, sheetIndex, formulaNameIndex, out string? conditionalFormattingReason)) {
                ThrowUnsupported(sheet, conditionalFormattingReason ?? "conditional formatting");
            }

            ExcelSheet.HeaderFooterSnapshot headerFooter = sheet.GetHeaderFooter();
            if (headerFooter.HeaderHasPicturePlaceholder || headerFooter.FooterHasPicturePlaceholder) {
                ThrowUnsupported(sheet, "header or footer images");
            }

            if (!SupportsHeaderFooterText(headerFooter, out string? headerFooterReason)) {
                ThrowUnsupported(sheet, headerFooterReason ?? "header or footer text");
            }

            if (!SupportsWorksheetPrinterSettings(worksheetPart, out string? printerSettingsReason)) {
                ThrowUnsupported(sheet, printerSettingsReason ?? "printer settings");
            }

            if (!SupportsWorksheetManualPageBreaks(sheet, out string? pageBreakReason)) {
                ThrowUnsupported(sheet, pageBreakReason ?? "manual page breaks");
            }

            if (!SupportsWorksheetPackageParts(worksheetPart, out string? worksheetPackageReason)) {
                ThrowUnsupported(sheet, worksheetPackageReason ?? "worksheet package parts");
            }

            if (!SupportsWorksheetViewModes(worksheet, out string? viewReason)) {
                ThrowUnsupported(sheet, viewReason ?? "worksheet view modes");
            }

            if (!SupportsWorksheetPanes(worksheet, out string? paneReason)) {
                ThrowUnsupported(sheet, paneReason ?? "worksheet panes");
            }

            if (!LegacyXlsSortWriter.SupportsWorksheetSortState(sheet, out string? sortReason)) {
                ThrowUnsupported(sheet, sortReason ?? "sort states");
            }

            if (!LegacyXlsRichTextCellWriter.SupportsWorksheetCellTextRuns(workbookPart, worksheet, fontTable, out string? richTextReason)) {
                ThrowUnsupported(sheet, richTextReason ?? "rich-text cells");
            }

            if (worksheet.Elements<TableParts>().Any(tableParts => tableParts.Elements<TablePart>().Any())) {
                ThrowUnsupported(sheet, "tables");
            }

            SheetProtection? sheetProtection = worksheet.Elements<SheetProtection>().FirstOrDefault();
            if (sheetProtection != null && !IsSupportedSheetProtection(sheetProtection, out string? sheetProtectionReason)) {
                ThrowUnsupported(sheet, sheetProtectionReason ?? "worksheet protection permission exceptions");
            }

            if (!LegacyXlsProtectedRangeWriter.SupportsWorksheetProtectedRanges(sheet, out string? protectedRangeReason)) {
                ThrowUnsupported(sheet, protectedRangeReason ?? "protected ranges");
            }

            if (!LegacyXlsIgnoredErrorWriter.SupportsWorksheetIgnoredErrors(sheet, out string? ignoredErrorReason)) {
                ThrowUnsupported(sheet, ignoredErrorReason ?? "ignored errors");
            }

            if (!LegacyXlsCellWatchWriter.SupportsWorksheetCellWatches(sheet, out string? cellWatchReason)) {
                ThrowUnsupported(sheet, cellWatchReason ?? "cell watches");
            }

            if (!LegacyXlsScenarioWriter.SupportsWorksheetScenarios(sheet, out string? scenarioReason)) {
                ThrowUnsupported(sheet, scenarioReason ?? "worksheet scenarios");
            }

            if (!LegacyXlsDataConsolidationWriter.SupportsWorksheetDataConsolidation(sheet, out string? dataConsolidationReason)) {
                ThrowUnsupported(sheet, dataConsolidationReason ?? "data consolidation");
            }

            if (HasSparklineMetadata(worksheet)) {
                ThrowUnsupported(sheet, "sparklines");
            }

            if (worksheet.Elements<WorksheetExtensionList>().Any(extensionList => extensionList.Elements<WorksheetExtension>().Any())) {
                ThrowUnsupported(sheet, "worksheet extension metadata");
            }

            if (!SupportsWorksheetMetadataElements(worksheet, out string? worksheetMetadataReason)) {
                ThrowUnsupported(sheet, worksheetMetadataReason ?? "worksheet metadata");
            }

            if (worksheetPart.TableDefinitionParts.Any()) {
                ThrowUnsupported(sheet, "tables");
            }

            if (!LegacyXlsCommentWriter.SupportsWorksheetComments(sheet, fontTable, out string? commentReason)) {
                ThrowUnsupported(sheet, commentReason ?? "comments");
            }

            if (worksheetPart.DrawingsPart != null) {
                ThrowUnsupported(sheet, "drawings, images, or charts");
            }
        }

        private static void ThrowUnsupported(ExcelSheet sheet, string feature) {
            ThrowUnsupported($"{feature} on worksheet '{sheet.Name}'");
        }

        private static void ThrowUnsupported(string feature) {
            throw new NotSupportedException($"Native XLS saving does not yet support {feature}. Save as .xlsx or remove this feature before saving as .xls.");
        }

        private static bool SupportsWorkbookViews(Workbook workbook, out string? reason) {
            reason = null;
            return true;
        }

        private static bool SupportsWorkbookPackageParts(WorkbookPart workbookPart, out string? reason) {
            reason = null;
            if (workbookPart.VbaProjectPart != null) {
                reason = "VBA projects or macros";
                return false;
            }

            if (workbookPart.CustomXmlParts.Any()) {
                reason = "custom XML parts";
                return false;
            }

            if (workbookPart.DataPartReferenceRelationships.Any()) {
                reason = "data part relationships";
                return false;
            }

            if (workbookPart.WorkbookPersonParts.Any()) {
                reason = "threaded comments";
                return false;
            }

            if (workbookPart.ChartsheetParts.Any()
                || workbookPart.DialogsheetParts.Any()
                || workbookPart.MacroSheetParts.Any()
                || workbookPart.InternationalMacroSheetParts.Any()) {
                reason = "unsupported sheet types";
                return false;
            }

            if (workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>().Any()
                || workbookPart.Workbook!.GetFirstChild<PivotCaches>()?.Elements<PivotCache>().Any() == true) {
                reason = "PivotTables";
                return false;
            }

            if (!LegacyXlsExternSheetTable.SupportsDeclaredExternalWorkbookLinks(workbookPart, out string? externalWorkbookLinkReason)) {
                reason = externalWorkbookLinkReason ?? "external workbook links";
                return false;
            }

            if (workbookPart.GetPartsOfType<ConnectionsPart>().Any()
                || HasRelatedPart(workbookPart, "connections")) {
                reason = "connections or query tables";
                return false;
            }

            if (workbookPart.SlicerCacheParts.Any()
                || workbookPart.TimeLineCacheParts.Any()
                || HasRelatedPart(workbookPart, "slicer")
                || HasRelatedPart(workbookPart, "timeline")) {
                reason = "slicers or timelines";
                return false;
            }

            if (workbookPart.CellMetadataPart != null
                || workbookPart.CustomDataPropertiesParts.Any()
                || workbookPart.CustomXmlMappingsPart != null
                || workbookPart.FeaturePropertyBagsPart != null) {
                reason = "workbook metadata extensions";
                return false;
            }

            if (workbookPart.RdArrayParts.Any()
                || workbookPart.RdRichValueParts.Any()
                || workbookPart.CT_RdRichValueStructureParts.Any()
                || workbookPart.RdRichValueTypesParts.Any()
                || workbookPart.RdRichValueWebImagePart != null
                || workbookPart.RdSupportingPropertyBagParts.Any()
                || workbookPart.RdSupportingPropertyBagStructureParts.Any()) {
                reason = "rich data features";
                return false;
            }

            if (workbookPart.RichStylesParts.Any()) {
                reason = "rich styles";
                return false;
            }

            if (workbookPart.VolatileDependenciesPart != null) {
                reason = "volatile dependency metadata";
                return false;
            }

            if (workbookPart.WorkbookRevisionHeaderPart != null
                || workbookPart.WorkbookUserDataPart != null) {
                reason = "revision or user data";
                return false;
            }

            if (workbookPart.ExcelAttachedToolbarsPart != null) {
                reason = "attached toolbars";
                return false;
            }

            return true;
        }

        private static bool SupportsWorkbookSingletonElements(Workbook workbook, out string? reason) {
            reason = null;
            return SupportsSingleWorkbookElement<WorkbookProperties>(workbook, "workbook property elements", out reason)
                && SupportsSingleWorkbookElement<FileSharing>(workbook, "write-reservation elements", out reason)
                && SupportsSingleWorkbookElement<WorkbookProtection>(workbook, "workbook protection elements", out reason)
                && SupportsSingleWorkbookElement<BookViews>(workbook, "book-view containers", out reason)
                && SupportsSingleWorkbookElement<Sheets>(workbook, "sheet collections", out reason)
                && SupportsSingleWorkbookElement<ExternalReferences>(workbook, "external-reference collections", out reason)
                && SupportsSingleWorkbookElement<DefinedNames>(workbook, "defined-name collections", out reason)
                && SupportsSingleWorkbookElement<CalculationProperties>(workbook, "calculation property elements", out reason)
                && SupportsSingleWorkbookElement<PivotCaches>(workbook, "PivotCache collections", out reason)
                && SupportsSingleWorkbookElement<CustomWorkbookViews>(workbook, "custom workbook-view containers", out reason)
                && SupportsSingleWorkbookElement<WorkbookExtensionList>(workbook, "workbook extension-list elements", out reason);
        }

        private static bool SupportsSingleWorkbookElement<TElement>(Workbook workbook, string featureName, out string? reason)
            where TElement : OpenXmlElement {
            reason = null;
            if (workbook.Elements<TElement>().Skip(1).Any()) {
                reason = $"multiple workbook {featureName}";
                return false;
            }

            return true;
        }

        private static bool SupportsCustomWorkbookViews(Workbook workbook, out string? reason) {
            reason = null;
            if (workbook.GetFirstChild<CustomWorkbookViews>()?.Elements<CustomWorkbookView>().Any() == true) {
                reason = "custom workbook views";
                return false;
            }

            return true;
        }

        private static bool SupportsWorkbookProtection(Workbook workbook, out string? reason) {
            reason = null;
            WorkbookProtection? protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection == null) {
                return true;
            }

            if (HasAnyAttribute(
                protection,
                "workbookAlgorithmName",
                "workbookHashValue",
                "workbookSaltValue",
                "workbookSpinCount")) {
                reason = "modern workbook protection hashes";
                return false;
            }

            if (!IsSupportedLegacyHash(protection.WorkbookPassword?.Value)) {
                reason = "invalid workbook protection password hashes";
                return false;
            }

            if (HasAnyAttribute(
                    protection,
                    "revisionsAlgorithmName",
                    "revisionsHashValue",
                    "revisionsSaltValue",
                    "revisionsSpinCount")) {
                reason = "workbook revision protection";
                return false;
            }

            if (!IsSupportedLegacyHash(protection.RevisionsPassword?.Value)) {
                reason = "invalid workbook revision protection password hashes";
                return false;
            }

            return true;
        }

        private static bool SupportsWriteReservation(Workbook workbook, out string? reason) {
            reason = null;
            FileSharing? fileSharing = workbook.GetFirstChild<FileSharing>();
            if (fileSharing == null) {
                return true;
            }

            string? userName = fileSharing.UserName?.Value;
            if (userName != null && userName.Length > 54) {
                reason = "write-reservation user names longer than 54 characters";
                return false;
            }

            if (!IsSupportedLegacyHash(GetFileSharingReservationPassword(fileSharing))) {
                reason = "invalid write-reservation password hashes";
                return false;
            }

            return true;
        }

        private static bool SupportsWorkbookCalculationProperties(Workbook workbook, out string? reason) {
            reason = null;
            CalculationProperties? properties = workbook.GetFirstChild<CalculationProperties>();
            if (properties == null) {
                return true;
            }

            if (properties.HasChildren) {
                reason = "workbook calculation properties";
                return false;
            }

            foreach (OpenXmlAttribute attribute in properties.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)
                    || !IsSupportedWorkbookCalculationAttribute(attribute.LocalName)) {
                    reason = "workbook calculation properties";
                    return false;
                }
            }

            return true;
        }

        private static bool IsSupportedWorkbookCalculationAttribute(string localName) {
            return string.Equals(localName, "calcMode", StringComparison.Ordinal)
                || string.Equals(localName, "iterateCount", StringComparison.Ordinal)
                || string.Equals(localName, "fullPrecision", StringComparison.Ordinal)
                || string.Equals(localName, "refMode", StringComparison.Ordinal)
                || string.Equals(localName, "iterateDelta", StringComparison.Ordinal)
                || string.Equals(localName, "iterate", StringComparison.Ordinal)
                || string.Equals(localName, "calcOnSave", StringComparison.Ordinal);
        }

        private static bool SupportsWorksheetPrinterSettings(WorksheetPart worksheetPart, out string? reason) {
            reason = null;
            IReadOnlyCollection<SpreadsheetPrinterSettingsPart> printerSettingsParts = worksheetPart.SpreadsheetPrinterSettingsParts.ToArray();
            if (printerSettingsParts.Count > 1) {
                reason = "multiple worksheet printer settings parts";
                return false;
            }

            foreach (SpreadsheetPrinterSettingsPart printerSettingsPart in printerSettingsParts) {
                using Stream stream = printerSettingsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (stream.Length > ushort.MaxValue - 2L) {
                    reason = "printer settings payload lengths outside BIFF8 limits";
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsWorksheetManualPageBreaks(ExcelSheet sheet, out string? reason) {
            reason = null;
            IReadOnlyList<int> rowPageBreaks = sheet.GetManualRowPageBreaks();
            if (rowPageBreaks.Count > 1026) {
                reason = "manual row page break counts outside BIFF8 limits";
                return false;
            }

            if (rowPageBreaks.Any(row => row <= 0 || row > ushort.MaxValue)) {
                reason = "manual row page breaks outside BIFF8 worksheet limits";
                return false;
            }

            IReadOnlyList<int> columnPageBreaks = sheet.GetManualColumnPageBreaks();
            if (columnPageBreaks.Count > 255) {
                reason = "manual column page break counts outside BIFF8 limits";
                return false;
            }

            if (columnPageBreaks.Any(column => column <= 0 || column >= 256)) {
                reason = "manual column page breaks outside BIFF8 worksheet limits";
                return false;
            }

            return true;
        }

        private static bool SupportsWorksheetPackageParts(WorksheetPart worksheetPart, out string? reason) {
            reason = null;
            if (worksheetPart.GetPartsOfType<PivotTablePart>().Any()) {
                reason = "PivotTables";
                return false;
            }

            if (worksheetPart.DataPartReferenceRelationships.Any()) {
                reason = "data part relationships";
                return false;
            }

            if (worksheetPart.WorksheetCommentsPart == null
                && (worksheetPart.VmlDrawingParts.Any() || HasRelatedPart(worksheetPart, "vmlDrawing"))) {
                reason = "legacy VML drawings or shapes";
                return false;
            }

            if (worksheetPart.GetPartsOfType<QueryTablePart>().Any()
                || HasRelatedPart(worksheetPart, "queryTable")) {
                reason = "connections or query tables";
                return false;
            }

            if (worksheetPart.EmbeddedPackageParts.Any()
                || worksheetPart.EmbeddedObjectParts.Any()
                || worksheetPart.Worksheet!.Elements<OleObjects>().Any(oleObjects => oleObjects.Elements<OleObject>().Any())) {
                reason = "embedded OLE objects or packages";
                return false;
            }

            if (worksheetPart.ControlPropertiesParts.Any()
                || worksheetPart.EmbeddedControlPersistenceParts.Any()
                || worksheetPart.EmbeddedControlPersistenceBinaryDataParts.Any()
                || worksheetPart.Worksheet!.Elements<Controls>().Any()) {
                reason = "form controls";
                return false;
            }

            if (worksheetPart.SlicersParts.Any()
                || worksheetPart.TimeLineParts.Any()
                || HasRelatedPart(worksheetPart, "slicer")
                || HasRelatedPart(worksheetPart, "timeline")) {
                reason = "slicers or timelines";
                return false;
            }

            if (worksheetPart.ImageParts.Any()
                || worksheetPart.Model3DReferenceRelationshipParts.Any()) {
                reason = "drawings, images, or charts";
                return false;
            }

            if (worksheetPart.SingleCellTablePart != null) {
                reason = "tables";
                return false;
            }

            if (worksheetPart.NamedSheetViewsParts.Any()) {
                reason = "named sheet views";
                return false;
            }

            if (worksheetPart.WorksheetSortMapPart != null) {
                reason = "worksheet sort maps";
                return false;
            }

            if (worksheetPart.CustomPropertyParts.Any()) {
                reason = "worksheet custom properties";
                return false;
            }

            return true;
        }

        private static bool HasRelatedPart(OpenXmlPartContainer container, string relationshipTypeFragment) {
            return container.Parts.Any(part =>
                part.OpenXmlPart.RelationshipType.IndexOf(relationshipTypeFragment, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool IsSupportedSheetProtection(SheetProtection protection, out string? reason) {
            reason = null;
            if (protection.Sheet?.Value == false) {
                reason = "worksheet protection permission exceptions";
                return false;
            }

            if (HasAnyAttribute(protection, "algorithmName", "hashValue", "saltValue", "spinCount")) {
                reason = "modern worksheet protection hashes";
                return false;
            }

            if (!IsSupportedLegacyHash(protection.Password?.Value)) {
                reason = "invalid worksheet protection password hashes";
                return false;
            }

            return true;
        }

        private static bool SupportsWorksheetMetadataElements(Worksheet worksheet, out string? reason) {
            reason = null;
            if (worksheet.Elements<CustomSheetViews>().Any(customSheetViews => customSheetViews.Elements<CustomSheetView>().Any())) {
                reason = "custom sheet views";
                return false;
            }

            if (worksheet.Elements<SheetCalculationProperties>().Any(properties => !SupportsWorksheetCalculationProperties(properties))) {
                reason = "worksheet calculation properties";
                return false;
            }

            foreach (PhoneticProperties properties in worksheet.Elements<PhoneticProperties>()) {
                if (!LegacyXlsWriter.SupportsWorksheetPhoneticProperties(properties, out reason)) {
                    reason ??= "worksheet phonetic settings";
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsWorksheetSingletonElements(Worksheet worksheet, out string? reason) {
            reason = null;
            return SupportsSingleWorksheetElement<SheetProperties>(worksheet, "property elements", out reason)
                && SupportsSingleWorksheetElement<SheetDimension>(worksheet, "dimension elements", out reason)
                && SupportsSingleWorksheetElement<SheetViews>(worksheet, "view containers", out reason)
                && SupportsSingleWorksheetElement<SheetFormatProperties>(worksheet, "format property elements", out reason)
                && SupportsSingleWorksheetElement<Columns>(worksheet, "column collections", out reason)
                && SupportsSingleWorksheetElement<SheetData>(worksheet, "sheet data elements", out reason)
                && SupportsSingleWorksheetElement<MergeCells>(worksheet, "merged-cell collections", out reason)
                && SupportsSingleWorksheetElement<AutoFilter>(worksheet, "AutoFilter", out reason)
                && SupportsSingleWorksheetElement<DataValidations>(worksheet, "data-validation collections", out reason)
                && SupportsSingleWorksheetElement<SortState>(worksheet, "sort-state elements", out reason)
                && SupportsSingleWorksheetElement<Hyperlinks>(worksheet, "hyperlink collections", out reason)
                && SupportsSingleWorksheetElement<PrintOptions>(worksheet, "print-option elements", out reason)
                && SupportsSingleWorksheetElement<PageMargins>(worksheet, "page-margin elements", out reason)
                && SupportsSingleWorksheetElement<PageSetup>(worksheet, "page-setup elements", out reason)
                && SupportsSingleWorksheetElement<HeaderFooter>(worksheet, "header/footer elements", out reason)
                && SupportsSingleWorksheetElement<RowBreaks>(worksheet, "row-break collections", out reason)
                && SupportsSingleWorksheetElement<ColumnBreaks>(worksheet, "column-break collections", out reason)
                && SupportsSingleWorksheetElement<SheetProtection>(worksheet, "worksheet protection elements", out reason)
                && SupportsSingleWorksheetElement<ProtectedRanges>(worksheet, "protected-range collections", out reason)
                && SupportsSingleWorksheetElement<IgnoredErrors>(worksheet, "ignored-error collections", out reason)
                && SupportsSingleWorksheetElement<CellWatches>(worksheet, "cell-watch collections", out reason)
                && SupportsSingleWorksheetElement<Scenarios>(worksheet, "scenario collections", out reason)
                && SupportsSingleWorksheetElement<DataConsolidate>(worksheet, "data-consolidation elements", out reason)
                && SupportsSingleWorksheetElement<SheetCalculationProperties>(worksheet, "worksheet calculation property elements", out reason)
                && SupportsSingleWorksheetElement<PhoneticProperties>(worksheet, "worksheet phonetic setting elements", out reason);
        }

        private static bool SupportsSingleWorksheetElement<TElement>(Worksheet worksheet, string featureName, out string? reason)
            where TElement : OpenXmlElement {
            reason = null;
            if (worksheet.Elements<TElement>().Skip(1).Any()) {
                reason = $"multiple worksheet {featureName}";
                return false;
            }

            return true;
        }

        private static bool HasSparklineMetadata(Worksheet worksheet) {
            foreach (OpenXmlElement descendant in worksheet.Descendants()) {
                if (string.Equals(descendant.LocalName, "sparklineGroups", StringComparison.Ordinal)
                    || string.Equals(descendant.LocalName, "sparkline", StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSupportedLegacyHash(string? value) {
            return string.IsNullOrWhiteSpace(value)
                || ushort.TryParse(value!.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _);
        }

        private static string? GetFileSharingReservationPassword(FileSharing fileSharing) {
            foreach (OpenXmlAttribute attribute in fileSharing.GetAttributes()) {
                if (string.Equals(attribute.LocalName, "reservationPassword", StringComparison.Ordinal)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }

            return null;
        }

        private static bool SupportsHeaderFooterText(ExcelSheet.HeaderFooterSnapshot headerFooter, out string? reason) {
            reason = null;
            string? oddHeaderText = BuildHeaderFooterText(headerFooter.HeaderLeft, headerFooter.HeaderCenter, headerFooter.HeaderRight);
            string? oddFooterText = BuildHeaderFooterText(headerFooter.FooterLeft, headerFooter.FooterCenter, headerFooter.FooterRight);
            if (!SupportsBiffUnicodeStringRecord(oddHeaderText) || !SupportsBiffUnicodeStringRecord(oddFooterText)) {
                reason = "header or footer text lengths outside BIFF8 limits";
                return false;
            }

            string? firstHeaderText = BuildHeaderFooterText(headerFooter.FirstHeaderLeft, headerFooter.FirstHeaderCenter, headerFooter.FirstHeaderRight);
            string? firstFooterText = BuildHeaderFooterText(headerFooter.FirstFooterLeft, headerFooter.FirstFooterCenter, headerFooter.FirstFooterRight);
            string? evenHeaderText = BuildHeaderFooterText(headerFooter.EvenHeaderLeft, headerFooter.EvenHeaderCenter, headerFooter.EvenHeaderRight);
            string? evenFooterText = BuildHeaderFooterText(headerFooter.EvenFooterLeft, headerFooter.EvenFooterCenter, headerFooter.EvenFooterRight);

            if (!SupportsHeaderFooterExtensionRecord(firstHeaderText, firstFooterText, evenHeaderText, evenFooterText)) {
                reason = "header or footer text lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static bool SupportsBiffUnicodeStringRecord(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return true;
            }

            string text = value!;
            return text.Length <= ushort.MaxValue
                && 3L + GetEncodedHeaderFooterByteCount(text) <= ushort.MaxValue;
        }

        private static bool SupportsHeaderFooterExtensionRecord(params string?[] values) {
            long payloadLength = 38L;
            foreach (string? value in values) {
                if (string.IsNullOrEmpty(value)) {
                    continue;
                }

                string text = value!;
                if (text.Length > ushort.MaxValue) {
                    return false;
                }

                payloadLength += 1L + GetEncodedHeaderFooterByteCount(text);
                if (payloadLength > ushort.MaxValue) {
                    return false;
                }
            }

            return true;
        }

        private static string? BuildHeaderFooterText(string? left, string? center, string? right) {
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(left)) {
                builder.Append("&L").Append(EscapeHeaderFooterText(left!));
            }

            if (!string.IsNullOrEmpty(center)) {
                builder.Append("&C").Append(EscapeHeaderFooterText(center!));
            }

            if (!string.IsNullOrEmpty(right)) {
                builder.Append("&R").Append(EscapeHeaderFooterText(right!));
            }

            return builder.Length == 0 ? null : builder.ToString();
        }

        private static string EscapeHeaderFooterText(string text) {
            var builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '&' && (i + 1 >= text.Length || !IsHeaderFooterTokenStarter(text[i + 1]))) {
                    builder.Append("&&");
                } else {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private static bool IsHeaderFooterTokenStarter(char ch) {
            return char.IsDigit(ch)
                || ch == '&'
                || ch == 'A'
                || ch == 'B'
                || ch == 'D'
                || ch == 'E'
                || ch == 'F'
                || ch == 'G'
                || ch == 'I'
                || ch == 'K'
                || ch == 'L'
                || ch == 'N'
                || ch == 'P'
                || ch == 'R'
                || ch == 'S'
                || ch == 'T'
                || ch == 'U'
                || ch == 'X'
                || ch == 'Y'
                || ch == 'Z';
        }

        private static long GetEncodedHeaderFooterByteCount(string value) {
            for (int i = 0; i < value.Length; i++) {
                if (value[i] > 0x7f) {
                    return (long)value.Length * 2L;
                }
            }

            return value.Length;
        }

        private static bool SupportsWorksheetCalculationProperties(SheetCalculationProperties properties) {
            if (properties.HasChildren) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in properties.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)
                    || !string.Equals(attribute.LocalName, "fullCalcOnLoad", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsWorksheetPanes(Worksheet worksheet, out string? reason) {
            reason = null;
            foreach (SheetView sheetView in worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>() ?? Enumerable.Empty<SheetView>()) {
                Pane? pane = sheetView.GetFirstChild<Pane>();
                if (pane == null) {
                    continue;
                }

                PaneStateValues? state = pane.State?.Value;
                if (state == PaneStateValues.Frozen || state == PaneStateValues.FrozenSplit) {
                    continue;
                }

                if (state.HasValue || HasSplit(pane.HorizontalSplit) || HasSplit(pane.VerticalSplit)) {
                    if (state.HasValue && state.Value != PaneStateValues.Split) {
                        reason = "worksheet panes";
                        return false;
                    }

                    if (!IsSupportedSplitPane(pane, out reason)) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsSupportedSplitPane(Pane pane, out string? reason) {
            reason = null;
            if (!IsSupportedSplitCoordinate(pane.HorizontalSplit?.Value) || !IsSupportedSplitCoordinate(pane.VerticalSplit?.Value)) {
                reason = "split pane coordinates";
                return false;
            }

            if (pane.TopLeftCell?.Value != null
                && (!A1.TryParseCellReferenceFast(pane.TopLeftCell.Value, out int row, out int column)
                    || row < 1
                    || row > 65536
                    || column < 1
                    || column > 256)) {
                reason = "split pane top-left cells";
                return false;
            }

            PaneValues? activePane = pane.ActivePane?.Value;
            if (activePane.HasValue
                && activePane.Value != PaneValues.BottomRight
                && activePane.Value != PaneValues.TopRight
                && activePane.Value != PaneValues.BottomLeft
                && activePane.Value != PaneValues.TopLeft) {
                reason = "split pane active panes";
                return false;
            }

            return true;
        }

        private static bool IsSupportedSplitCoordinate(double? value) {
            if (!value.HasValue || Math.Abs(value.Value) <= double.Epsilon) {
                return true;
            }

            return !double.IsNaN(value.Value)
                && !double.IsInfinity(value.Value)
                && value.Value >= 0D
                && value.Value <= ushort.MaxValue
                && Math.Abs(value.Value - Math.Round(value.Value)) <= double.Epsilon;
        }

        private static bool SupportsWorksheetViewModes(Worksheet worksheet, out string? reason) {
            reason = null;
            IReadOnlyList<SheetView> sheetViews = GetWorksheetSheetViews(worksheet);
            foreach (SheetView sheetView in sheetViews) {
                SheetViewValues? view = sheetView.View?.Value;
                if (!view.HasValue
                    || view.Value == SheetViewValues.Normal
                    || view.Value == SheetViewValues.PageBreakPreview
                    || view.Value == SheetViewValues.PageLayout) {
                    continue;
                }

                reason = $"worksheet view mode '{view.Value}'";
                return false;
            }

            if (sheetViews.Count > 1) {
                for (int i = 0; i < sheetViews.Count; i++) {
                    SheetView sheetView = sheetViews[i];
                    if (sheetView.View?.Value == SheetViewValues.PageLayout) {
                        reason = "multiple worksheet views with page layout mode";
                        return false;
                    }

                    if (sheetView.GetFirstChild<Pane>() != null) {
                        reason = "multiple worksheet views with pane metadata";
                        return false;
                    }

                }
            }

            return true;
        }

        private static IReadOnlyList<SheetView> GetWorksheetSheetViews(Worksheet worksheet) {
            return worksheet.Elements<SheetViews>()
                .SelectMany(sheetViews => sheetViews.Elements<SheetView>())
                .ToArray();
        }

        private static bool EquivalentSheetView(SheetView left, SheetView right) {
            return Same(left.ShowFormulas?.Value, right.ShowFormulas?.Value)
                && Same(left.ShowGridLines?.Value, right.ShowGridLines?.Value)
                && Same(left.ShowRowColHeaders?.Value, right.ShowRowColHeaders?.Value)
                && Same(left.ShowZeros?.Value, right.ShowZeros?.Value)
                && Same(left.RightToLeft?.Value, right.RightToLeft?.Value)
                && Same(left.DefaultGridColor?.Value, right.DefaultGridColor?.Value)
                && Same(left.ColorId?.Value, right.ColorId?.Value)
                && Same(left.ShowOutlineSymbols?.Value, right.ShowOutlineSymbols?.Value)
                && Same(left.TabSelected?.Value, right.TabSelected?.Value)
                && Same(left.View?.Value, right.View?.Value)
                && Same(left.ZoomScale?.Value, right.ZoomScale?.Value)
                && Same(left.ZoomScaleNormal?.Value, right.ZoomScaleNormal?.Value)
                && string.Equals(left.TopLeftCell?.Value, right.TopLeftCell?.Value, StringComparison.Ordinal)
                && EquivalentPane(left.GetFirstChild<Pane>(), right.GetFirstChild<Pane>())
                && EquivalentSelections(left.Elements<Selection>().ToArray(), right.Elements<Selection>().ToArray());
        }

        private static bool EquivalentPane(Pane? left, Pane? right) {
            if (left == null || right == null) {
                return left == null && right == null;
            }

            return Same(left.HorizontalSplit?.Value, right.HorizontalSplit?.Value)
                && Same(left.VerticalSplit?.Value, right.VerticalSplit?.Value)
                && string.Equals(left.TopLeftCell?.Value, right.TopLeftCell?.Value, StringComparison.Ordinal)
                && Same(left.ActivePane?.Value, right.ActivePane?.Value)
                && Same(left.State?.Value, right.State?.Value);
        }

        private static bool EquivalentSelections(IReadOnlyList<Selection> left, IReadOnlyList<Selection> right) {
            if (left.Count != right.Count) {
                return false;
            }

            for (int i = 0; i < left.Count; i++) {
                if (!EquivalentSelection(left[i], right[i])) {
                    return false;
                }
            }

            return true;
        }

        private static bool EquivalentSelection(Selection left, Selection right) {
            return Same(left.Pane?.Value, right.Pane?.Value)
                && string.Equals(left.ActiveCell?.Value, right.ActiveCell?.Value, StringComparison.Ordinal)
                && Same(left.ActiveCellId?.Value, right.ActiveCellId?.Value)
                && string.Equals(left.SequenceOfReferences?.InnerText, right.SequenceOfReferences?.InnerText, StringComparison.Ordinal);
        }

        private static bool Same<T>(T? left, T? right) where T : struct {
            return EqualityComparer<T?>.Default.Equals(left, right);
        }

        private static bool HasAnyAttribute(OpenXmlElement element, params string[] localNames) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                foreach (string localName in localNames) {
                    if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string? GetAttributeValue(OpenXmlElement element, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool HasSplit(DoubleValue? split) {
            return split != null && Math.Abs(split.Value) > double.Epsilon;
        }

        private static bool IsAllowedSelectionFlag(BooleanValue? value) {
            return value == null || value.Value == false;
        }

        private static bool IsLockedPermissionFlag(BooleanValue? value) {
            return value == null || value.Value;
        }
    }
}
