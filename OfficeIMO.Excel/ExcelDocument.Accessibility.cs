using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        // English plus the default sheet prefixes of the most common Excel display languages.
        private static readonly Regex DefaultSheetNamePattern = new Regex(
            @"^(Sheet|Arkusz|Tabelle|Feuil|Hoja|Foglio|Planilha|Plan|Blad|Blatt|List|Munka|Sayfa|Ark|Лист|Аркуш|Taulukko|Φύλλο)\s?\d+$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>
        /// Inspects the workbook against an accessibility policy: alternative text for pictures, charts and shapes,
        /// table header rows, merged cells, default sheet names, hidden sheets with content, hyperlink text, text contrast
        /// and blank rows or columns used for layout.
        /// </summary>
        /// <remarks>Locations use sheet-qualified A1 references such as <c>Data!B3</c>; <see cref="OfficeAccessibilityFinding.LocationIndex"/> is the zero-based sheet index.</remarks>
        public OfficeAccessibilityReport InspectAccessibility(ExcelAccessibilityOptions? options = null) {
            options ??= new ExcelAccessibilityOptions();
            var findings = new List<OfficeAccessibilityFinding>();
            foreach (ExcelAccessibilityHit hit in InspectAccessibilityHits(options)) findings.Add(hit.Finding);
            return new OfficeAccessibilityReport("Excel", options.Profile, findings);
        }

        /// <summary>
        /// Reports workbook accessibility and compliance issues such as missing alt text, hidden sheets, merged cells, and missing table headers.
        /// </summary>
        [Obsolete("Use InspectAccessibility, which returns the OfficeAccessibilityReport shared with Word and PowerPoint.")]
        public ExcelAccessibilityReport AnalyzeAccessibility() {
            var findings = new List<ExcelAccessibilityFinding>();
            foreach (ExcelAccessibilityHit hit in InspectAccessibilityHits(new ExcelAccessibilityOptions())) {
                OfficeAccessibilityFinding finding = hit.Finding;
                findings.Add(new ExcelAccessibilityFinding(LegacyAccessibilityCategory(finding.Code),
                    (ExcelFindingSeverity)(int)finding.Severity, finding.Message, hit.SheetName, hit.Address));
            }
            return new ExcelAccessibilityReport(findings);
        }

        private static string LegacyAccessibilityCategory(string code) => code switch {
            OfficeDiagnosticCodes.AccessibilityMissingAlternativeText => "ImageAltText",
            OfficeDiagnosticCodes.AccessibilityHiddenSheetContent => "HiddenData",
            OfficeDiagnosticCodes.AccessibilityMergedCells => "MergedCells",
            OfficeDiagnosticCodes.AccessibilityMissingTableHeader => "TableHeaders",
            _ => code
        };

        private List<ExcelAccessibilityHit> InspectAccessibilityHits(ExcelAccessibilityOptions options) {
            var hits = new List<ExcelAccessibilityHit>();
            if (options.RequireDocumentTitle && string.IsNullOrWhiteSpace(Properties.Title)) {
                hits.Add(new ExcelAccessibilityHit(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error,
                    OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle, "The workbook does not define a title.", "Workbook"), null, null));
            }

            var context = new ExcelAccessibilityContext(_spreadSheetDocument.WorkbookPart, options);
            IReadOnlyList<ExcelSheet> sheets = Sheets;
            for (int index = 0; index < sheets.Count; index++) {
                InspectSheetAccessibility(sheets[index], index, context, hits);
            }
            return hits;
        }

        private static void InspectSheetAccessibility(ExcelSheet sheet, int sheetIndex, ExcelAccessibilityContext context, List<ExcelAccessibilityHit> hits) {
            ExcelAccessibilityOptions options = context.Options;
            WorksheetPart part = sheet.WorksheetPart;
            Worksheet? worksheet = part.Worksheet;
            string sheetName = sheet.Name;

            void Add(OfficeAccessibilitySeverity severity, string code, string message, string? address,
                uint? objectId = null, string? objectName = null, bool isHeuristic = false,
                double? measuredValue = null, double? requiredValue = null) {
                hits.Add(new ExcelAccessibilityHit(new OfficeAccessibilityFinding(severity, code, message,
                    SheetLocation(sheetName, address), sheetIndex, objectId, objectName, isHeuristic, measuredValue, requiredValue),
                    sheetName, address));
            }

            var hyperlinkCells = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            if (options.CheckMeaningfulLinks && worksheet != null) {
                foreach (Hyperlink link in worksheet.Elements<Hyperlinks>().SelectMany(list => list.Elements<Hyperlink>())) {
                    string? reference = FirstCell(link.Reference?.Value);
                    if (reference != null && !hyperlinkCells.ContainsKey(reference)) hyperlinkCells[reference] = null;
                }
            }

            SheetScan scan = ScanCells(worksheet, context, hyperlinkCells);

            SheetStateValues? state = sheet.SheetElement.State?.Value;
            if (options.CheckHiddenSheets && scan.HasContent && state != null && state.Value != SheetStateValues.Visible) {
                string kind = state.Value == SheetStateValues.VeryHidden ? "very hidden" : "hidden";
                Add(options.Structural, OfficeDiagnosticCodes.AccessibilityHiddenSheetContent,
                    "The worksheet is " + kind + " but contains data, so people navigating the visible sheets can miss it.", null);
            }
            if (options.CheckDefaultSheetNames && scan.HasContent && DefaultSheetNamePattern.IsMatch(sheetName.Trim())) {
                Add(options.Structural, OfficeDiagnosticCodes.AccessibilityDefaultSheetName,
                    "The worksheet keeps the default name '" + sheetName + "', which does not say what it contains.", null);
            }

            if (options.RequireAlternativeText) InspectDrawingAccessibility(part, Add);

            if (options.RequireTableHeaders) {
                foreach (TableDefinitionPart tablePart in part.TableDefinitionParts) {
                    Table? table = tablePart.Table;
                    if (table == null || table.HeaderRowCount == null || table.HeaderRowCount.Value != 0) continue;
                    string name = table.DisplayName?.Value ?? table.Name?.Value ?? string.Empty;
                    Add(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingTableHeader,
                        "Table '" + name + "' has no header row, so its columns have no names for screen readers.",
                        table.Reference?.Value, table.Id?.Value, name);
                }
            }

            if (options.CheckMergedCells && worksheet != null) {
                foreach (MergeCell merge in worksheet.Elements<MergeCells>().SelectMany(list => list.Elements<MergeCell>())) {
                    string? range = merge.Reference?.Value;
                    if (string.IsNullOrWhiteSpace(range)) continue;
                    Add(options.Structural, OfficeDiagnosticCodes.AccessibilityMergedCells,
                        "Merged cells make sorting, filtering and screen-reader navigation harder.", range);
                }
            }

            foreach (KeyValuePair<string, string?> link in hyperlinkCells) {
                if (OfficeAccessibilityText.IsMeaningfulLinkLabel(link.Value)) continue;
                Add(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityUnclearLinkLabel,
                    string.IsNullOrWhiteSpace(link.Value)
                        ? "The hyperlink cell has no visible text."
                        : "The hyperlink text '" + link.Value!.Trim() + "' does not describe its destination.",
                    link.Key);
            }

            if (options.CheckContrast) {
                foreach (ContrastIssue issue in scan.ContrastIssues) {
                    string ratio = issue.Measured.ToString("0.##", CultureInfo.InvariantCulture);
                    string required = issue.Required.ToString("0.##", CultureInfo.InvariantCulture);
                    string cells = issue.CellCount == 1 ? "The cell" : issue.CellCount.ToString(CultureInfo.InvariantCulture) + " cells with this formatting";
                    Add(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityLowContrast,
                        cells + " have text contrast " + ratio + ":1 against the fill; the selected policy requires at least " + required + ":1.",
                        issue.FirstAddress, measuredValue: issue.Measured, requiredValue: issue.Required);
                }
            }

            if (options.CheckBlankCellLayout) InspectBlankLayout(scan, Add);
        }

        private delegate void AccessibilityAdd(OfficeAccessibilitySeverity severity, string code, string message, string? address,
            uint? objectId = null, string? objectName = null, bool isHeuristic = false,
            double? measuredValue = null, double? requiredValue = null);

        private static void InspectDrawingAccessibility(WorksheetPart part, AccessibilityAdd add) {
            Xdr.WorksheetDrawing? drawing = part.DrawingsPart?.WorksheetDrawing;
            if (drawing == null) return;
            foreach (OpenXmlElement anchor in drawing.ChildElements) {
                if (anchor is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor)) continue;
                OpenXmlElement? graphic = anchor.ChildElements.FirstOrDefault(child =>
                    child is Xdr.Picture or Xdr.GraphicFrame or Xdr.Shape or Xdr.GroupShape);
                if (graphic == null) continue;
                Xdr.NonVisualDrawingProperties? properties = graphic.Descendants<Xdr.NonVisualDrawingProperties>().FirstOrDefault();
                if (properties == null || ExcelDrawingAccessibility.IsDecorative(properties)) continue;
                if (!string.IsNullOrWhiteSpace(properties.Description?.Value) || !string.IsNullOrWhiteSpace(properties.Title?.Value)) continue;
                if (graphic is Xdr.Shape shape && !string.IsNullOrWhiteSpace(shape.TextBody?.InnerText)) continue;

                string kind = graphic switch {
                    Xdr.Picture => "picture",
                    Xdr.Shape => "shape",
                    Xdr.GroupShape => "group",
                    _ => DescribeGraphicFrame(graphic.Descendants<A.GraphicData>().FirstOrDefault()?.Uri?.Value)
                };
                add(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText,
                    "The " + kind + " is neither decorative nor described with alternative text.",
                    AnchorAddress(anchor), properties.Id?.Value, properties.Name?.Value);
            }
        }

        private static void InspectBlankLayout(SheetScan scan, AccessibilityAdd add) {
            if (!scan.HasContent || scan.Truncated) return;
            foreach ((int start, int end) in BlankRuns(scan.MinRow, scan.MaxRow, scan.Rows)) {
                string range = ExcelA1.CellReference(start, scan.MinColumn) + ":" + ExcelA1.CellReference(end, scan.MaxColumn);
                string rows = start == end ? "Row " + start.ToString(CultureInfo.InvariantCulture)
                    : "Rows " + start.ToString(CultureInfo.InvariantCulture) + "-" + end.ToString(CultureInfo.InvariantCulture);
                add(OfficeAccessibilitySeverity.Warning, OfficeDiagnosticCodes.AccessibilityBlankCellLayout,
                    rows + " are blank inside the data area; blank cells used for spacing make screen readers announce empty cells.",
                    range, isHeuristic: true);
            }
            foreach ((int start, int end) in BlankRuns(scan.MinColumn, scan.MaxColumn, scan.Columns)) {
                string first = ExcelA1.ColumnIndexToLetters(start);
                string last = ExcelA1.ColumnIndexToLetters(end);
                string range = ExcelA1.CellReference(scan.MinRow, start) + ":" + ExcelA1.CellReference(scan.MaxRow, end);
                string columns = start == end ? "Column " + first : "Columns " + first + "-" + last;
                add(OfficeAccessibilitySeverity.Warning, OfficeDiagnosticCodes.AccessibilityBlankCellLayout,
                    columns + " are blank inside the data area; blank cells used for spacing make screen readers announce empty cells.",
                    range, isHeuristic: true);
            }
        }

        private static IEnumerable<(int Start, int End)> BlankRuns(int min, int max, HashSet<int> used) {
            int runStart = 0;
            for (int index = min; index <= max; index++) {
                if (!used.Contains(index)) {
                    if (runStart == 0) runStart = index;
                } else if (runStart != 0) {
                    yield return (runStart, index - 1);
                    runStart = 0;
                }
            }
        }

        private static SheetScan ScanCells(Worksheet? worksheet, ExcelAccessibilityContext context, Dictionary<string, string?> hyperlinkCells) {
            var scan = new SheetScan();
            SheetData? data = worksheet?.GetFirstChild<SheetData>();
            if (data == null) return scan;
            var contrastByStyle = new Dictionary<uint, ContrastIssue?>();
            var issues = new Dictionary<uint, ContrastIssue>();
            int visited = 0;
            uint previousRow = 0;
            foreach (Row row in data.Elements<Row>()) {
                uint rowIndex = row.RowIndex?.Value ?? previousRow + 1;
                previousRow = rowIndex;
                int previousColumn = 0;
                foreach (Cell cell in row.Elements<Cell>()) {
                    if (++visited > context.Options.MaximumCellCount) {
                        scan.Truncated = true;
                        break;
                    }
                    int column = previousColumn + 1;
                    string? reference = cell.CellReference?.Value;
                    if (!string.IsNullOrEmpty(reference)) {
                        string letters = new string(reference!.TakeWhile(char.IsLetter).ToArray());
                        if (letters.Length > 0) column = ExcelA1.ColumnLettersToIndex(letters);
                    }
                    previousColumn = column;
                    string address = ExcelA1.CellReference((int)rowIndex, column);
                    string? text = context.CellText(cell);
                    if (hyperlinkCells.ContainsKey(address)) hyperlinkCells[address] = text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    scan.Add((int)rowIndex, column);
                    if (!context.Options.CheckContrast) continue;
                    uint style = cell.StyleIndex?.Value ?? 0;
                    if (!contrastByStyle.TryGetValue(style, out ContrastIssue? template)) {
                        template = context.MeasureContrast(style);
                        contrastByStyle[style] = template;
                    }
                    if (template == null) continue;
                    if (issues.TryGetValue(style, out ContrastIssue? issue)) issue.CellCount++;
                    else issues[style] = new ContrastIssue(template.Measured, template.Required) { FirstAddress = address, CellCount = 1 };
                }
                if (scan.Truncated) break;
            }
            scan.ContrastIssues.AddRange(issues.Values);
            return scan;
        }

        private static string DescribeGraphicFrame(string? uri) {
            if (uri == null) return "object";
            if (uri.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0) return "chart";
            if (uri.EndsWith("/diagram", StringComparison.Ordinal)) return "SmartArt graphic";
            return "object";
        }

        private static string? AnchorAddress(OpenXmlElement anchor) {
            Xdr.FromMarker? from = anchor.GetFirstChild<Xdr.FromMarker>();
            if (from == null) return null;
            if (!int.TryParse(from.RowId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int row) ||
                !int.TryParse(from.ColumnId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int column)) return null;
            return ExcelA1.CellReference(row + 1, column + 1);
        }

        private static string? FirstCell(string? reference) {
            if (string.IsNullOrWhiteSpace(reference)) return null;
            string first = reference!.Split(':')[0].Replace("$", string.Empty).Trim();
            return first.Length == 0 ? null : first.ToUpperInvariant();
        }

        private static string SheetLocation(string sheetName, string? address) {
            bool plain = sheetName.Length > 0 && !char.IsDigit(sheetName[0]) && sheetName.All(character => char.IsLetterOrDigit(character) || character == '_');
            string sheet = plain ? sheetName : "'" + sheetName.Replace("'", "''") + "'";
            return string.IsNullOrEmpty(address) ? sheet : sheet + "!" + address;
        }

        private sealed class ExcelAccessibilityHit {
            internal ExcelAccessibilityHit(OfficeAccessibilityFinding finding, string? sheetName, string? address) {
                Finding = finding;
                SheetName = sheetName;
                Address = address;
            }

            internal OfficeAccessibilityFinding Finding { get; }
            internal string? SheetName { get; }
            internal string? Address { get; }
        }

        private sealed class SheetScan {
            internal bool HasContent { get; private set; }
            internal bool Truncated { get; set; }
            internal int MinRow { get; private set; } = int.MaxValue;
            internal int MaxRow { get; private set; }
            internal int MinColumn { get; private set; } = int.MaxValue;
            internal int MaxColumn { get; private set; }
            internal HashSet<int> Rows { get; } = new HashSet<int>();
            internal HashSet<int> Columns { get; } = new HashSet<int>();
            internal List<ContrastIssue> ContrastIssues { get; } = new List<ContrastIssue>();

            internal void Add(int row, int column) {
                HasContent = true;
                Rows.Add(row);
                Columns.Add(column);
                MinRow = Math.Min(MinRow, row);
                MaxRow = Math.Max(MaxRow, row);
                MinColumn = Math.Min(MinColumn, column);
                MaxColumn = Math.Max(MaxColumn, column);
            }
        }

        private sealed class ContrastIssue {
            internal ContrastIssue(double measured, double required) {
                Measured = measured;
                Required = required;
            }

            internal double Measured { get; }
            internal double Required { get; }
            internal string? FirstAddress { get; set; }
            internal int CellCount { get; set; }
        }

        private sealed class ExcelAccessibilityContext {
            private readonly WorkbookPart? _workbookPart;
            private readonly SharedStringItem[] _sharedStrings;
            private readonly CellFormat[] _cellFormats;
            private readonly Font[] _fonts;
            private readonly Fill[] _fills;

            internal ExcelAccessibilityContext(WorkbookPart? workbookPart, ExcelAccessibilityOptions options) {
                _workbookPart = workbookPart;
                Options = options;
                _sharedStrings = workbookPart?.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();
                Stylesheet? styles = workbookPart?.WorkbookStylesPart?.Stylesheet;
                _cellFormats = styles?.CellFormats?.Elements<CellFormat>().ToArray() ?? Array.Empty<CellFormat>();
                _fonts = styles?.Fonts?.Elements<Font>().ToArray() ?? Array.Empty<Font>();
                _fills = styles?.Fills?.Elements<Fill>().ToArray() ?? Array.Empty<Fill>();
            }

            internal ExcelAccessibilityOptions Options { get; }

            internal string? CellText(Cell cell) {
                CellValues? type = cell.DataType?.Value;
                if (type == CellValues.InlineString) return cell.InlineString?.InnerText;
                string? value = cell.CellValue?.Text;
                if (type == CellValues.SharedString) {
                    return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) &&
                        index >= 0 && index < _sharedStrings.Length ? _sharedStrings[index].InnerText : null;
                }
                if (string.IsNullOrEmpty(value) && cell.CellFormula != null) return "=" + cell.CellFormula.Text;
                return value;
            }

            /// <summary>Returns the contrast failure for a cell format, or null when it passes or uses default colours.</summary>
            internal ContrastIssue? MeasureContrast(uint styleIndex) {
                if (styleIndex >= _cellFormats.Length) return null;
                CellFormat format = _cellFormats[styleIndex];
                Font? font = format.FontId?.Value is uint fontId && fontId < _fonts.Length ? _fonts[fontId] : null;
                Fill? fill = format.FillId?.Value is uint fillId && fillId < _fills.Length ? _fills[fillId] : null;
                OfficeColor? foreground = ResolveColor(font?.Color);
                OfficeColor? background = null;
                PatternFill? pattern = fill?.PatternFill;
                if (pattern?.PatternType?.Value == PatternValues.Solid) background = ResolveColor(pattern.ForegroundColor);
                if (foreground == null && background == null) return null;

                double size = font?.FontSize?.Val?.Value ?? 11D;
                bool bold = font?.Bold != null && (font.Bold.Val == null || font.Bold.Val.Value);
                bool large = size >= 18D || bold && size >= 14D;
                double required = large ? Options.MinimumLargeTextContrastRatio : Options.MinimumTextContrastRatio;
                double measured = OfficeColorContrast.ContrastRatio(foreground ?? OfficeColor.Black, background ?? OfficeColor.White);
                return measured + 0.000001D < required ? new ContrastIssue(measured, required) : null;
            }

            private OfficeColor? ResolveColor(ColorType? color) {
                if (color == null || color.Auto?.Value == true) return null;
                string? argb = ExcelThemeColorResolver.Resolve(color, _workbookPart);
                if (argb == null || argb.Length < 6) return null;
                string rgb = argb.Substring(argb.Length - 6);
                return OfficeColor.TryParseHex("#" + rgb, out OfficeColor parsed) ? parsed : null;
            }
        }
    }
}
