using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Provenance;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Inspects XLSX, XLSM, XLSB, and legacy XLS content using the normal first-party workbook loader.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective, inspectZipPackage: true), Path.GetFileName(filePath), effective);
    }

    /// <summary>Inspects encoded Excel workbook bytes. The file name helps preserve exact template and binary routing.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] workbookBytes,
        string fileName = "workbook.xlsx",
        OfficeContentSafetyOptions? options = null) {
        if (workbookBytes == null) throw new ArgumentNullException(nameof(workbookBytes));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(workbookBytes, effective, inspectZipPackage: true);
        using ExcelDocument document = LoadContentSafetyWorkbook(workbookBytes, fileName, readOnly: true);
        return InspectContentSafetyDocument(document, effective, targets: null);
    }

    /// <summary>Removes exact selected concealed-content findings and emits the same physical Excel format.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] workbookBytes,
        OfficeContentCleanupSelection selection,
        string fileName = "workbook.xlsx",
        OfficeContentCleanupOptions? options = null) {
        if (workbookBytes == null) throw new ArgumentNullException(nameof(workbookBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();

        OfficeContentSafetyReport before = InspectContentSafety(workbookBytes, fileName, options.Inspection);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult((byte[])workbookBytes.Clone(), before, before, Array.Empty<OfficeContentCleanupChange>());

        ExcelFileFormat sourceFormat = ExcelDocumentLoadRouting.DetectFormat(workbookBytes, fileName);
        byte[] mutableBytes = PrepareExcelContentSafetyMutation(workbookBytes, sourceFormat, options);
        using ExcelDocument document = LoadContentSafetyWorkbook(mutableBytes, fileName, readOnly: false);
        var targets = new Dictionary<string, ExcelCleanupTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport current = InspectContentSafetyDocument(document, options.Inspection, targets);
        IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
        foreach (IGrouping<ExcelCleanupTarget, OfficeContentSafetyFinding> group in currentSelection
                .OrderByDescending(item => targets[item.Id].RemovalPriority)
                .ThenByDescending(item => targets[item.Id].Sequence)
                .GroupBy(item => targets[item.Id])) {
            group.Key.Remove();
        }
        document.MarkPackageDirty();

        byte[] output = document.ToBytes(sourceFormat, new ExcelSaveOptions {
            SignatureMutationPolicy = options.SignatureMutationPolicy
        });
        OfficeContentSafetyReport after = InspectContentSafety(output, fileName, options.Inspection);
        OfficeContentCleanupChange[] changes = selected
            .Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability))
            .ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned Excel artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentCleanupResult result = RemoveSelectedContent(OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection, inspectZipPackage: true), selection, Path.GetFileName(inputPath), options);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static ExcelDocument LoadContentSafetyWorkbook(byte[] bytes, string fileName, bool readOnly) =>
        LoadFromByteArray(bytes, new ExcelLoadOptions {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        }, fileName);

    private static OfficeContentSafetyReport InspectContentSafetyDocument(
        ExcelDocument document,
        OfficeContentSafetyOptions? options,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        WorkbookPart workbookPart = document.WorkbookPartRoot ?? throw new InvalidDataException("The package has no workbook part.");
        Workbook workbook = workbookPart.Workbook ?? throw new InvalidDataException("The workbook part has no workbook root.");
        var builder = new OfficeContentSafetyBuilder("Excel " + document.SourceFormat, options);
        var styles = new ExcelContentStyleResolver(workbookPart);
        IReadOnlyList<SharedStringItem> sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray()
            ?? Array.Empty<SharedStringItem>();
        List<Sheet> sheets = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();

        for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
            Sheet sheet = sheets[sheetIndex];
            string? relationshipId = sheet.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId) || workbookPart.GetPartById(relationshipId!) is not WorksheetPart worksheetPart || worksheetPart.Worksheet == null) continue;
            string sheetName = sheet.Name?.Value ?? "Sheet" + (sheetIndex + 1).ToString(CultureInfo.InvariantCulture);
            SheetStateValues? state = sheet.State?.Value;
            bool hiddenSheet = state == SheetStateValues.Hidden || state == SheetStateValues.VeryHidden;
            InspectExcelWorksheet(workbookPart, worksheetPart, sheetName, hiddenSheet, sharedStrings, styles, builder, targets);
        }

        if (builder.Options.IncludeNonPrimaryContent) {
            int nameIndex = 0;
            foreach (DefinedName name in workbook.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
                if (name.Hidden?.Value != true || string.IsNullOrWhiteSpace(name.Text)) continue;
                string location = "DefinedName[" + (++nameIndex).ToString(CultureInfo.InvariantCulture) + "](" + (name.Name?.Value ?? "unnamed") + ")";
                OfficeContentSafetyFinding finding = builder.Add(
                    OfficeContentConcealmentKind.HiddenByProperty,
                    OfficeContentSafetyRisk.ContextDependent,
                    location,
                    "The workbook defined name has its native hidden flag enabled.",
                    name.Text,
                    OfficeContentCleanupCapability.RemoveElement);
                if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForElement(name);
            }
        }

        if (workbookPart.WorkbookStylesPart?.Stylesheet?.DifferentialFormats?.ChildElements.Count > 0) {
            builder.AddDiagnostic("Conditional-format differential styles are retained as context but are not evaluated as a final rendered color because their formulas and precedence depend on the active calculation state.");
        }
        foreach (string diagnostic in document.SourceFormat == ExcelFileFormat.Xlsb
            ? document.XlsbImportDiagnostics.Select(item => item.ToString())
            : document.SourceFormat == ExcelFileFormat.Xls
                ? document.LegacyXlsImportDiagnostics.Select(item => item.ToString())
                : Enumerable.Empty<string>()) {
            builder.AddDiagnostic(diagnostic);
        }
        return builder.Build();
    }

    private static void InspectExcelWorksheet(
        WorkbookPart workbookPart,
        WorksheetPart part,
        string sheetName,
        bool hiddenSheet,
        IReadOnlyList<SharedStringItem> sharedStrings,
        ExcelContentStyleResolver styles,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        Worksheet worksheet = part.Worksheet!;
        List<Column> columns = worksheet.GetFirstChild<Columns>()?.Elements<Column>().ToList() ?? new List<Column>();
        SheetFormatProperties? format = worksheet.GetFirstChild<SheetFormatProperties>();
        bool defaultRowsHidden = format?.ZeroHeight?.Value == true;
        int fallbackRow = 0;
        foreach (Row row in worksheet.GetFirstChild<SheetData>()?.Elements<Row>() ?? Enumerable.Empty<Row>()) {
            int rowIndex = checked((int)(row.RowIndex?.Value ?? (uint)(++fallbackRow)));
            bool hiddenRow = defaultRowsHidden || row.Hidden?.Value == true || (row.Height?.Value == 0D && row.CustomHeight?.Value == true);
            foreach (Cell cell in row.Elements<Cell>()) {
                string reference = cell.CellReference?.Value ?? "R" + rowIndex.ToString(CultureInfo.InvariantCulture) + "C?";
                int columnIndex = TryGetExcelColumnIndex(reference);
                Column? column = columns.LastOrDefault(item => columnIndex > 0 && item.Min?.Value <= (uint)columnIndex && item.Max?.Value >= (uint)columnIndex);
                bool hiddenColumn = column?.Hidden?.Value == true || (column?.Width?.Value == 0D && column.CustomWidth?.Value == true);
                string text = GetExcelCellText(cell, sharedStrings);
                string formula = cell.CellFormula?.Text ?? string.Empty;
                string payload = string.IsNullOrEmpty(formula) ? text : text + "\nFormula: " + formula;
                if (string.IsNullOrWhiteSpace(payload)) continue;
                string location = "Worksheet(" + sheetName + ")/Cell(" + reference + ")";
                ExcelEffectiveCellStyle style = styles.Resolve(cell, row, column);
                OfficeContentConcealmentKind? kind = null;
                string? evidence = null;
                if (hiddenSheet) {
                    kind = OfficeContentConcealmentKind.HiddenContainer;
                    evidence = "The owning worksheet is hidden or very hidden.";
                } else if (hiddenRow) {
                    kind = OfficeContentConcealmentKind.HiddenContainer;
                    evidence = "The owning row is hidden or has zero effective height.";
                } else if (hiddenColumn) {
                    kind = OfficeContentConcealmentKind.HiddenContainer;
                    evidence = "The owning column is hidden or has zero effective width.";
                } else if (style.FontSizePoints.HasValue && style.FontSizePoints.Value <= builder.Options.MaximumTinyFontSizePoints) {
                    kind = OfficeContentConcealmentKind.TinyText;
                    evidence = "The effective cell font size is " + style.FontSizePoints.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
                } else if (style.TransparentText) {
                    kind = OfficeContentConcealmentKind.TransparentText;
                    evidence = "The effective SpreadsheetML font color has zero alpha.";
                } else if (style.HiddenDisplayValue) {
                    kind = OfficeContentConcealmentKind.HiddenDisplayValue;
                    evidence = "The effective custom number format is the canonical all-empty format ';;;'.";
                } else if (style.ContrastRatio.HasValue && style.ContrastRatio.Value < builder.Options.MinimumVisibleContrastRatio) {
                    kind = OfficeContentConcealmentKind.LowContrastText;
                    evidence = "The effective cell foreground/background contrast ratio is " + style.ContrastRatio.Value.ToString("0.###", CultureInfo.InvariantCulture) + ".";
                }

                if (kind.HasValue) {
                    OfficeContentSafetyFinding finding = builder.Add(
                        kind.Value,
                        OfficeContentSafetyRisk.ContextDependent,
                        location,
                        evidence!,
                        payload,
                        OfficeContentCleanupCapability.RemoveText,
                        inspectTextIntegrityEvidence: false);
                    if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForCellPayload(cell);
                    if (TryGetExcelRichTextOwner(cell, sharedStrings, out OpenXmlCompositeElement? concealedRichOwner)) {
                        InspectExcelChargedTextNodes(cell, concealedRichOwner!, location, builder, targets);
                    } else {
                        InspectExcelCellText(cell, sharedStrings, location, builder, targets, alreadyCharged: true);
                    }
                    if (!string.IsNullOrEmpty(formula) && cell.CellFormula != null) {
                        IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectChargedTextIntegrity(location + "/Formula", formula, OfficeContentCleanupCapability.RemoveText);
                        if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = ExcelCleanupTarget.ForTextRange(cell.CellFormula, item);
                    }
                } else if (TryGetExcelRichTextOwner(cell, sharedStrings, out OpenXmlCompositeElement? richOwner)) {
                    InspectExcelRichText(cell, richOwner!, location, style, workbookPart, builder, targets);
                    if (!string.IsNullOrEmpty(formula) && cell.CellFormula != null) {
                        IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectVisibleText(location + "/Formula", formula, OfficeContentCleanupCapability.RemoveText);
                        if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = ExcelCleanupTarget.ForTextRange(cell.CellFormula, item);
                    }
                } else {
                    InspectExcelCellText(cell, sharedStrings, location, builder, targets, alreadyCharged: false);
                    if (!string.IsNullOrEmpty(formula) && cell.CellFormula != null) {
                        IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectVisibleText(location + "/Formula", formula, OfficeContentCleanupCapability.RemoveText);
                        if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = ExcelCleanupTarget.ForTextRange(cell.CellFormula, item);
                    }
                }
            }
        }

        if (builder.Options.IncludeNonPrimaryContent) {
            foreach (Comment comment in part.WorksheetCommentsPart?.Comments?.CommentList?.Elements<Comment>() ?? Enumerable.Empty<Comment>()) {
                string text = comment.CommentText?.InnerText ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text)) continue;
                string reference = comment.Reference?.Value ?? "unknown";
                OfficeContentSafetyFinding finding = builder.Add(
                    OfficeContentConcealmentKind.NonPrimaryContent,
                    OfficeContentSafetyRisk.Informational,
                    "Worksheet(" + sheetName + ")/Comment(" + reference + ")",
                    "The text is stored in a legacy cell comment rather than the visible cell value.",
                    text,
                    OfficeContentCleanupCapability.RemoveElement);
                if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForElement(comment);
            }
            foreach (WorksheetThreadedCommentsPart threadedPart in part.WorksheetThreadedCommentsParts) {
                foreach (Threaded.ThreadedComment comment in threadedPart.ThreadedComments?.Elements<Threaded.ThreadedComment>() ?? Enumerable.Empty<Threaded.ThreadedComment>()) {
                    string text = comment.GetFirstChild<Threaded.ThreadedCommentText>()?.Text ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    string reference = comment.Ref?.Value ?? "unknown";
                    OfficeContentSafetyFinding finding = builder.Add(
                        OfficeContentConcealmentKind.NonPrimaryContent,
                        OfficeContentSafetyRisk.Informational,
                        "Worksheet(" + sheetName + ")/ThreadedComment(" + reference + ")",
                        "The text is stored in a threaded cell comment rather than the visible cell value.",
                        text,
                        OfficeContentCleanupCapability.RemoveElement);
                    if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForElement(comment);
                }
            }
            InspectExcelDrawingAlternativeText(part, sheetName, builder, targets);
        }
        InspectExcelDrawingText(part, sheetName, workbookPart, builder, targets);
    }

    private static bool TryGetExcelRichTextOwner(Cell cell, IReadOnlyList<SharedStringItem> sharedStrings, out OpenXmlCompositeElement? owner) {
        if (cell.InlineString != null && cell.InlineString.Elements<Run>().Any()) { owner = cell.InlineString; return true; }
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) &&
            index >= 0 && index < sharedStrings.Count && sharedStrings[index].Elements<Run>().Any()) {
            owner = sharedStrings[index];
            return true;
        }
        owner = null;
        return false;
    }

    private static void InspectExcelCellText(
        Cell cell,
        IReadOnlyList<SharedStringItem> sharedStrings,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets,
        bool alreadyCharged) {
        if (cell.InlineString != null) {
            int index = 0;
            foreach (Text textNode in cell.InlineString.Descendants<Text>()) InspectExcelTextNode(cell, cell.InlineString, textNode, location + "/Text[" + (++index).ToString(CultureInfo.InvariantCulture) + "]", false, builder, targets, alreadyCharged);
            return;
        }
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedIndex) &&
            sharedIndex >= 0 && sharedIndex < sharedStrings.Count) {
            int index = 0;
            SharedStringItem owner = sharedStrings[sharedIndex];
            foreach (Text textNode in owner.Descendants<Text>()) InspectExcelTextNode(cell, owner, textNode, location + "/Text[" + (++index).ToString(CultureInfo.InvariantCulture) + "]", true, builder, targets, alreadyCharged);
            return;
        }
        if (cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text)) {
            IReadOnlyList<OfficeContentSafetyFinding> unicode = alreadyCharged
                ? builder.InspectChargedTextIntegrity(location + "/Value", cell.CellValue.Text, OfficeContentCleanupCapability.RemoveText)
                : builder.InspectVisibleText(location + "/Value", cell.CellValue.Text, OfficeContentCleanupCapability.RemoveText);
            if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = ExcelCleanupTarget.ForTextRange(cell.CellValue, item);
        }
    }

    private static void InspectExcelTextNode(
        Cell cell,
        OpenXmlCompositeElement owner,
        Text textNode,
        string location,
        bool shared,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets,
        bool alreadyCharged = false) {
        string text = textNode.Text ?? string.Empty;
        if (text.Length == 0) return;
        IReadOnlyList<OfficeContentSafetyFinding> unicode = alreadyCharged
            ? builder.InspectChargedTextIntegrity(location, text, OfficeContentCleanupCapability.RemoveText)
            : builder.InspectVisibleText(location, text, OfficeContentCleanupCapability.RemoveText);
        if (targets == null) return;
        foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = shared
            ? ExcelCleanupTarget.ForSharedTextRange(cell, owner, textNode, item)
            : ExcelCleanupTarget.ForTextRange(textNode, item);
    }

    private static void InspectExcelRichText(
        Cell cell,
        OpenXmlCompositeElement owner,
        string cellLocation,
        ExcelEffectiveCellStyle cellStyle,
        WorkbookPart workbookPart,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        int index = 0;
        foreach (Run run in owner.Elements<Run>()) {
            string text = run.InnerText;
            if (string.IsNullOrWhiteSpace(text)) continue;
            string location = cellLocation + "/RichRun[" + (++index).ToString(CultureInfo.InvariantCulture) + "]";
            RunProperties? properties = run.RunProperties;
            string? foregroundArgb = ExcelThemeColorResolver.Resolve(properties?.GetFirstChild<Color>(), workbookPart);
            bool transparent = foregroundArgb?.StartsWith("00", StringComparison.OrdinalIgnoreCase) == true;
            double? fontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value;
            OfficeContentConcealmentKind? kind = null;
            string? evidence = null;
            if (fontSize.HasValue && fontSize.Value <= builder.Options.MaximumTinyFontSizePoints) {
                kind = OfficeContentConcealmentKind.TinyText;
                evidence = "The SpreadsheetML rich-text run font size is " + fontSize.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
            } else if (transparent) {
                kind = OfficeContentConcealmentKind.TransparentText;
                evidence = "The SpreadsheetML rich-text run color has zero alpha.";
            } else if (ExcelContentStyleResolver.TryParseExcelColor(foregroundArgb, OfficeColor.Black, out OfficeColor foreground) && cellStyle.Background.HasValue &&
                       OfficeColorContrast.ContrastRatio(foreground, cellStyle.Background.Value) < builder.Options.MinimumVisibleContrastRatio) {
                double ratio = OfficeColorContrast.ContrastRatio(foreground, cellStyle.Background.Value);
                kind = OfficeContentConcealmentKind.LowContrastText;
                evidence = "The rich-text run/cell-background contrast ratio is " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".";
            }
            if (kind.HasValue) {
                OfficeContentSafetyFinding finding = builder.Add(kind.Value, OfficeContentSafetyRisk.ContextDependent, location, evidence!, text, OfficeContentCleanupCapability.RemoveElement, inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForRichRun(cell, owner, run);
            }
            Text? textNode = run.GetFirstChild<Text>();
            if (textNode != null) {
                bool shared = owner is SharedStringItem;
                IReadOnlyList<OfficeContentSafetyFinding> unicode = kind.HasValue
                    ? builder.InspectChargedTextIntegrity(location + "/Text", text, OfficeContentCleanupCapability.RemoveText)
                    : builder.InspectVisibleText(location + "/Text", text, OfficeContentCleanupCapability.RemoveText);
                if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = shared
                    ? ExcelCleanupTarget.ForSharedTextRange(cell, owner, textNode, item)
                    : ExcelCleanupTarget.ForTextRange(textNode, item);
            }
        }
    }

    private static void InspectExcelDrawingText(
        WorksheetPart worksheetPart,
        string sheetName,
        WorkbookPart workbookPart,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        Xdr.WorksheetDrawing? drawing = worksheetPart.DrawingsPart?.WorksheetDrawing;
        if (drawing == null) return;
        int index = 0;
        foreach (A.Run run in drawing.Descendants<A.Run>()) {
            if (run.Text != null) InspectExcelDrawingTextNode(run, run.Text, run.RunProperties, "DrawingRun", ++index, sheetName, workbookPart, builder, targets);
        }
        foreach (A.Field field in drawing.Descendants<A.Field>()) {
            if (field.Text != null) InspectExcelDrawingTextNode(field, field.Text, field.RunProperties, "DrawingField", ++index, sheetName, workbookPart, builder, targets);
        }
    }

    private static void InspectExcelDrawingTextNode(
        OpenXmlElement owner,
        A.Text textNode,
        A.RunProperties? properties,
        string label,
        int index,
        string sheetName,
        WorkbookPart workbookPart,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        string text = textNode.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text)) return;
        string location = "Worksheet(" + sheetName + ")/" + label + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
        Xdr.Shape? shape = owner.Ancestors<Xdr.Shape>().FirstOrDefault();
        int? fontSize = properties?.FontSize?.Value;
        A.SolidFill? fill = properties?.GetFirstChild<A.SolidFill>();
        int? alpha = fill?.Descendants<A.Alpha>().FirstOrDefault()?.Val?.Value;
        OfficeContentConcealmentKind? kind = null;
        string? evidence = null;
        if (TryGetExcelDrawingOwnerConcealment(owner, out OfficeContentConcealmentKind ownerKind, out string ownerEvidence)) {
            kind = ownerKind;
            evidence = ownerEvidence;
        } else if (fontSize.HasValue && fontSize.Value / 100D <= builder.Options.MaximumTinyFontSizePoints) {
            kind = OfficeContentConcealmentKind.TinyText;
            evidence = "The worksheet drawing text font size is " + (fontSize.Value / 100D).ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
        } else if (alpha.HasValue && alpha.Value <= 1000) {
            kind = OfficeContentConcealmentKind.TransparentText;
            evidence = "The worksheet drawing text color is fully or nearly transparent.";
        } else if (TryGetExcelDrawingContrast(fill, shape, workbookPart, out double ratio, out string colors) && ratio < builder.Options.MinimumVisibleContrastRatio) {
            kind = OfficeContentConcealmentKind.LowContrastText;
            evidence = colors + " has contrast ratio " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".";
        }
        if (kind.HasValue) {
            OfficeContentSafetyFinding finding = builder.Add(kind.Value, OfficeContentSafetyRisk.ContextDependent, location, evidence!, text, OfficeContentCleanupCapability.RemoveElement, inspectTextIntegrityEvidence: false);
            if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForElement(owner);
        }
        IReadOnlyList<OfficeContentSafetyFinding> unicode = kind.HasValue
            ? builder.InspectChargedTextIntegrity(location + "/Text", text, OfficeContentCleanupCapability.RemoveText)
            : builder.InspectVisibleText(location + "/Text", text, OfficeContentCleanupCapability.RemoveText);
        if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = ExcelCleanupTarget.ForTextRange(textNode, item);
    }

    private static bool TryGetExcelDrawingOwnerConcealment(
        OpenXmlElement owner,
        out OfficeContentConcealmentKind kind,
        out string evidence) {
        foreach (OpenXmlElement ancestor in owner.Ancestors()) {
            if (ancestor is Xdr.Shape shape) {
                if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value == true) {
                    kind = OfficeContentConcealmentKind.HiddenByProperty;
                    evidence = "The worksheet drawing shape has its native hidden flag enabled.";
                    return true;
                }
                A.Extents? extents = shape.ShapeProperties?.Transform2D?.Extents;
                if (HasZeroExcelDrawingExtents(extents)) {
                    kind = OfficeContentConcealmentKind.ZeroDimension;
                    evidence = "The worksheet drawing shape has zero width or height.";
                    return true;
                }
            } else if (ancestor is Xdr.GroupShape group) {
                if (group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value == true) {
                    kind = OfficeContentConcealmentKind.HiddenContainer;
                    evidence = "An owning worksheet drawing group has its native hidden flag enabled.";
                    return true;
                }
                A.Extents? extents = group.GroupShapeProperties?.TransformGroup?.Extents;
                if (HasZeroExcelDrawingExtents(extents)) {
                    kind = OfficeContentConcealmentKind.ZeroDimension;
                    evidence = "An owning worksheet drawing group has zero width or height.";
                    return true;
                }
            }
        }
        kind = default;
        evidence = string.Empty;
        return false;
    }

    private static bool HasZeroExcelDrawingExtents(A.Extents? extents) =>
        extents != null && ((extents.Cx?.Value ?? 0L) <= 0L || (extents.Cy?.Value ?? 0L) <= 0L);

    private static bool TryGetExcelDrawingContrast(
        A.SolidFill? foregroundFill,
        Xdr.Shape? shape,
        WorkbookPart workbookPart,
        out double ratio,
        out string evidence) {
        ratio = 0D;
        evidence = string.Empty;
        string? foregroundArgb = ExcelThemeColorResolver.Resolve(foregroundFill, workbookPart);
        string? backgroundArgb = ExcelThemeColorResolver.Resolve(shape?.ShapeProperties?.GetFirstChild<A.SolidFill>(), workbookPart);
        if (!ExcelContentStyleResolver.TryParseExcelColor(foregroundArgb, OfficeColor.Black, out OfficeColor foreground) ||
            !ExcelContentStyleResolver.TryParseExcelColor(backgroundArgb, OfficeColor.White, out OfficeColor background)) return false;
        ratio = OfficeColorContrast.ContrastRatio(foreground, background);
        evidence = "Worksheet drawing foreground " + foreground.ToHex() + " against explicit shape background " + background.ToHex();
        return true;
    }

    private static void InspectExcelChargedTextNodes(
        Cell cell,
        OpenXmlCompositeElement owner,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        bool shared = owner is SharedStringItem;
        int index = 0;
        foreach (Text textNode in owner.Descendants<Text>()) {
            InspectExcelTextNode(cell, owner, textNode, location + "/Text[" + (++index).ToString(CultureInfo.InvariantCulture) + "]", shared, builder, targets, alreadyCharged: true);
        }
    }

    private static void InspectExcelDrawingAlternativeText(
        WorksheetPart worksheetPart,
        string sheetName,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, ExcelCleanupTarget>? targets) {
        int index = 0;
        foreach (DrawingsPart drawingPart in worksheetPart.DrawingsPart == null ? Enumerable.Empty<DrawingsPart>() : new[] { worksheetPart.DrawingsPart }) {
            foreach (Xdr.NonVisualDrawingProperties properties in drawingPart.WorksheetDrawing?.Descendants<Xdr.NonVisualDrawingProperties>() ?? Enumerable.Empty<Xdr.NonVisualDrawingProperties>()) {
                foreach ((string Attribute, string Text) item in new[] {
                    ("descr", properties.Description?.Value ?? string.Empty),
                    ("title", properties.Title?.Value ?? string.Empty)
                }) {
                    if (string.IsNullOrWhiteSpace(item.Text)) continue;
                    string location = "Worksheet(" + sheetName + ")/Drawing[" + (++index).ToString(CultureInfo.InvariantCulture) + "]/@" + item.Attribute;
                    OfficeContentSafetyFinding finding = builder.Add(
                        OfficeContentConcealmentKind.NonPrimaryContent,
                        OfficeContentSafetyRisk.Informational,
                        location,
                        "The text is stored as drawing alternative text and is not ordinary visible cell content.",
                        item.Text,
                        OfficeContentCleanupCapability.RemoveText);
                    if (targets != null) targets[finding.Id] = ExcelCleanupTarget.ForAttribute(properties, item.Attribute);
                }
            }
        }
    }

    private static string GetExcelCellText(Cell cell, IReadOnlyList<SharedStringItem> sharedStrings) {
        if (cell.InlineString != null) return cell.InlineString.InnerText;
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) &&
            index >= 0 && index < sharedStrings.Count) return sharedStrings[index].InnerText;
        return cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
    }

    private static int TryGetExcelColumnIndex(string reference) {
        int column = 0;
        foreach (char value in reference) {
            if (value is < 'A' or > 'Z' && value is < 'a' or > 'z') break;
            int digit = char.ToUpperInvariant(value) - 'A' + 1;
            if (column > (16384 - digit) / 26) return 0;
            column = column * 26 + digit;
        }
        return column;
    }

    private static byte[] PrepareExcelContentSafetyMutation(
        byte[] data,
        ExcelFileFormat sourceFormat,
        OfficeContentCleanupOptions cleanupOptions) {
        OfficeSignatureMutationPolicy policy = cleanupOptions.SignatureMutationPolicy;
        if (sourceFormat == ExcelFileFormat.Xls) {
            using ExcelDocument document = LoadContentSafetyWorkbook(data, "workbook.xls", readOnly: true);
            if (document.LegacyXlsCompoundFeatures.Any(item => item.Kind == LegacyXlsCompoundFeatureRecordKind.DigitalSignature)) {
                throw new InvalidOperationException("Legacy XLS digital signatures cannot be preserved or safely stripped during content cleanup. Work from an explicitly unsigned copy.");
            }
            return (byte[])data.Clone();
        }

        OfficeProvenanceRemovalOptions provenanceOptions =
            OfficeContentSafetyProvenanceOptions.CreateSignatureRemovalOptions(cleanupOptions);
        bool hasSignatures = HasPackageSignatures(data, provenanceOptions);
        if (!hasSignatures) return (byte[])data.Clone();
        if (policy == OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException("Content cleanup would invalidate existing Excel package signatures. Select RemoveInvalidatedSignatures or PreserveSignatureMarkup explicitly.");
        }
        return policy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            ? StripPackageSignatures(data, provenanceOptions).Data
            : (byte[])data.Clone();
    }

    private sealed class ExcelContentStyleResolver {
        private readonly WorkbookPart _workbookPart;
        private readonly CellFormat[] _formats;
        private readonly Font[] _fonts;
        private readonly Fill[] _fills;
        private readonly Dictionary<uint, string> _numberFormats;
        internal ExcelContentStyleResolver(WorkbookPart workbookPart) {
            _workbookPart = workbookPart;
            Stylesheet? stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
            _formats = stylesheet?.CellFormats?.Elements<CellFormat>().ToArray() ?? Array.Empty<CellFormat>();
            _fonts = stylesheet?.Fonts?.Elements<Font>().ToArray() ?? Array.Empty<Font>();
            _fills = stylesheet?.Fills?.Elements<Fill>().ToArray() ?? Array.Empty<Fill>();
            _numberFormats = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                .Where(item => item.NumberFormatId?.Value != null && item.FormatCode?.Value != null)
                .ToDictionary(item => item.NumberFormatId!.Value, item => item.FormatCode!.Value!, EqualityComparer<uint>.Default)
                ?? new Dictionary<uint, string>();
        }

        internal ExcelEffectiveCellStyle Resolve(Cell cell, Row row, Column? column) {
            uint styleIndex = cell.StyleIndex?.Value
                ?? (row.CustomFormat?.Value == true ? row.StyleIndex?.Value : null)
                ?? column?.Style?.Value
                ?? 0U;
            CellFormat? format = styleIndex < _formats.Length ? _formats[styleIndex] : null;
            Font? font = format?.FontId?.Value is uint fontId && fontId < _fonts.Length ? _fonts[fontId] : null;
            Fill? fill = format?.FillId?.Value is uint fillId && fillId < _fills.Length ? _fills[fillId] : null;
            string? foregroundArgb = ExcelThemeColorResolver.Resolve(font?.Color, _workbookPart);
            PatternFill? pattern = fill?.PatternFill;
            string? backgroundArgb = pattern?.PatternType?.Value == PatternValues.Solid
                ? ExcelThemeColorResolver.Resolve(pattern.ForegroundColor, _workbookPart)
                : null;
            bool transparent = foregroundArgb?.StartsWith("00", StringComparison.OrdinalIgnoreCase) == true;
            OfficeColor background = OfficeColor.White;
            double? contrast = TryParseExcelColor(foregroundArgb, OfficeColor.Black, out OfficeColor foreground) &&
                TryParseExcelColor(backgroundArgb, OfficeColor.White, out background)
                ? OfficeColorContrast.ContrastRatio(foreground, background)
                : null;
            string? numberFormat = format?.NumberFormatId?.Value is uint numberFormatId && _numberFormats.TryGetValue(numberFormatId, out string? custom)
                ? custom
                : null;
            return new ExcelEffectiveCellStyle(
                font?.FontSize?.Val?.Value,
                transparent,
                contrast,
                string.Equals(RemoveAsciiWhitespace(numberFormat), ";;;", StringComparison.Ordinal),
                background);
        }

        internal static bool TryParseExcelColor(string? value, OfficeColor fallback, out OfficeColor color) {
            if (string.IsNullOrWhiteSpace(value)) { color = fallback; return true; }
            string rgb = value!.Length == 8 ? value.Substring(2) : value;
            if (rgb.Length == 6 && OfficeColor.TryParseHex(rgb, out color)) return true;
            color = default;
            return false;
        }

        private static string RemoveAsciiWhitespace(string? value) => string.IsNullOrEmpty(value)
            ? string.Empty
            : new string(value!.Where(character => character is not ' ' and not '\t' and not '\r' and not '\n').ToArray());
    }

    private sealed class ExcelEffectiveCellStyle {
        internal ExcelEffectiveCellStyle(double? fontSizePoints, bool transparentText, double? contrastRatio, bool hiddenDisplayValue, OfficeColor? background) {
            FontSizePoints = fontSizePoints;
            TransparentText = transparentText;
            ContrastRatio = contrastRatio;
            HiddenDisplayValue = hiddenDisplayValue;
            Background = background;
        }
        internal double? FontSizePoints { get; }
        internal bool TransparentText { get; }
        internal double? ContrastRatio { get; }
        internal bool HiddenDisplayValue { get; }
        internal OfficeColor? Background { get; }
    }

    private sealed class ExcelCleanupTarget : IEquatable<ExcelCleanupTarget> {
        private readonly OpenXmlElement _element;
        private readonly ExcelCleanupOperation _operation;
        private readonly string? _attribute;
        private readonly Cell? _cell;
        private readonly OpenXmlCompositeElement? _richOwner;
        private readonly int? _offset;
        private readonly int? _length;
        private readonly string? _expected;
        private readonly int _sequence;
        private ExcelCleanupTarget(OpenXmlElement element, ExcelCleanupOperation operation, string? attribute = null, Cell? cell = null, OpenXmlCompositeElement? richOwner = null, int? offset = null, int? length = null, string? expected = null, int sequence = 0) {
            _element = element;
            _operation = operation;
            _attribute = attribute;
            _cell = cell;
            _richOwner = richOwner;
            _offset = offset;
            _length = length;
            _expected = expected;
            _sequence = sequence;
        }
        internal static ExcelCleanupTarget ForCellPayload(Cell cell) => new ExcelCleanupTarget(cell, ExcelCleanupOperation.CellPayload);
        internal static ExcelCleanupTarget ForElement(OpenXmlElement element) => new ExcelCleanupTarget(element, ExcelCleanupOperation.Element);
        internal static ExcelCleanupTarget ForAttribute(OpenXmlElement element, string attribute) => new ExcelCleanupTarget(element, ExcelCleanupOperation.Attribute, attribute);
        internal static ExcelCleanupTarget ForRichRun(Cell cell, OpenXmlCompositeElement owner, Run run) => new ExcelCleanupTarget(
            run, ExcelCleanupOperation.RichRun, cell: cell, richOwner: owner,
            sequence: owner.Elements<Run>().TakeWhile(item => !ReferenceEquals(item, run)).Count());
        internal static ExcelCleanupTarget ForTextRange(OpenXmlLeafTextElement text, OfficeContentSafetyFinding finding) => new ExcelCleanupTarget(
            text, ExcelCleanupOperation.TextRange, offset: finding.SourceTextOffset, length: finding.SourceTextLength,
            expected: text.Text.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value), sequence: finding.SourceTextOffset.Value);
        internal static ExcelCleanupTarget ForSharedTextRange(Cell cell, OpenXmlCompositeElement owner, Text text, OfficeContentSafetyFinding finding) => new ExcelCleanupTarget(
            text, ExcelCleanupOperation.SharedTextRange, cell: cell, richOwner: owner, offset: finding.SourceTextOffset, length: finding.SourceTextLength,
            expected: text.Text.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value),
            sequence: checked(owner.Descendants<Text>().TakeWhile(item => !ReferenceEquals(item, text)).Count() * 1_000_000 + finding.SourceTextOffset.Value));
        internal int Sequence => _sequence;
        internal int RemovalPriority => _operation is ExcelCleanupOperation.TextRange or ExcelCleanupOperation.SharedTextRange ? 2 :
            _operation == ExcelCleanupOperation.RichRun ? 1 : 0;
        internal void Remove() {
            if (_operation == ExcelCleanupOperation.TextRange && _element is OpenXmlLeafTextElement text && _offset.HasValue && _length.HasValue) {
                text.Text = RemoveVerifiedRange(text.Text, _offset.Value, _length.Value, _expected);
                return;
            }
            if (_operation == ExcelCleanupOperation.SharedTextRange && _element is Text sourceText && _cell != null && _richOwner != null && _offset.HasValue && _length.HasValue) {
                int textIndex = _richOwner.Descendants<Text>().TakeWhile(item => !ReferenceEquals(item, sourceText)).Count();
                InlineString inline;
                if (_cell.InlineString != null && _cell.DataType?.Value == CellValues.InlineString) {
                    inline = _cell.InlineString;
                } else {
                    inline = new InlineString();
                    foreach (OpenXmlElement child in _richOwner.ChildElements) inline.Append(child.CloneNode(true));
                }
                Text clonedText = inline.Descendants<Text>().ElementAt(textIndex);
                clonedText.Text = RemoveVerifiedRange(clonedText.Text, _offset.Value, _length.Value, _expected);
                _cell.InlineString = inline;
                _cell.CellValue = null;
                _cell.DataType = CellValues.InlineString;
                return;
            }
            if (_operation == ExcelCleanupOperation.Element) { _element.Remove(); return; }
            if (_operation == ExcelCleanupOperation.CellPayload && _element is Cell cell) {
                cell.CellFormula = null;
                cell.CellValue = null;
                cell.InlineString = null;
                cell.DataType = null;
                return;
            }
            if (_operation == ExcelCleanupOperation.RichRun && _element is Run richRun && _cell != null && _richOwner != null) {
                if (_richOwner is InlineString) {
                    richRun.Remove();
                } else {
                    InlineString inline;
                    if (_cell.InlineString != null && _cell.DataType?.Value == CellValues.InlineString) {
                        inline = _cell.InlineString;
                    } else {
                        inline = new InlineString();
                        foreach (OpenXmlElement child in _richOwner.ChildElements) inline.Append(child.CloneNode(true));
                    }
                    int runIndex = _richOwner.Elements<Run>().TakeWhile(item => !ReferenceEquals(item, richRun)).Count();
                    inline.Elements<Run>().ElementAt(runIndex).Remove();
                    _cell.InlineString = inline;
                    _cell.CellValue = null;
                    _cell.DataType = CellValues.InlineString;
                }
                return;
            }
            if (_operation == ExcelCleanupOperation.Attribute && _element is Xdr.NonVisualDrawingProperties properties) {
                if (string.Equals(_attribute, "descr", StringComparison.Ordinal)) properties.Description = null;
                else if (string.Equals(_attribute, "title", StringComparison.Ordinal)) properties.Title = null;
            }
        }
        private static string RemoveVerifiedRange(string current, int offset, int length, string? expected) {
            if (offset > current.Length - length || !string.Equals(current.Substring(offset, length), expected, StringComparison.Ordinal)) {
                throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected Excel text node.");
            }
            return current.Remove(offset, length);
        }
        public bool Equals(ExcelCleanupTarget? other) => other != null && ReferenceEquals(_element, other._element) && ReferenceEquals(_cell, other._cell) && ReferenceEquals(_richOwner, other._richOwner) && _operation == other._operation && string.Equals(_attribute, other._attribute, StringComparison.Ordinal) && _offset == other._offset && _length == other._length;
        public override bool Equals(object? obj) => Equals(obj as ExcelCleanupTarget);
        public override int GetHashCode() { unchecked { return (_element.GetHashCode() * 397) ^ ((_cell?.GetHashCode() ?? 0) * 131) ^ ((_richOwner?.GetHashCode() ?? 0) * 17) ^ ((int)_operation * 31) ^ (_attribute?.GetHashCode() ?? 0) ^ (_offset ?? 0); } }
        private enum ExcelCleanupOperation { Element, CellPayload, Attribute, RichRun, TextRange, SharedTextRange }
    }
}
