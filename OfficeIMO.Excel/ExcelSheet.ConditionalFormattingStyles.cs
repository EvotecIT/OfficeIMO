using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private uint? CreateOrReuseStandardConditionalDifferentialFormat(
            ExcelConditionalFormattingInfo definition) {
            if (!HasProjectedConditionalFormattingStyle(definition) && !definition.DifferentialFormatId.HasValue) return null;

            WorkbookStylesPart stylesPart = _excelDocument.WorkbookPartRoot.WorkbookStylesPart
                ?? _excelDocument.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);
            DifferentialFormat candidate = definition.DifferentialFormatId.HasValue
                ? GetDifferentialFormat(stylesheet, definition.DifferentialFormatId)?.CloneNode(true) as DifferentialFormat
                    ?? new DifferentialFormat()
                : new DifferentialFormat();
            ApplyProjectedConditionalFormattingStyle(candidate, definition);
            if (!candidate.ChildElements.Any() && !candidate.ExtendedAttributes.Any()) return null;
            DifferentialFormats formats = stylesheet.DifferentialFormats ??= new DifferentialFormats();
            int existingIndex = formats.Elements<DifferentialFormat>()
                .Select((format, index) => new { format, index })
                .Where(item => string.Equals(item.format.OuterXml, candidate.OuterXml, StringComparison.Ordinal))
                .Select(item => item.index)
                .DefaultIfEmpty(-1)
                .First();
            if (existingIndex >= 0) return (uint)existingIndex;

            formats.Append(candidate);
            formats.Count = (uint)formats.Elements<DifferentialFormat>().Count();
            stylesheet.Save();
            return formats.Count!.Value - 1U;
        }

        private static void ApplyOffice2010ConditionalDifferentialFormat(
            X14.ConditionalFormattingRule rule,
            ExcelConditionalFormattingInfo definition) {
            X14.DifferentialType? differential = rule.GetFirstChild<X14.DifferentialType>();
            if (!HasProjectedConditionalFormattingStyle(definition)) {
                if (differential == null) return;
                differential.Fill?.Remove();
                differential.Font?.Remove();
                differential.Border?.Remove();
                if (!differential.ChildElements.Any() && !differential.ExtendedAttributes.Any()) differential.Remove();
                return;
            }

            if (differential == null) {
                differential = new X14.DifferentialType();
                InsertBeforeRuleExtension(rule, differential);
            }
            ApplyProjectedConditionalFormattingStyle(differential, definition);
        }

        private static void ApplyProjectedConditionalFormattingStyle(
            OpenXmlCompositeElement differential,
            ExcelConditionalFormattingInfo definition) {
            differential.GetFirstChild<Fill>()?.Remove();
            differential.GetFirstChild<Font>()?.Remove();
            if (definition.DifferentialBorder == null) differential.GetFirstChild<Border>()?.Remove();
            if (!string.IsNullOrWhiteSpace(definition.DifferentialFillColorArgb)) {
                string fill = NormalizeHexColor(definition.DifferentialFillColorArgb!);
                InsertDifferentialStyleChild(differential, new Fill(new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = fill },
                    BackgroundColor = new BackgroundColor { Rgb = fill }
                }));
            }

            if (HasProjectedConditionalFormattingFont(definition)) {
                var font = new Font();
                if (definition.DifferentialFontBold == true) font.Append(new Bold());
                if (definition.DifferentialFontItalic == true) font.Append(new Italic());
                if (definition.DifferentialFontUnderline == true) font.Append(new Underline());
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFontColorArgb)) {
                    font.Append(new DocumentFormat.OpenXml.Spreadsheet.Color {
                        Rgb = NormalizeHexColor(definition.DifferentialFontColorArgb!)
                    });
                }
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFontName)) {
                    font.Append(new FontName { Val = definition.DifferentialFontName });
                }
                if (definition.DifferentialFontSize.HasValue) {
                    font.Append(new FontSize { Val = definition.DifferentialFontSize.Value });
                }
                if (font.ChildElements.Any()) InsertDifferentialStyleChild(differential, font);
            }

            if (definition.DifferentialBorder != null) {
                Border border = differential.GetFirstChild<Border>() ?? new Border();
                ApplyProjectedConditionalFormattingBorder(border, definition.DifferentialBorder);
                if (border.Parent == null) InsertDifferentialStyleChild(differential, border);
            }
        }

        private static void ApplyProjectedConditionalFormattingBorder(
            Border border,
            ExcelCellBorderSnapshot snapshot) {
            border.LeftBorder = CreateProjectedConditionalFormattingBorderSide<LeftBorder>(snapshot.Left);
            border.RightBorder = CreateProjectedConditionalFormattingBorderSide<RightBorder>(snapshot.Right);
            border.TopBorder = CreateProjectedConditionalFormattingBorderSide<TopBorder>(snapshot.Top);
            border.BottomBorder = CreateProjectedConditionalFormattingBorderSide<BottomBorder>(snapshot.Bottom);
            border.DiagonalBorder = CreateProjectedConditionalFormattingBorderSide<DiagonalBorder>(snapshot.Diagonal);
            border.DiagonalUp = snapshot.Diagonal != null && snapshot.DiagonalUp;
            border.DiagonalDown = snapshot.Diagonal != null && snapshot.DiagonalDown;
        }

        private static TBorder? CreateProjectedConditionalFormattingBorderSide<TBorder>(
            ExcelBorderSideSnapshot? snapshot)
            where TBorder : BorderPropertiesType, new() {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.Style)) return null;
            var side = new TBorder();
            SetRuleAttribute(side, "style", NormalizeConditionalFormattingBorderStyle(snapshot.Style));
            if (!string.IsNullOrWhiteSpace(snapshot.ColorArgb)) {
                side.Append(new DocumentFormat.OpenXml.Spreadsheet.Color {
                    Rgb = NormalizeHexColor(snapshot.ColorArgb!)
                });
            }
            return side;
        }

        private static void InsertDifferentialStyleChild(OpenXmlCompositeElement differential, OpenXmlElement child) {
            OpenXmlElement? before = child is Font
                ? differential.ChildElements.FirstOrDefault(element =>
                    element is NumberingFormat || element is Fill || element is Border ||
                    element is Alignment || element is Protection ||
                    element is DocumentFormat.OpenXml.Spreadsheet.ExtensionList)
                : differential.ChildElements.FirstOrDefault(element =>
                    element is Border || element is Alignment || element is Protection ||
                    element is DocumentFormat.OpenXml.Spreadsheet.ExtensionList);
            if (before == null) differential.Append(child);
            else differential.InsertBefore(child, before);
        }

        private static bool HasProjectedConditionalFormattingStyle(ExcelConditionalFormattingInfo definition) =>
            !string.IsNullOrWhiteSpace(definition.DifferentialFillColorArgb) ||
            HasProjectedConditionalFormattingFont(definition) ||
            definition.DifferentialBorder != null;

        private static bool HasProjectedConditionalFormattingFont(ExcelConditionalFormattingInfo definition) =>
            !string.IsNullOrWhiteSpace(definition.DifferentialFontColorArgb) ||
            definition.DifferentialFontBold.HasValue ||
            definition.DifferentialFontItalic.HasValue ||
            definition.DifferentialFontUnderline.HasValue ||
            !string.IsNullOrWhiteSpace(definition.DifferentialFontName) ||
            definition.DifferentialFontSize.HasValue;
    }
}
