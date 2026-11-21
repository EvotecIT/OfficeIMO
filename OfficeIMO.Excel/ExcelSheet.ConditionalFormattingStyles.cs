using System.Runtime.CompilerServices;
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
            return GetOrAddConditionalDifferentialFormat(stylesPart, stylesheet, candidate);
        }

        private static uint GetOrAddConditionalDifferentialFormat(WorkbookStylesPart stylesPart, Stylesheet stylesheet,
            DifferentialFormat candidate) {
            DifferentialFormats formats = stylesheet.DifferentialFormats ??= new DifferentialFormats();
            return GetOrAppendDifferentialFormat(stylesPart, formats, candidate);
        }

        // Each rule looked up its differential format by serializing the candidate once per existing format, which
        // made adding thousands of rules spend most of its time in XML writers. Formats are only appended by the
        // library, so their serialized forms are kept per formats element while its children keep their count and
        // last element; a package handed out for direct editing is always searched.
        private static readonly ConditionalWeakTable<DifferentialFormats, DifferentialFormatIndex> DifferentialFormatIndexes = new();

        private sealed class DifferentialFormatIndex {
            internal readonly Dictionary<string, uint> Indexes = new(StringComparer.Ordinal);
            internal int Count;
            internal OpenXmlElement? LastChild;
        }

        /// <summary>Returns the index of a format equal to <paramref name="candidate"/>, appending it when there is none.</summary>
        private static uint GetOrAppendDifferentialFormat(WorkbookStylesPart stylesPart, DifferentialFormats formats, DifferentialFormat candidate) {
            string key = candidate.OuterXml;
            DifferentialFormatIndex? index = null;
            if (!IsPackageExposed(stylesPart.OpenXmlPackage)) {
                index = DifferentialFormatIndexes.GetValue(formats, static _ => new DifferentialFormatIndex());
                int count = 0;
                for (OpenXmlElement? child = formats.FirstChild; child != null; child = child.NextSibling()) count++;
                if (index.Count != count || !ReferenceEquals(index.LastChild, formats.LastChild)) {
                    index.Indexes.Clear();
                    uint position = 0;
                    foreach (DifferentialFormat format in formats.Elements<DifferentialFormat>()) {
                        string existing = format.OuterXml;
                        if (!index.Indexes.ContainsKey(existing)) index.Indexes.Add(existing, position);
                        position++;
                    }
                    index.Count = count;
                    index.LastChild = formats.LastChild;
                }
                if (index.Indexes.TryGetValue(key, out uint found)) return found;
            } else {
                uint position = 0;
                foreach (DifferentialFormat format in formats.Elements<DifferentialFormat>()) {
                    if (string.Equals(format.OuterXml, key, StringComparison.Ordinal)) return position;
                    position++;
                }
            }

            formats.Append(candidate);
            formats.Count = (uint)formats.Elements<DifferentialFormat>().Count();
            uint appended = formats.Count!.Value - 1U;
            if (index != null) {
                if (!index.Indexes.ContainsKey(key)) index.Indexes.Add(key, appended);
                index.Count++;
                index.LastChild = candidate;
            }
            SaveStylesheet(stylesPart);
            return appended;
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
                if (definition.DifferentialFontBold.HasValue) font.AddChild(new Bold { Val = definition.DifferentialFontBold.Value }, true);
                if (definition.DifferentialFontItalic.HasValue) font.AddChild(new Italic { Val = definition.DifferentialFontItalic.Value }, true);
                if (definition.DifferentialFontStrike.HasValue) font.AddChild(new Strike { Val = definition.DifferentialFontStrike.Value }, true);
                if (definition.DifferentialFontUnderline.HasValue) font.AddChild(new Underline {
                    Val = definition.DifferentialFontUnderline.Value ? UnderlineValues.Single : UnderlineValues.None
                }, true);
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFontColorArgb)) {
                    font.AddChild(new DocumentFormat.OpenXml.Spreadsheet.Color {
                        Rgb = NormalizeHexColor(definition.DifferentialFontColorArgb!)
                    }, true);
                }
                if (definition.DifferentialFontSize.HasValue) {
                    font.AddChild(new FontSize { Val = definition.DifferentialFontSize.Value.ToPoints() }, true);
                }
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFontName)) {
                    font.AddChild(new FontName { Val = definition.DifferentialFontName }, true);
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
            ExcelCellBorderInfo snapshot) {
            border.LeftBorder = CreateProjectedConditionalFormattingBorderSide<LeftBorder>(snapshot.Left);
            border.RightBorder = CreateProjectedConditionalFormattingBorderSide<RightBorder>(snapshot.Right);
            border.TopBorder = CreateProjectedConditionalFormattingBorderSide<TopBorder>(snapshot.Top);
            border.BottomBorder = CreateProjectedConditionalFormattingBorderSide<BottomBorder>(snapshot.Bottom);
            border.DiagonalBorder = CreateProjectedConditionalFormattingBorderSide<DiagonalBorder>(snapshot.Diagonal);
            border.DiagonalUp = snapshot.Diagonal != null && snapshot.DiagonalUp;
            border.DiagonalDown = snapshot.Diagonal != null && snapshot.DiagonalDown;
        }

        private static TBorder? CreateProjectedConditionalFormattingBorderSide<TBorder>(
            ExcelBorderSideInfo? snapshot)
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
            definition.DifferentialFontStrike.HasValue ||
            definition.DifferentialFontUnderline.HasValue ||
            !string.IsNullOrWhiteSpace(definition.DifferentialFontName) ||
            definition.DifferentialFontSize.HasValue;
    }
}
