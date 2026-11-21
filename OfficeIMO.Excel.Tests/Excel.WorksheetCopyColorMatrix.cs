using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Utilities;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [MemberData(nameof(FormulaMatrixRoutes))]
        public void Test_WorksheetCopy_ColorMatrixKeepsSourceAppearance(bool merge, ExcelWorksheetCopyMode mode) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"ColorMatrixSource-{merge}-{mode}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"ColorMatrixTarget-{merge}-{mode}.xlsx");
            CreateColorMatrixSource(sourcePath);
            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            // The target gets its own theme accents and palette, so every theme and indexed color resolves differently.
            RecolorWorkbook(targetPath, "1A2B3C", "FF0F0F0F");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                if (merge) {
                    targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        CopyMode = mode,
                        TargetSheetNames = new Dictionary<string, string> { ["Data"] = "Copy" }
                    });
                } else {
                    targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = mode });
                }

                targetDocument.Save();
            }

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, false))
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                string[] sourceColors = ReadColorMatrix(source, "Data");
                string[] targetColors = ReadColorMatrix(target, "Copy");
                Assert.Equal(sourceColors, targetColors);
                Assert.Contains("auto", targetColors);
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        /// <summary>
        /// Resolves, in its own workbook, every color a copy carries: the column style's font, fill and border colors
        /// (theme, indexed, rgb and auto), the conditional-format dxf color, the color-scale and data-bar colors and the
        /// tab color.
        /// </summary>
        private static string[] ReadColorMatrix(SpreadsheetDocument document, string sheetName) {
            WorkbookPart workbookPart = document.WorkbookPart!;
            Stylesheet styles = workbookPart.WorkbookStylesPart!.Stylesheet!;
            Worksheet worksheet = GetWorksheetPartByNameForOperations(document, sheetName).Worksheet!;
            string Resolve(ColorType? color) =>
                color == null ? "none"
                : color.Auto?.Value == true ? "auto"
                : ExcelThemeColorResolver.Resolve(color, workbookPart) ?? "unresolved";

            CellFormat columnFormat = styles.CellFormats!.Elements<CellFormat>()
                .ElementAt((int)worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single().Style!.Value);
            Font font = styles.Fonts!.Elements<Font>().ElementAt((int)columnFormat.FontId!.Value);
            Fill fill = styles.Fills!.Elements<Fill>().ElementAt((int)columnFormat.FillId!.Value);
            Border border = styles.Borders!.Elements<Border>().ElementAt((int)columnFormat.BorderId!.Value);
            ConditionalFormattingRule expressionRule = worksheet.Descendants<ConditionalFormattingRule>()
                .Single(rule => rule.Type!.Value == ConditionalFormatValues.Expression);
            DifferentialFormat dxf = styles.DifferentialFormats!.Elements<DifferentialFormat>().ElementAt((int)expressionRule.FormatId!.Value);

            var colors = new List<string> {
                Resolve(font.Color),
                Resolve(fill.PatternFill!.ForegroundColor),
                Resolve(border.LeftBorder!.Color),
                Resolve(border.BottomBorder!.Color),
                Resolve(dxf.Font!.Color),
                Resolve(worksheet.GetFirstChild<SheetProperties>()!.TabColor)
            };
            colors.AddRange(worksheet.Descendants<ColorScale>().Single().Elements<Color>().Select(Resolve));
            colors.Add(Resolve(worksheet.Descendants<DataBar>().Single().GetFirstChild<Color>()));
            return colors.ToArray();
        }

        private static void CreateColorMatrixSource(string path) {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet data = document.AddSheet("Data");
                for (int row = 1; row <= 3; row++) data.Cell(row, 1).SetValue(row);
                document.Save();
            }

            RecolorWorkbook(path, "C04080", "FFAA5500");
            using SpreadsheetDocument source = SpreadsheetDocument.Open(path, true);
            Stylesheet styles = source.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            styles.Fonts!.Append(new Font(new FontSize { Val = 11D }, new Color { Theme = 4U, Tint = 0.4D }, new FontName { Val = "Calibri" }));
            styles.Fonts.Count = (uint)styles.Fonts.ChildElements.Count;
            styles.Fills!.Append(new Fill(new PatternFill(new ForegroundColor { Indexed = 10U }) { PatternType = PatternValues.Solid }));
            styles.Fills.Count = (uint)styles.Fills.ChildElements.Count;
            styles.Borders!.Append(new Border(
                new LeftBorder(new Color { Rgb = "FF112233" }) { Style = BorderStyleValues.Thin },
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder()));
            styles.Borders.Count = (uint)styles.Borders.ChildElements.Count;
            styles.CellFormats!.Append(new CellFormat {
                NumberFormatId = 0U,
                FontId = styles.Fonts.Count.Value - 1U,
                FillId = styles.Fills.Count.Value - 1U,
                BorderId = styles.Borders.Count.Value - 1U,
                FormatId = 0U,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true
            });
            styles.CellFormats.Count = (uint)styles.CellFormats.ChildElements.Count;
            styles.DifferentialFormats ??= new DifferentialFormats();
            styles.DifferentialFormats.Append(new DifferentialFormat(new Font(new Color { Theme = 5U })));
            styles.DifferentialFormats.Count = (uint)styles.DifferentialFormats.ChildElements.Count;
            styles.Save();

            Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!;
            worksheet.GetFirstChild<SheetProperties>()?.Remove();
            worksheet.InsertAt(new SheetProperties(new TabColor { Theme = 8U }), 0);
            worksheet.GetFirstChild<Columns>()?.Remove();
            worksheet.InsertBefore(new Columns(new Column { Min = 1U, Max = 1U, Width = 12D, CustomWidth = true, Style = styles.CellFormats.Count.Value - 1U }),
                worksheet.GetFirstChild<SheetData>()!);
            worksheet.InsertAfter(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("A1>1")) { Type = ConditionalFormatValues.Expression, Priority = 1, FormatId = styles.DifferentialFormats.Count.Value - 1U },
                new ConditionalFormattingRule(new ColorScale(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                    new Color { Theme = 6U },
                    new Color { Indexed = 12U })) { Type = ConditionalFormatValues.ColorScale, Priority = 2 },
                new ConditionalFormattingRule(new DataBar(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                    new Color { Theme = 7U, Tint = -0.25D })) { Type = ConditionalFormatValues.DataBar, Priority = 3 }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3" }
            }, worksheet.GetFirstChild<SheetData>()!);
            worksheet.Save();
            Assert.Empty(new OpenXmlValidator().Validate(source));
        }

        /// <summary>Gives a workbook distinctive theme accents and a custom indexed palette.</summary>
        private static void RecolorWorkbook(string path, string accentPrefix, string paletteColor) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            ThemePart themePart = workbookPart.GetPartsOfType<ThemePart>().Single();
            A.ColorScheme scheme = themePart.Theme!.ThemeElements!.ColorScheme!;
            int accent = 0;
            foreach (OpenXmlCompositeElement slot in scheme.Elements<OpenXmlCompositeElement>()
                .Where(element => element.LocalName.StartsWith("accent", StringComparison.Ordinal))) {
                slot.RemoveAllChildren();
                slot.Append(new A.RgbColorModelHex { Val = accentPrefix.Substring(0, 4) + (10 + accent++).ToString(System.Globalization.CultureInfo.InvariantCulture) });
            }

            themePart.Theme.Save();
            Stylesheet styles = workbookPart.WorkbookStylesPart!.Stylesheet!;
            var palette = new IndexedColors();
            for (int index = 0; index < 64; index++) {
                palette.Append(new RgbColor { Rgb = index is 10 or 12 ? paletteColor : "FF" + (index * 4).ToString("X2", System.Globalization.CultureInfo.InvariantCulture) + "0000" });
            }

            styles.Colors = new Colors(palette);
            styles.Save();
        }
    }
}
