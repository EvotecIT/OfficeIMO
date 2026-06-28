using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesInlineRichTextCells() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Rich Cells");
                    sheet.CellAt(1, 1).SetRichText(
                        new ExcelRichTextRun("Inline ") {
                            Bold = true,
                            FontColor = "#123456",
                            FontName = "Consolas",
                            FontSize = 13D
                        },
                        new ExcelRichTextRun("cell") {
                            Italic = true,
                            Underline = true,
                            FontName = "Arial",
                            FontSize = 11D
                        });

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell legacyCell = Assert.Single(legacySheet.Cells);
                Assert.Equal("Inline cell", legacyCell.Value);
                Assert.Equal(2, legacyCell.TextFormattingRuns.Count);

                IReadOnlyList<ExcelRichTextRun> projectedRuns = result.Document.Sheets[0].CellAt(1, 1).GetRichText();
                Assert.Equal(2, projectedRuns.Count);
                AssertRichTextRun(projectedRuns[0], "Inline ", bold: true, italic: false, underline: false, "FF123456", "Consolas", 13D);
                AssertRichTextRun(projectedRuns[1], "cell", bold: false, italic: true, underline: true, null, "Arial", 11D);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesRichTextCellFontFamilyAndCharset() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Rich Font Bytes");
                    sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Run font bytes") {
                        FontName = "Arial",
                        FontFamily = 2,
                        FontCharacterSet = 238
                    });

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell legacyCell = Assert.Single(legacySheet.Cells);
                LegacyXlsTextFormattingRun formattingRun = Assert.Single(legacyCell.TextFormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal((byte)2, font.Family);
                Assert.Equal((byte)238, font.CharacterSet);

                ExcelRichTextRun projectedRun = Assert.Single(result.Document.Sheets[0].CellAt(1, 1).GetRichText());
                Assert.Equal((byte)2, projectedRun.FontFamily);
                Assert.Equal((byte)238, projectedRun.FontCharacterSet);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesRichTextCellVerticalTextAlignment() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Rich Escapement");
                    sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Raised") {
                        FontName = "Arial",
                        VerticalTextAlignment = VerticalAlignmentRunValues.Superscript
                    });

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell legacyCell = Assert.Single(legacySheet.Cells);
                LegacyXlsTextFormattingRun formattingRun = Assert.Single(legacyCell.TextFormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal(LegacyXlsFontEscapement.Superscript, font.Escapement);

                ExcelRichTextRun projectedRun = Assert.Single(result.Document.Sheets[0].CellAt(1, 1).GetRichText());
                Assert.Equal(VerticalAlignmentRunValues.Superscript, projectedRun.VerticalTextAlignment);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesRichTextCellFontOptionFlags() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Rich Font Flags");
                    sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Run flags") {
                        FontName = "Arial",
                        Outline = true,
                        Shadow = true,
                        Condense = true,
                        Extend = true
                    });

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell legacyCell = Assert.Single(legacySheet.Cells);
                LegacyXlsTextFormattingRun formattingRun = Assert.Single(legacyCell.TextFormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.True(font.Outline);
                Assert.True(font.Shadow);
                Assert.True(font.Condense);
                Assert.True(font.Extend);

                ExcelRichTextRun projectedRun = Assert.Single(result.Document.Sheets[0].CellAt(1, 1).GetRichText());
                Assert.True(projectedRun.Outline);
                Assert.True(projectedRun.Shadow);
                Assert.True(projectedRun.Condense);
                Assert.True(projectedRun.Extend);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesRichTextCellUnderlineStyle() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Rich Underline");
                    sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Accounting") {
                        FontName = "Arial",
                        UnderlineStyle = UnderlineValues.SingleAccounting
                    });

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell legacyCell = Assert.Single(legacySheet.Cells);
                LegacyXlsTextFormattingRun formattingRun = Assert.Single(legacyCell.TextFormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal((byte)0x21, font.UnderlineStyle);

                ExcelRichTextRun projectedRun = Assert.Single(result.Document.Sheets[0].CellAt(1, 1).GetRichText());
                Assert.Equal(UnderlineValues.SingleAccounting, projectedRun.UnderlineStyle);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSharedStringRichTextCells() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Shared Rich Cells");
                    sheet.CellValue(1, 1, "Shared rich text");

                    SharedStringTablePart sharedStringPart = document.WorkbookPartRoot.SharedStringTablePart
                        ?? document.WorkbookPartRoot.AddNewPart<SharedStringTablePart>();
                    sharedStringPart.SharedStringTable = new SharedStringTable(
                        new SharedStringItem(
                            new Run(
                                new RunProperties(
                                    new Italic(),
                                    new Color { Rgb = "FF654321" },
                                    new RunFont { Val = "Calibri" },
                                    new FontSize { Val = 12D }),
                                new Text("Shared ")),
                            new Run(
                                new RunProperties(new Bold()),
                                new Text("cell"))));
                    sharedStringPart.SharedStringTable.Save();

                    Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                        .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                    cell.DataType = CellValues.SharedString;
                    cell.RemoveAllChildren<CellValue>();
                    cell.RemoveAllChildren<InlineString>();
                    cell.Append(new CellValue("0"));
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell legacyCell = Assert.Single(legacySheet.Cells);
                Assert.Equal("Shared cell", legacyCell.Value);
                Assert.Equal(2, legacyCell.TextFormattingRuns.Count);

                IReadOnlyList<ExcelRichTextRun> projectedRuns = result.Document.Sheets[0].CellAt(1, 1).GetRichText();
                Assert.Equal(2, projectedRuns.Count);
                AssertRichTextRun(projectedRuns[0], "Shared ", bold: false, italic: true, underline: false, "FF654321", "Calibri", 12D);
                AssertRichTextRun(projectedRuns[1], "cell", bold: true, italic: false, underline: false, null, "Calibri", 11D);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateRichTextFontPropertiesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("rich-text font properties with duplicate font name elements", (document, sheet) => {
                sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Duplicate font") {
                    FontName = "Arial"
                });

                RunProperties properties = sheet.WorksheetPart.Worksheet!
                    .Descendants<RunProperties>()
                    .Single();
                properties.Append(new RunFont { Val = "Calibri" });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedRichTextRunMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("rich-text cell run metadata", (document, sheet) => {
                sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Run metadata") {
                    FontName = "Arial"
                });

                Run run = sheet.WorksheetPart.Worksheet!
                    .Descendants<Run>()
                    .Single();
                run.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedRichTextRunPropertyMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("rich-text font properties with unsupported metadata", (document, sheet) => {
                sheet.CellAt(1, 1).SetRichText(new ExcelRichTextRun("Run property metadata") {
                    FontName = "Arial"
                });

                RunProperties properties = sheet.WorksheetPart.Worksheet!
                    .Descendants<RunProperties>()
                    .Single();
                properties.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        private static void AssertRichTextRun(
            ExcelRichTextRun run,
            string text,
            bool bold,
            bool italic,
            bool underline,
            string? fontColor,
            string? fontName,
            double? fontSize) {
            Assert.Equal(text, run.Text);
            Assert.Equal(bold, run.Bold);
            Assert.Equal(italic, run.Italic);
            Assert.Equal(underline, run.Underline);
            Assert.Equal(fontColor, run.FontColor);
            Assert.Equal(fontName, run.FontName);
            Assert.Equal(fontSize, run.FontSize);
        }
    }
}
