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
        public void LegacyXls_NativeSave_WritesVisibleWorksheetCommentAnchors() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("CommentAnchor");
                    sheet.CellValue(3, 2, "Anchored note");
                    sheet.SetLegacyComment(
                        3,
                        2,
                        "Visible anchored note",
                        "Reviewer",
                        visible: true,
                        new ExcelCommentAnchor(1, 10, 2, 20, 3, 30, 4, 40));

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                Assert.Equal(3, legacyComment.Row);
                Assert.Equal(2, legacyComment.Column);
                Assert.Equal("Visible anchored note", legacyComment.Text);
                Assert.True(legacyComment.Visible);

                LegacyXlsDrawingAnchor anchor = legacyComment.Anchor!;
                Assert.Equal((ushort)0, anchor.Flags);
                Assert.Equal((ushort)1, anchor.StartColumn);
                Assert.Equal((ushort)10, anchor.StartDx);
                Assert.Equal((ushort)2, anchor.StartRow);
                Assert.Equal((ushort)20, anchor.StartDy);
                Assert.Equal((ushort)3, anchor.EndColumn);
                Assert.Equal((ushort)30, anchor.EndDx);
                Assert.Equal((ushort)4, anchor.EndRow);
                Assert.Equal((ushort)40, anchor.EndDy);

                VmlDrawingPart vmlPart = Assert.Single(result.Document.Sheets[0].WorksheetPart.VmlDrawingParts);
                using var reader = new StreamReader(vmlPart.GetStream());
                string vml = reader.ReadToEnd();
                Assert.Contains("<x:Anchor>1, 10, 2, 20, 3, 30, 4, 40</x:Anchor>", vml, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("visibility:visible", vml, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("<x:Visible", vml, StringComparison.OrdinalIgnoreCase);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCommentRichTextFontFamilyAndCharset() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("CommentFonts");
                    sheet.CellValue(1, 1, "Comment");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Comment font bytes") {
                                FontName = "Arial",
                                FontFamily = 2,
                                FontCharacterSet = 238
                            }
                        },
                        "Reviewer");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                LegacyXlsCommentFormattingRun formattingRun = Assert.Single(legacyComment.FormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal((byte)2, font.Family);
                Assert.Equal((byte)238, font.CharacterSet);

                ExcelCommentInfo projectedComment = Assert.Single(result.Document.Sheets[0].GetComments());
                ExcelRichTextRun projectedRun = Assert.Single(projectedComment.RichTextRuns);
                Assert.Equal((byte)2, projectedRun.FontFamily);
                Assert.Equal((byte)238, projectedRun.FontCharacterSet);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCommentRichTextVerticalTextAlignment() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("CommentEscapement");
                    sheet.CellValue(1, 1, "Comment");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Raised comment") {
                                FontName = "Arial",
                                VerticalTextAlignment = VerticalAlignmentRunValues.Subscript
                            }
                        },
                        "Reviewer");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                LegacyXlsCommentFormattingRun formattingRun = Assert.Single(legacyComment.FormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal(LegacyXlsFontEscapement.Subscript, font.Escapement);

                ExcelCommentInfo projectedComment = Assert.Single(result.Document.Sheets[0].GetComments());
                ExcelRichTextRun projectedRun = Assert.Single(projectedComment.RichTextRuns);
                Assert.Equal(VerticalAlignmentRunValues.Subscript, projectedRun.VerticalTextAlignment);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCommentRichTextFontOptionFlags() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("CommentFontFlags");
                    sheet.CellValue(1, 1, "Comment");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Comment flags") {
                                FontName = "Arial",
                                Outline = true,
                                Shadow = true,
                                Condense = true,
                                Extend = true
                            }
                        },
                        "Reviewer");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                LegacyXlsCommentFormattingRun formattingRun = Assert.Single(legacyComment.FormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.True(font.Outline);
                Assert.True(font.Shadow);
                Assert.True(font.Condense);
                Assert.True(font.Extend);

                ExcelCommentInfo projectedComment = Assert.Single(result.Document.Sheets[0].GetComments());
                ExcelRichTextRun projectedRun = Assert.Single(projectedComment.RichTextRuns);
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
        public void LegacyXls_NativeSave_WritesCommentRichTextUnderlineStyle() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("CommentUnderline");
                    sheet.CellValue(1, 1, "Comment");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Double accounting") {
                                FontName = "Arial",
                                UnderlineStyle = UnderlineValues.DoubleAccounting
                            }
                        },
                        "Reviewer");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                LegacyXlsCommentFormattingRun formattingRun = Assert.Single(legacyComment.FormattingRuns);
                LegacyXlsFont font = GetLegacyFont(result.Workbook, formattingRun.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal((byte)0x22, font.UnderlineStyle);

                ExcelCommentInfo projectedComment = Assert.Single(result.Document.Sheets[0].GetComments());
                ExcelRichTextRun projectedRun = Assert.Single(projectedComment.RichTextRuns);
                Assert.Equal(UnderlineValues.DoubleAccounting, projectedRun.UnderlineStyle);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedCommentTextPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment text payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Comment");
                sheet.SetLegacyComment(1, 1, new string('C', 9000), "Reviewer", visible: false, anchor: null);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedCommentAuthorPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment author payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Comment");
                sheet.SetLegacyComment(1, 1, "Supported text", new string('A', 9000), visible: false, anchor: null);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedCommentRichTextRunMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment rich-text run metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Comment");
                sheet.SetCommentRichText(
                    1,
                    1,
                    new[] {
                        new ExcelRichTextRun("Comment metadata") {
                            FontName = "Arial"
                        }
                    },
                    "Reviewer");

                Run run = sheet.WorksheetPart.WorksheetCommentsPart!
                    .Comments!
                    .Descendants<Run>()
                    .Single();
                run.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.WorksheetCommentsPart.Comments.Save();
            });
        }
    }
}
