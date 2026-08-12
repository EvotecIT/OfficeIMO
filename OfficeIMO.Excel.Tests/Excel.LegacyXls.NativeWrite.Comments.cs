using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Write;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesVisibleWorksheetCommentAnchors() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CommentAnchor");
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
        public void LegacyXls_NativeSave_WritesCommentsAcrossNonAdjacentWorksheets() {
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create()) {
                    ExcelSheet first = document.AddWorksheet("First");
                    first.CellValue(1, 1, "First value");
                    first.SetLegacyComment(1, 1, "First comment", "Reviewer", visible: false, anchor: null);

                    ExcelSheet middle = document.AddWorksheet("Middle");
                    middle.CellValue(1, 1, "No comment");

                    ExcelSheet third = document.AddWorksheet("Third");
                    third.CellValue(2, 2, "Third value");
                    third.SetLegacyComment(2, 2, "Third comment", "Reviewer", visible: true, anchor: null);

                    document.Save(xlsOutputPath, new ExcelSaveOptions { DisableFastPackageWriter = true });
                }

                AssertWorkbookOpensViaExcelComWhenAvailable(
                    xlsOutputPath,
                    "The BIFF8 workbook with comments on non-adjacent worksheets failed to open in desktop Excel.");

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                Assert.Equal(3, result.Workbook.Worksheets.Count);
                Assert.Equal("First comment", Assert.Single(result.Workbook.Worksheets[0].Comments).Text);
                Assert.Empty(result.Workbook.Worksheets[1].Comments);
                Assert.Equal("Third comment", Assert.Single(result.Workbook.Worksheets[2].Comments).Text);
            } finally {
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_ClampsCommentAnchorsToBiff8Grid() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("EdgeComment");
                    sheet.CellValue(65536, 256, "Edge note");
                    sheet.SetLegacyComment(
                        65536,
                        256,
                        "Comment near the BIFF8 edge",
                        "Reviewer",
                        visible: true,
                        new ExcelCommentAnchor(255, 15, 65535, 2, 260, 15, 65540, 16));

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                LegacyXlsDrawingAnchor anchor = legacyComment.Anchor!;
                Assert.Equal((ushort)255, anchor.StartColumn);
                Assert.Equal((ushort)255, anchor.EndColumn);
                Assert.Equal((ushort)65535, anchor.StartRow);
                Assert.Equal((ushort)65535, anchor.EndRow);
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
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CommentFonts");
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
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CommentEscapement");
                    sheet.CellValue(1, 1, "Comment");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Raised comment") {
                                FontName = "Arial",
                                VerticalTextAlignment = ExcelVerticalTextAlignment.Subscript
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
                Assert.Equal(ExcelVerticalTextAlignment.Subscript, projectedRun.VerticalTextAlignment);
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
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CommentFontFlags");
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
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CommentUnderline");
                    sheet.CellValue(1, 1, "Comment");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Double accounting") {
                                FontName = "Arial",
                                UnderlineStyle = ExcelUnderlineStyle.DoubleAccounting
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
                Assert.Equal(ExcelUnderlineStyle.DoubleAccounting, projectedRun.UnderlineStyle);
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

        [Theory]
        [InlineData(1023, true)]
        [InlineData(1024, false)]
        public void LegacyXls_CommentDrawingCountBoundary_MatchesOfficeArtCluster(int commentCount, bool expected) {
            Assert.Equal(expected, LegacyXlsCommentWriter.SupportsCommentCount(commentCount));
        }

        [Theory]
        [InlineData(4093, true)]
        [InlineData(4094, false)]
        public void LegacyXls_CommentDrawingSheetBoundary_MatchesOfficeArtIdentifierLimit(int sheetIndex, bool expected) {
            Assert.Equal(expected, LegacyXlsCommentWriter.SupportsCommentDrawingSheetIndex(sheetIndex));
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentCountsBeyondOneOfficeArtClusterBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment counts outside BIFF8 limits", (document, sheet) => {
                for (int row = 1; row <= 1024; row++) {
                    sheet.CellValue(row, 1, row);
                    sheet.SetLegacyComment(row, 1, "Comment " + row, "Reviewer", visible: false, anchor: null);
                }
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
