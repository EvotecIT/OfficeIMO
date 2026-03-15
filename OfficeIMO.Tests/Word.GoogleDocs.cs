using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordInspectionSnapshot_ExposesOfficeIMOBodyModel() {
            string filePath = Path.Combine(_directoryWithFiles, "WordInspectionSnapshot.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using var document = BuildGoogleDocsSampleDocument(filePath, imagePath);

            var snapshot = document.CreateInspectionSnapshot();

            Assert.Equal(filePath, snapshot.FilePath);
            Assert.Equal("Google Docs Snapshot", snapshot.Title);
            Assert.Equal(2, snapshot.Sections.Count);

            var firstSection = snapshot.Sections[0];
            Assert.Equal(0, firstSection.Index);
            Assert.Null(firstSection.SectionBreakType);
            Assert.Equal(3, firstSection.Elements.Count);

            var intro = Assert.IsType<WordParagraphSnapshot>(firstSection.Elements[0]);
            Assert.Equal("Heading1", intro.StyleName);
            Assert.False(string.IsNullOrWhiteSpace(intro.Alignment));
            Assert.Equal("Intro Bold Portal", intro.Text);
            Assert.Equal(3, intro.Runs.Count);
            Assert.True(intro.Runs[1].Bold);
            Assert.NotNull(intro.Runs[2].HyperlinkUri);
            Assert.Equal("https://example.com/", intro.Runs[2].HyperlinkUri);

            var imageParagraph = Assert.IsType<WordParagraphSnapshot>(firstSection.Elements[1]);
            Assert.Single(imageParagraph.Runs);
            Assert.Equal("Image ", imageParagraph.Text);
            Assert.NotNull(imageParagraph.Runs[0].InlineImage);
            Assert.Equal("Logo", imageParagraph.Runs[0].InlineImage!.Description);
            Assert.Equal("Brand", imageParagraph.Runs[0].InlineImage!.Title);
            Assert.True(imageParagraph.Runs[0].InlineImage!.IsInline);

            var table = Assert.IsType<WordTableSnapshot>(firstSection.Elements[2]);
            Assert.Equal("TableGrid", table.StyleName);
            Assert.Equal("Summary", table.Title);
            Assert.Equal("Demo table", table.Description);
            Assert.True(table.RepeatHeaderRow);
            Assert.Equal(2, table.RowCount);
            Assert.Equal(2, table.ColumnCount);
            Assert.Equal("Name", table.Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal("Value", table.Rows[0].Cells[1].Paragraphs[0].Text);
            Assert.Equal("Alpha", table.Rows[1].Cells[0].Paragraphs[0].Text);
            Assert.Equal("42", table.Rows[1].Cells[1].Paragraphs[0].Text);

            var secondSection = snapshot.Sections[1];
            Assert.Equal("NextPage", secondSection.SectionBreakType);
            Assert.Single(secondSection.Elements);
            var trailingParagraph = Assert.IsType<WordParagraphSnapshot>(secondSection.Elements[0]);
            Assert.Equal("Second section", trailingParagraph.Text);
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_EmitsParagraphAndTableRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsBatch.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using var document = BuildGoogleDocsSampleDocument(filePath, imagePath);

            var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                Title = "OfficeIMO Docs Export"
            });

            Assert.Equal("OfficeIMO Docs Export", batch.Title);
            Assert.Equal(2, batch.Snapshot.Sections.Count);
            Assert.Equal(4, batch.Requests.Count);

            var intro = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
            Assert.Equal(0, intro.SectionIndex);
            Assert.Equal(0, intro.ElementIndex);
            Assert.Equal("Heading1", intro.Paragraph.StyleName);
            Assert.Equal("Intro Bold Portal", intro.Paragraph.Text);
            Assert.True(intro.Paragraph.Runs[1].Bold);
            Assert.Equal("https://example.com/", intro.Paragraph.Runs[2].Link!.Uri);

            var imageParagraph = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[1]);
            Assert.NotNull(imageParagraph.Paragraph.Runs[0].InlineImage);
            Assert.Equal("Logo", imageParagraph.Paragraph.Runs[0].InlineImage!.Description);

            var table = Assert.IsType<GoogleDocsInsertTableRequest>(batch.Requests[2]);
            Assert.Equal(2, table.Table.RowCount);
            Assert.Equal(2, table.Table.ColumnCount);
            Assert.Equal("Summary", table.Table.Title);
            Assert.Equal("Alpha", table.Table.Rows[1].Cells[0].Paragraphs[0].Text);
            Assert.Equal("42", table.Table.Rows[1].Cells[1].Paragraphs[0].Text);

            var secondSection = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[3]);
            Assert.Equal(1, secondSection.SectionIndex);
            Assert.Equal("Second section", secondSection.Paragraph.Text);

            Assert.Contains(batch.Report.Notices, notice => notice.Feature == "Sections");
            Assert.Contains(batch.Report.Notices, notice => notice.Feature == "TextStyles");
            Assert.Contains(batch.Report.Notices, notice => notice.Feature == "InlineImages");
            Assert.Contains(batch.Report.Notices, notice => notice.Feature == "Tables");
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsParagraphStyleAndTableRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsPayload.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            try {
                using var document = BuildGoogleDocsSampleDocument(filePath, imagePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Payload Export"
                });
                var imageParagraph = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[1]);
                var inlineImage = Assert.IsType<GoogleDocsInlineImage>(imageParagraph.Paragraph.Runs[0].InlineImage);
                var imageUris = new Dictionary<GoogleDocsInlineImage, string> {
                    [inlineImage] = "https://drive.google.com/uc?export=download&id=image123"
                };

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch, imageUris);

                Assert.Contains(payload.Requests, request => request.InsertTable != null && request.InsertTable.Rows == 2 && request.InsertTable.Columns == 2);
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "Intro Bold Portal\n");
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "Image \n");
                var inlineImageRequest = Assert.Single(payload.Requests, request => request.InsertInlineImage?.Uri == "https://drive.google.com/uc?export=download&id=image123");
                Assert.Equal(7, inlineImageRequest.InsertInlineImage!.Location.Index);
                Assert.NotNull(inlineImageRequest.InsertInlineImage.ObjectSize);

                var headingStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.NamedStyleType == "HEADING_1");
                Assert.Equal("alignment,namedStyleType", string.Join(",", headingStyle.UpdateParagraphStyle!.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.Equal("CENTER", headingStyle.UpdateParagraphStyle.ParagraphStyle.Alignment);

                var boldStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.Bold == true);
                Assert.Equal(7, boldStyle.UpdateTextStyle!.Range.StartIndex);
                Assert.Equal(11, boldStyle.UpdateTextStyle.Range.EndIndex);

                var hyperlinkStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.Link?.Url == "https://example.com/");
                Assert.Equal(11, hyperlinkStyle.UpdateTextStyle!.Range.StartIndex);
                Assert.Equal(18, hyperlinkStyle.UpdateTextStyle.Range.EndIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsNativeBulletAndNumberingRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsLists.docx");

            try {
                using var document = BuildGoogleDocsListDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "List Export"
                });

                var listParagraphs = batch.Requests
                    .OfType<GoogleDocsInsertParagraphRequest>()
                    .Where(request => request.Paragraph.IsListItem)
                    .ToList();

                Assert.Equal(4, listParagraphs.Count);
                Assert.False(listParagraphs[0].Paragraph.IsOrderedList);
                Assert.Equal(0, listParagraphs[0].Paragraph.ListLevel);
                Assert.False(listParagraphs[1].Paragraph.IsOrderedList);
                Assert.Equal(1, listParagraphs[1].Paragraph.ListLevel);
                Assert.True(listParagraphs[2].Paragraph.IsOrderedList);
                Assert.True(listParagraphs[3].Paragraph.IsOrderedList);
                Assert.Equal(1, listParagraphs[3].Paragraph.ListLevel);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);

                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "First bullet\n");
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "\tNested bullet\n");
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "First step\n");
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "\tNested step\n");

                Assert.Contains(payload.Requests, request => request.CreateParagraphBullets?.BulletPreset == "BULLET_DISC_CIRCLE_SQUARE");
                Assert.Contains(payload.Requests, request => request.CreateParagraphBullets?.BulletPreset == "NUMBERED_DECIMAL_NESTED");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsHighlightedRunStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsHighlight.docx");

            try {
                using var document = BuildGoogleDocsHighlightDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Highlight Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                var highlightedRun = Assert.Single(paragraphRequest.Paragraph.Runs, run => string.Equals(run.HighlightColor, "Yellow", StringComparison.OrdinalIgnoreCase));
                Assert.Equal("Highlighted", highlightedRun.Text);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var highlightStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.BackgroundColor != null);
                Assert.Equal("backgroundColor", highlightStyle.UpdateTextStyle!.Fields);
                Assert.Equal(1d, highlightStyle.UpdateTextStyle.TextStyle.BackgroundColor!.Color.RgbColor.Red);
                Assert.Equal(1d, highlightStyle.UpdateTextStyle.TextStyle.BackgroundColor.Color.RgbColor.Green);
                Assert.Equal(0d, highlightStyle.UpdateTextStyle.TextStyle.BackgroundColor.Color.RgbColor.Blue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsBaselineOffsets() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsBaseline.docx");

            try {
                using var document = BuildGoogleDocsBaselineDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Baseline Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Contains(paragraphRequest.Paragraph.Runs, run => string.Equals(run.VerticalTextAlignment, "Superscript", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(paragraphRequest.Paragraph.Runs, run => string.Equals(run.VerticalTextAlignment, "Subscript", StringComparison.OrdinalIgnoreCase));

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                Assert.Contains(payload.Requests, request => request.UpdateTextStyle?.TextStyle.BaselineOffset == "SUPERSCRIPT");
                Assert.Contains(payload.Requests, request => request.UpdateTextStyle?.TextStyle.BaselineOffset == "SUBSCRIPT");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsSmallCapsStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSmallCaps.docx");

            try {
                using var document = BuildGoogleDocsSmallCapsDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "SmallCaps Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Contains(paragraphRequest.Paragraph.Runs, run => string.Equals(run.CapsStyle, "SmallCaps", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(paragraphRequest.Paragraph.Runs, run => string.Equals(run.CapsStyle, "Caps", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "TextStyles" && notice.Message.Contains("smallCaps", StringComparison.OrdinalIgnoreCase));

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var smallCapsStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.SmallCaps == true);
                Assert.Equal("smallCaps", smallCapsStyle.UpdateTextStyle!.Fields);
                Assert.DoesNotContain(payload.Requests, request => request.UpdateTextStyle?.TextStyle.SmallCaps == false);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsWeightedFontFamilyStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsFontFamily.docx");

            try {
                using var document = BuildGoogleDocsFontFamilyDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Font Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                var fontRun = Assert.Single(paragraphRequest.Paragraph.Runs, run => string.Equals(run.FontFamily, "Consolas", StringComparison.Ordinal));
                Assert.Equal("Mono", fontRun.Text);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var fontStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.WeightedFontFamily?.FontFamily == "Consolas");
                Assert.Equal("weightedFontFamily", fontStyle.UpdateTextStyle!.Fields);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsParagraphIndentationAndSpacing() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphLayout.docx");

            try {
                using var document = BuildGoogleDocsParagraphLayoutDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Paragraph Layout Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Equal(24d, paragraphRequest.Paragraph.IndentStartPoints);
                Assert.Equal(12d, paragraphRequest.Paragraph.IndentEndPoints);
                Assert.Equal(18d, paragraphRequest.Paragraph.IndentFirstLinePoints);
                Assert.Equal(6d, paragraphRequest.Paragraph.SpaceAbovePoints);
                Assert.Equal(9d, paragraphRequest.Paragraph.SpaceBelowPoints);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.IndentStart != null);
                Assert.Equal(
                    "indentEnd,indentFirstLine,indentStart,spaceAbove,spaceBelow",
                    string.Join(",", paragraphStyle.UpdateParagraphStyle!.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.Equal(24d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.IndentStart!.Magnitude);
                Assert.Equal(12d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.IndentEnd!.Magnitude);
                Assert.Equal(18d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.IndentFirstLine!.Magnitude);
                Assert.Equal(6d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.SpaceAbove!.Magnitude);
                Assert.Equal(9d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.SpaceBelow!.Magnitude);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsRightToLeftParagraphDirection() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphDirection.docx");

            try {
                using var document = BuildGoogleDocsRightToLeftParagraphDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "RTL Paragraph Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.True(paragraphRequest.Paragraph.IsRightToLeft);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.Direction == "RIGHT_TO_LEFT");
                Assert.Equal("direction", paragraphStyle.UpdateParagraphStyle!.Fields);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsParagraphPaginationControls() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphPagination.docx");

            try {
                using var document = BuildGoogleDocsParagraphPaginationDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Paragraph Pagination Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.True(paragraphRequest.Paragraph.KeepWithNext);
                Assert.True(paragraphRequest.Paragraph.KeepLinesTogether);
                Assert.True(paragraphRequest.Paragraph.AvoidWidowAndOrphan);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.KeepWithNext == true);
                Assert.Equal(
                    "avoidWidowAndOrphan,keepLinesTogether,keepWithNext",
                    string.Join(",", paragraphStyle.UpdateParagraphStyle!.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.True(paragraphStyle.UpdateParagraphStyle.ParagraphStyle.KeepWithNext);
                Assert.True(paragraphStyle.UpdateParagraphStyle.ParagraphStyle.KeepLinesTogether);
                Assert.True(paragraphStyle.UpdateParagraphStyle.ParagraphStyle.AvoidWidowAndOrphan);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsParagraphTabStops() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphTabStops.docx");

            try {
                using var document = BuildGoogleDocsParagraphTabStopsDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Paragraph Tab Stops Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Equal(2, paragraphRequest.Paragraph.TabStops.Count);
                Assert.Equal("Left", paragraphRequest.Paragraph.TabStops[0].Alignment, StringComparer.OrdinalIgnoreCase);
                Assert.Equal(72d, paragraphRequest.Paragraph.TabStops[0].OffsetPoints);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "ParagraphStyles" && notice.Message.Contains("tab", StringComparison.OrdinalIgnoreCase));

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.TabStops?.Count == 2);
                Assert.Equal("tabStops", paragraphStyle.UpdateParagraphStyle!.Fields);
                Assert.Equal("START", paragraphStyle.UpdateParagraphStyle.ParagraphStyle.TabStops![0].Alignment);
                Assert.Equal(72d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.TabStops[0].Offset.Magnitude);
                Assert.Equal("DECIMAL", paragraphStyle.UpdateParagraphStyle.ParagraphStyle.TabStops[1].Alignment);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsSectionLayout() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSectionLayout.docx");

            try {
                using var document = BuildGoogleDocsSectionLayoutDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Section Layout Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.NotNull(paragraphRequest.SectionStyle);
                Assert.Equal("landscape", paragraphRequest.SectionStyle!.Orientation, ignoreCase: true);
                Assert.Equal(595.3d, paragraphRequest.SectionStyle!.PageWidthPoints);
                Assert.Equal(419.55d, paragraphRequest.SectionStyle.PageHeightPoints);
                Assert.Equal(36d, paragraphRequest.SectionStyle.MarginTopPoints);
                Assert.Equal(54d, paragraphRequest.SectionStyle.MarginLeftPoints);
                Assert.True(paragraphRequest.SectionStyle.UseFirstPageHeaderFooter);
                Assert.Equal(3, paragraphRequest.SectionStyle.PageNumberStart);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var sectionStyle = Assert.Single(payload.Requests, request => request.UpdateSectionStyle?.SectionStyle.PageSize != null);
                Assert.Equal(
                    "flipPageOrientation,marginBottom,marginFooter,marginHeader,marginLeft,marginRight,marginTop,pageNumberStart,pageSize,useFirstPageHeaderFooter",
                    string.Join(",", sectionStyle.UpdateSectionStyle!.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.Equal(595.3d, sectionStyle.UpdateSectionStyle.SectionStyle.PageSize!.Width!.Magnitude);
                Assert.Equal(419.55d, sectionStyle.UpdateSectionStyle.SectionStyle.PageSize.Height!.Magnitude);
                Assert.Equal(36d, sectionStyle.UpdateSectionStyle.SectionStyle.MarginTop!.Magnitude);
                Assert.Equal(18d, sectionStyle.UpdateSectionStyle.SectionStyle.MarginHeader!.Magnitude);
                Assert.True(sectionStyle.UpdateSectionStyle.SectionStyle.FlipPageOrientation);
                Assert.True(sectionStyle.UpdateSectionStyle.SectionStyle.UseFirstPageHeaderFooter);
                Assert.Equal(3, sectionStyle.UpdateSectionStyle.SectionStyle.PageNumberStart);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsSectionColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSectionColumns.docx");

            try {
                using var document = BuildGoogleDocsSectionColumnsDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Section Columns Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.NotNull(paragraphRequest.SectionStyle);
                Assert.Equal(2, paragraphRequest.SectionStyle!.ColumnCount);
                Assert.Equal(18d, paragraphRequest.SectionStyle.ColumnSpacingPoints);
                Assert.True(paragraphRequest.SectionStyle.HasColumnSeparator);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var sectionStyle = Assert.Single(payload.Requests, request => request.UpdateSectionStyle?.SectionStyle.ColumnProperties?.Count == 2);
                Assert.Contains("columnProperties", sectionStyle.UpdateSectionStyle!.Fields);
                Assert.Contains("columnSeparatorStyle", sectionStyle.UpdateSectionStyle.Fields);
                Assert.Equal(18d, sectionStyle.UpdateSectionStyle.SectionStyle.ColumnProperties![0].PaddingEnd!.Magnitude);
                Assert.Equal("BETWEEN_EACH_COLUMN", sectionStyle.UpdateSectionStyle.SectionStyle.ColumnSeparatorStyle);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsAutoLineSpacing() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphLineSpacing.docx");

            try {
                using var document = BuildGoogleDocsAutoLineSpacingDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Line Spacing Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Equal(150d, paragraphRequest.Paragraph.LineSpacingPercent);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.LineSpacing == 150d);
                Assert.Equal("lineSpacing", paragraphStyle.UpdateParagraphStyle!.Fields);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_ApproximatesExactLineSpacing() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphExactLineSpacing.docx");

            try {
                using var document = BuildGoogleDocsExactLineSpacingDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Exact Line Spacing Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Equal(150d, paragraphRequest.Paragraph.LineSpacingPercent);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "ParagraphStyles" && notice.Message.Contains("approximations", StringComparison.OrdinalIgnoreCase));

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.LineSpacing == 150d);
                Assert.Equal("lineSpacing", paragraphStyle.UpdateParagraphStyle!.Fields);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_ApproximatesAtLeastLineSpacing() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphAtLeastLineSpacing.docx");

            try {
                using var document = BuildGoogleDocsAtLeastLineSpacingDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "AtLeast Line Spacing Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Equal(150d, paragraphRequest.Paragraph.LineSpacingPercent);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "ParagraphStyles" && notice.Message.Contains("approximations", StringComparison.OrdinalIgnoreCase));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsParagraphShading() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphShading.docx");

            try {
                using var document = BuildGoogleDocsParagraphShadingDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Paragraph Shading Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.Equal("d9eaf7", paragraphRequest.Paragraph.ShadingFillColorHex);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.Shading?.BackgroundColor != null);
                Assert.Equal("shading", paragraphStyle.UpdateParagraphStyle!.Fields);
                Assert.Equal(0.8509803921568627d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.Shading!.BackgroundColor!.Color.RgbColor.Red);
                Assert.Equal(0.9176470588235294d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.Shading.BackgroundColor.Color.RgbColor.Green);
                Assert.Equal(0.9686274509803922d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.Shading.BackgroundColor.Color.RgbColor.Blue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsParagraphBorders() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsParagraphBorders.docx");

            try {
                using var document = BuildGoogleDocsParagraphBorderDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Paragraph Border Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                Assert.NotNull(paragraphRequest.Paragraph.TopBorder);
                Assert.Equal("single", paragraphRequest.Paragraph.TopBorder!.Style, StringComparer.OrdinalIgnoreCase);
                Assert.Equal("336699", paragraphRequest.Paragraph.TopBorder.ColorHex);
                Assert.Equal(8U, paragraphRequest.Paragraph.TopBorder.Size);
                Assert.Equal(6U, paragraphRequest.Paragraph.TopBorder.Space);
                Assert.NotNull(paragraphRequest.Paragraph.LeftBorder);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var paragraphStyle = Assert.Single(payload.Requests, request => request.UpdateParagraphStyle?.ParagraphStyle.BorderTop != null);
                Assert.Equal(
                    "borderLeft,borderTop",
                    string.Join(",", paragraphStyle.UpdateParagraphStyle!.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.Equal(1d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.BorderTop!.Width!.Magnitude);
                Assert.Equal(6d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.BorderTop.Padding!.Magnitude);
                Assert.Equal("SOLID", paragraphStyle.UpdateParagraphStyle.ParagraphStyle.BorderTop.DashStyle);
                Assert.Equal(0.2d, paragraphStyle.UpdateParagraphStyle.ParagraphStyle.BorderTop.Color!.Color.RgbColor.Red);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsNativePageBreakRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsPageBreak.docx");

            try {
                using var document = BuildGoogleDocsPageBreakDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Page Break Export"
                });

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);

                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "Intro paragraph\n");
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "Starts on next page\n");
                var pageBreakRequest = Assert.Single(payload.Requests, request => request.InsertPageBreak != null);
                Assert.Equal(1, pageBreakRequest.InsertPageBreak!.Location.Index);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "PageBreaks");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_CompilesFootnotes() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsFootnotes.docx");

            try {
                using var document = BuildGoogleDocsFootnoteDocument(filePath);
                var snapshot = document.CreateInspectionSnapshot();
                var paragraph = Assert.IsType<WordParagraphSnapshot>(snapshot.Sections[0].Elements[0]);
                var footnoteRun = Assert.Single(paragraph.Runs, run => run.Footnote != null);
                Assert.Single(footnoteRun.Footnote!.Paragraphs);
                Assert.Equal("Footnote text", footnoteRun.Footnote.Paragraphs[0].Text);

                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Footnote Export"
                });

                var paragraphRequest = Assert.Single(batch.Requests.OfType<GoogleDocsInsertParagraphRequest>());
                var compiledFootnoteRun = Assert.Single(paragraphRequest.Paragraph.Runs, run => run.Footnote != null);
                Assert.Single(compiledFootnoteRun.Footnote!.Paragraphs);
                Assert.Equal("Footnote text", compiledFootnoteRun.Footnote.Paragraphs[0].Text);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "Footnotes");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_CompilesBookmarksAndInternalAnchors() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsBookmarks.docx");

            try {
                using var document = BuildGoogleDocsBookmarkDocument(filePath);
                var snapshot = document.CreateInspectionSnapshot();
                var targetParagraph = Assert.IsType<WordParagraphSnapshot>(snapshot.Sections[0].Elements[1]);
                Assert.Equal("TargetBookmark", targetParagraph.BookmarkName);

                var sourceParagraph = Assert.IsType<WordParagraphSnapshot>(snapshot.Sections[0].Elements[0]);
                var internalLinkRun = Assert.Single(sourceParagraph.Runs, run => !string.IsNullOrWhiteSpace(run.HyperlinkAnchor));
                Assert.Equal("TargetBookmark", internalLinkRun.HyperlinkAnchor);

                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Bookmark Export"
                });

                var compiledTarget = batch.Requests
                    .OfType<GoogleDocsInsertParagraphRequest>()
                    .Single(request => request.Paragraph.BookmarkName == "TargetBookmark");
                Assert.Equal("Target paragraph", compiledTarget.Paragraph.Text);

                var compiledSource = batch.Requests
                    .OfType<GoogleDocsInsertParagraphRequest>()
                    .First(request => request.Paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Link?.Anchor)));
                Assert.Equal("TargetBookmark", compiledSource.Paragraph.Runs.Single(run => !string.IsNullOrWhiteSpace(run.Link?.Anchor)).Link!.Anchor);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "Bookmarks");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsCreateFootnoteRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsFootnotePayload.docx");

            try {
                using var document = BuildGoogleDocsFootnoteDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Footnote Export"
                });

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);

                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "Body text\n");
                var footnoteRequest = Assert.Single(payload.Requests, request => request.CreateFootnote != null);
                Assert.Equal(10, footnoteRequest.CreateFootnote!.Location.Index);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsCreateNamedRangeRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsNamedRangePayload.docx");

            try {
                using var document = BuildGoogleDocsBookmarkDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Bookmark Export"
                });

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);

                var namedRangeRequest = Assert.Single(payload.Requests, request => request.CreateNamedRange != null);
                Assert.Equal("TargetBookmark", namedRangeRequest.CreateNamedRange!.Name);
                Assert.Equal(1, namedRangeRequest.CreateNamedRange.Range.StartIndex);
                Assert.Equal(17, namedRangeRequest.CreateNamedRange.Range.EndIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsCreateNamedRangeRequestsForFootnotes() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsFootnoteBookmarkPayload.docx");

            try {
                using var document = BuildGoogleDocsFootnoteBookmarkDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Footnote Bookmark Export"
                });

                var paragraphRequest = Assert.IsType<GoogleDocsInsertParagraphRequest>(batch.Requests[0]);
                var footnote = Assert.Single(paragraphRequest.Paragraph.Runs, run => run.Footnote != null).Footnote;
                Assert.NotNull(footnote);

                var payload = GoogleDocsApiPayloadBuilder.BuildFootnoteBatchUpdatePayload(footnote!, batch.Report, "fn-bookmark-123", null);
                var namedRangeRequest = Assert.Single(payload.Requests, request => request.CreateNamedRange != null);
                Assert.Equal("FootnoteBookmark", namedRangeRequest.CreateNamedRange!.Name);
                Assert.Equal("fn-bookmark-123", namedRangeRequest.CreateNamedRange.Range.SegmentId);
                Assert.Equal(1, namedRangeRequest.CreateNamedRange.Range.StartIndex);
                Assert.Equal(14, namedRangeRequest.CreateNamedRange.Range.EndIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_CompilesBodyTableBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsTableBookmarkPayload.docx");

            try {
                using var document = BuildGoogleDocsTableBookmarkDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Table Bookmark Export"
                });

                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateSingleCellBodyTableDocumentStateJson("doc-table-bookmark", "Table Bookmark Export"));
                Assert.NotNull(documentState);

                var prepared = GoogleDocsApiPayloadBuilder.BuildPreparedTableContentBatchUpdate(batch, documentState!);
                var namedRangeRequest = Assert.Single(prepared.Payload.Requests, request => request.CreateNamedRange != null);
                Assert.Equal("CellBookmark", namedRangeRequest.CreateNamedRange!.Name);
                Assert.Equal(11, namedRangeRequest.CreateNamedRange.Range.StartIndex);
                Assert.Equal(26, namedRangeRequest.CreateNamedRange.Range.EndIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsSegmentCreateNamedRangeRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsHeaderBookmarkPayload.docx");

            try {
                using var document = BuildGoogleDocsHeaderBookmarkDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Header Bookmark Export"
                });

                var headerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "header");
                var payload = GoogleDocsApiPayloadBuilder.BuildSegmentBatchUpdatePayload(headerSegment, batch.Report, "headerBookmark123", null);

                var namedRangeRequest = Assert.Single(payload.Requests, request => request.CreateNamedRange != null);
                Assert.Equal("HeaderBookmark", namedRangeRequest.CreateNamedRange!.Name);
                Assert.Equal("headerBookmark123", namedRangeRequest.CreateNamedRange.Range.SegmentId);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsNativeSectionBreakRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSectionBreak.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            try {
                using var document = BuildGoogleDocsSampleDocument(filePath, imagePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Section Break Export"
                });

                var secondSectionParagraph = batch.Requests
                    .OfType<GoogleDocsInsertParagraphRequest>()
                    .Single(request => request.SectionIndex == 1);
                Assert.True(secondSectionParagraph.Paragraph.StartsNewSectionBefore);
                Assert.Equal("NextPage", secondSectionParagraph.Paragraph.SectionBreakType);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var sectionBreakRequest = Assert.Single(payload.Requests, request => request.InsertSectionBreak != null);
                Assert.Equal("NEXT_PAGE", sectionBreakRequest.InsertSectionBreak!.SectionType);
                Assert.Equal(1, sectionBreakRequest.InsertSectionBreak.Location.Index);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "SectionBreaks");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsNativeSectionBreakRequestsBeforeLeadingTable() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSectionBreakLeadingTable.docx");

            try {
                using var document = BuildGoogleDocsLeadingTableSectionDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Section Break Leading Table Export"
                });

                var secondSectionTable = batch.Requests
                    .OfType<GoogleDocsInsertTableRequest>()
                    .Single(request => request.SectionIndex == 1);
                Assert.True(secondSectionTable.StartsNewSectionBefore);
                Assert.Equal("NextPage", secondSectionTable.SectionBreakType);

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                Assert.Contains(payload.Requests, request => request.InsertTable?.Rows == 2 && request.InsertTable.Columns == 2);

                var sectionBreakRequest = Assert.Single(payload.Requests, request => request.InsertSectionBreak != null);
                Assert.Equal("NEXT_PAGE", sectionBreakRequest.InsertSectionBreak!.SectionType);
                Assert.Equal(1, sectionBreakRequest.InsertSectionBreak.Location.Index);
                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "SectionBreaks");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_CompilesDefaultHeaderAndFooterSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsHeaderFooter.docx");

            try {
                using var document = BuildGoogleDocsHeaderFooterDocument(filePath);
                var snapshot = document.CreateInspectionSnapshot();
                Assert.NotNull(snapshot.Sections[0].DefaultHeader);
                Assert.NotNull(snapshot.Sections[0].DefaultFooter);
                Assert.Equal("Header text", snapshot.Sections[0].DefaultHeader!.Paragraphs[0].Text);
                Assert.Equal("Footer text", snapshot.Sections[0].DefaultFooter!.Paragraphs[0].Text);

                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Header Footer Export"
                });

                Assert.Equal(2, batch.Segments.Count);
                var headerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "header");
                Assert.Equal("default", headerSegment.Variant);
                Assert.Equal("Header text", headerSegment.Paragraphs[0].Text);

                var footerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "footer");
                Assert.Equal("Footer text", footerSegment.Paragraphs[0].Text);

                var payload = GoogleDocsApiPayloadBuilder.BuildSegmentBatchUpdatePayload(headerSegment, batch.Report, "header123", null);
                Assert.Contains(payload.Requests, request => request.InsertText?.Location?.SegmentId == "header123");
                Assert.Contains(payload.Requests, request => request.InsertText?.Text == "Header text\n");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_CompilesDefaultHeaderTableSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsHeaderTable.docx");

            try {
                using var document = BuildGoogleDocsHeaderTableDocument(filePath);
                var snapshot = document.CreateInspectionSnapshot();
                var headerSnapshot = Assert.IsType<WordHeaderFooterSnapshot>(snapshot.Sections[0].DefaultHeader);
                Assert.Single(headerSnapshot.Tables);
                Assert.Equal("H1", headerSnapshot.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("H4", headerSnapshot.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text);

                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Header Table Export"
                });

                var headerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "header");
                Assert.Single(headerSegment.Tables);
                Assert.Equal(2, headerSegment.Tables[0].RowCount);
                Assert.Equal(2, headerSegment.Tables[0].ColumnCount);

                var initialPayload = GoogleDocsApiPayloadBuilder.BuildSegmentBatchUpdatePayload(headerSegment, batch.Report, "headerTable123", null);
                Assert.Contains(initialPayload.Requests, request => request.InsertTable?.Location.SegmentId == "headerTable123");

                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateHeaderTableDocumentStateJson("headerTable123"));
                Assert.NotNull(documentState);

                var tablePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableContentBatchUpdatePayload(
                    headerSegment,
                    documentState!,
                    batch.Report,
                    "headerTable123");

                Assert.Contains(tablePayload.Requests, request => request.InsertText?.Location?.SegmentId == "headerTable123");
                Assert.Contains(tablePayload.Requests, request => request.InsertText?.Text == "H1\n");
                Assert.Contains(tablePayload.Requests, request => request.InsertText?.Text == "H4\n");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_CompilesFirstPageHeaderAndFooterSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsFirstPageHeaderFooter.docx");

            try {
                using var document = BuildGoogleDocsFirstPageHeaderFooterDocument(filePath);
                var snapshot = document.CreateInspectionSnapshot();
                Assert.NotNull(snapshot.Sections[0].FirstHeader);
                Assert.NotNull(snapshot.Sections[0].FirstFooter);
                Assert.Equal("First header text", snapshot.Sections[0].FirstHeader!.Paragraphs[0].Text);
                Assert.Equal("First footer text", snapshot.Sections[0].FirstFooter!.Paragraphs[0].Text);

                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "First Page Header Footer Export"
                });

                Assert.Equal(2, batch.Segments.Count);
                var headerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "header");
                Assert.Equal("first", headerSegment.Variant);
                Assert.Equal("First header text", headerSegment.Paragraphs[0].Text);

                var footerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "footer");
                Assert.Equal("first", footerSegment.Variant);
                Assert.Equal("First footer text", footerSegment.Paragraphs[0].Text);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsBatchCompiler_CompilesEvenPageHeaderAndFooterSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsEvenPageHeaderFooter.docx");

            try {
                using var document = BuildGoogleDocsEvenPageHeaderFooterDocument(filePath);
                var snapshot = document.CreateInspectionSnapshot();
                Assert.NotNull(snapshot.Sections[0].EvenHeader);
                Assert.NotNull(snapshot.Sections[0].EvenFooter);
                Assert.Equal("Even header text", snapshot.Sections[0].EvenHeader!.Paragraphs[0].Text);
                Assert.Equal("Even footer text", snapshot.Sections[0].EvenFooter!.Paragraphs[0].Text);

                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Even Page Header Footer Export"
                });

                Assert.Contains(batch.Segments, segment => segment.Kind == "header" && segment.Variant == "even");
                Assert.Contains(batch.Segments, segment => segment.Kind == "footer" && segment.Variant == "even");

                var payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
                var documentStyle = Assert.Single(payload.Requests, request => request.UpdateDocumentStyle != null);
                Assert.Equal("useEvenPageHeaderFooter", documentStyle.UpdateDocumentStyle!.Fields);
                Assert.True(documentStyle.UpdateDocumentStyle.DocumentStyle.UseEvenPageHeaderFooter);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_CompilesHeaderTableBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsHeaderTableBookmark.docx");

            try {
                using var document = BuildGoogleDocsHeaderTableBookmarkDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Header Table Bookmark Export"
                });

                var headerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "header");
                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateHeaderTableDocumentStateJson("headerTableBookmark123", "doc-header-table-bookmark", "Header Table Bookmark Export"));
                Assert.NotNull(documentState);

                var payload = GoogleDocsApiPayloadBuilder.BuildSegmentTableContentBatchUpdatePayload(
                    headerSegment,
                    documentState!,
                    batch.Report,
                    "headerTableBookmark123");

                var namedRangeRequest = Assert.Single(payload.Requests, request => request.CreateNamedRange != null);
                Assert.Equal("HeaderCellBookmark", namedRangeRequest.CreateNamedRange!.Name);
                Assert.Equal("headerTableBookmark123", namedRangeRequest.CreateNamedRange.Range.SegmentId);
                Assert.Equal(5, namedRangeRequest.CreateNamedRange.Range.StartIndex);
                Assert.Equal(7, namedRangeRequest.CreateNamedRange.Range.EndIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_ReplaysStyledAndLinkedTableCellContent() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsStyledTableCell.docx");

            try {
                using var document = BuildGoogleDocsStyledTableDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Styled Table Cell Export"
                });

                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateBodyTableDocumentStateJson("doc-table-style", "Styled Table Cell Export"));
                Assert.NotNull(documentState);

                var payload = GoogleDocsApiPayloadBuilder.BuildTableContentBatchUpdatePayload(batch, documentState!);

                var insertedCellText = Assert.Single(payload.Requests, request => request.InsertText?.Text == "Cell Bold Link\n");
                var insertedTextPayload = Assert.IsType<GoogleDocsApiInsertTextRequestPayload>(insertedCellText.InsertText);
                var insertedTextLocation = Assert.IsType<GoogleDocsApiLocationPayload>(insertedTextPayload.Location);
                Assert.Equal(25, insertedTextLocation.Index);

                var boldStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.Bold == true);
                Assert.Equal(30, boldStyle.UpdateTextStyle!.Range.StartIndex);
                Assert.Equal(34, boldStyle.UpdateTextStyle.Range.EndIndex);

                var hyperlinkStyle = Assert.Single(payload.Requests, request => request.UpdateTextStyle?.TextStyle.Link?.Url == "https://example.com/");
                Assert.Equal(34, hyperlinkStyle.UpdateTextStyle!.Range.StartIndex);
                Assert.Equal(39, hyperlinkStyle.UpdateTextStyle.Range.EndIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsNativeTablePresentationRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsTablePresentation.docx");

            try {
                using var document = BuildGoogleDocsStyledTableDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Styled Table Cell Export"
                });

                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateBodyTableDocumentStateJson("doc-table-style", "Styled Table Cell Export"));
                Assert.NotNull(documentState);

                var payload = GoogleDocsApiPayloadBuilder.BuildTableStyleBatchUpdatePayload(batch, documentState!);

                var pinnedHeaderRows = Assert.Single(payload.Requests, request => request.PinTableHeaderRows != null);
                Assert.Equal(1, pinnedHeaderRows.PinTableHeaderRows!.TableStartLocation.Index);
                Assert.Equal(1, pinnedHeaderRows.PinTableHeaderRows.PinnedHeaderRowsCount);

                var columnWidthRequests = payload.Requests.Where(request => request.UpdateTableColumnProperties != null).ToList();
                Assert.Equal(2, columnWidthRequests.Count);
                Assert.All(columnWidthRequests, request => Assert.Equal(1, request.UpdateTableColumnProperties!.TableStartLocation.Index));
                var firstColumnWidth = Assert.IsType<GoogleDocsApiUpdateTableColumnPropertiesRequestPayload>(columnWidthRequests[0].UpdateTableColumnProperties);
                var secondColumnWidth = Assert.IsType<GoogleDocsApiUpdateTableColumnPropertiesRequestPayload>(columnWidthRequests[1].UpdateTableColumnProperties);
                Assert.Equal(new[] { 0 }, firstColumnWidth.ColumnIndices);
                Assert.Equal("width,widthType", firstColumnWidth.Fields);
                Assert.Equal("FIXED_WIDTH", firstColumnWidth.TableColumnProperties.WidthType);
                Assert.Equal(72d, firstColumnWidth.TableColumnProperties.Width!.Magnitude);
                Assert.Equal(new[] { 1 }, secondColumnWidth.ColumnIndices);
                Assert.Equal(144d, secondColumnWidth.TableColumnProperties.Width!.Magnitude);

                var shadedCell = Assert.Single(payload.Requests, request => request.UpdateTableCellStyle != null);
                Assert.Equal(0, shadedCell.UpdateTableCellStyle!.TableRange.TableCellLocation.RowIndex);
                Assert.Equal(0, shadedCell.UpdateTableCellStyle.TableRange.TableCellLocation.ColumnIndex);
                Assert.Equal(
                    "backgroundColor,borderBottom,borderLeft,borderTop",
                    string.Join(",", shadedCell.UpdateTableCellStyle.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.NotNull(shadedCell.UpdateTableCellStyle.TableCellStyle.BackgroundColor);
                Assert.Equal(1d, shadedCell.UpdateTableCellStyle.TableCellStyle.BackgroundColor!.Color!.RgbColor!.Red);
                Assert.Equal(0.8d, shadedCell.UpdateTableCellStyle.TableCellStyle.BackgroundColor.Color.RgbColor.Green);
                Assert.Equal(0d, shadedCell.UpdateTableCellStyle.TableCellStyle.BackgroundColor.Color.RgbColor.Blue);
                Assert.Equal("SOLID", shadedCell.UpdateTableCellStyle.TableCellStyle.BorderLeft!.DashStyle);
                Assert.Equal(1d, shadedCell.UpdateTableCellStyle.TableCellStyle.BorderLeft.Width!.Magnitude);
                Assert.Equal(1d, shadedCell.UpdateTableCellStyle.TableCellStyle.BorderLeft.Color!.Color.RgbColor.Red);
                Assert.Equal(1.5d, shadedCell.UpdateTableCellStyle.TableCellStyle.BorderTop!.Width!.Magnitude);
                Assert.Equal(2d, shadedCell.UpdateTableCellStyle.TableCellStyle.BorderBottom!.Width!.Magnitude);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsHeaderTableCellStyleRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsHeaderTableStyle.docx");

            try {
                using var document = BuildGoogleDocsHeaderTableDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Header Table Export"
                });

                var headerSegment = Assert.Single(batch.Segments, segment => segment.Kind == "header");
                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateHeaderTableDocumentStateJson("headerTable123"));
                Assert.NotNull(documentState);

                var payload = GoogleDocsApiPayloadBuilder.BuildSegmentTableStyleBatchUpdatePayload(
                    headerSegment,
                    documentState!,
                    batch.Report,
                    "headerTable123");

                Assert.DoesNotContain(payload.Requests, request => request.PinTableHeaderRows != null);
                var shadedCell = Assert.Single(payload.Requests, request => request.UpdateTableCellStyle != null);
                Assert.Equal("headerTable123", shadedCell.UpdateTableCellStyle!.TableRange.TableCellLocation.TableStartLocation.SegmentId);
                Assert.Equal(0, shadedCell.UpdateTableCellStyle.TableRange.TableCellLocation.RowIndex);
                Assert.Equal(0, shadedCell.UpdateTableCellStyle.TableRange.TableCellLocation.ColumnIndex);
                Assert.Equal(
                    "backgroundColor,borderRight",
                    string.Join(",", shadedCell.UpdateTableCellStyle.Fields.Split(',').OrderBy(value => value, StringComparer.Ordinal)));
                Assert.Equal("SOLID", shadedCell.UpdateTableCellStyle.TableCellStyle.BorderRight!.DashStyle);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_CompilesBodyTableFootnotes() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsTableFootnotePayload.docx");

            try {
                using var document = BuildGoogleDocsTableFootnoteDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Table Footnote Export"
                });

                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateSingleCellBodyTableDocumentStateJson("doc-table-footnote", "Table Footnote Export"));
                Assert.NotNull(documentState);

                var prepared = GoogleDocsApiPayloadBuilder.BuildPreparedTableContentBatchUpdate(batch, documentState!);
                Assert.Single(prepared.Footnotes);
                var footnoteRequest = Assert.Single(prepared.Payload.Requests, request => request.CreateFootnote != null);
                Assert.Equal(20, footnoteRequest.CreateFootnote!.Location.Index);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_UploadsTableCellImages() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterTableImage.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            try {
                using var document = BuildGoogleDocsTableImageDocument(filePath, imagePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                int batchUpdateCount = 0;

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-table-image\",\"title\":\"Table Image Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id") {
                        return CreateJsonResponse("{\"id\":\"img-table-123\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/img-table-123/permissions?supportsAllDrives=true") {
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-image:batchUpdate") {
                        batchUpdateCount++;
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-image") {
                        return CreateJsonResponse(CreateBodyTableDocumentStateJson("doc-table-image", "Table Image Export"));
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Table Image Export",
                });

                Assert.Equal("doc-table-image", result.DocumentId);
                Assert.Equal(7, recordedRequests.Count);
                Assert.Equal(3, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                Assert.Contains(recordedRequests, request => request.Uri.AbsoluteUri == "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id");
                Assert.Contains(recordedRequests, request => request.Uri.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/img-table-123/permissions?supportsAllDrives=true");

                var tableBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-image:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"insertInlineImage\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"Cell \\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"insertInlineImage\"", tableBatchRequest.Body!);
                var tableStyleBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-image:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"updateTableColumnProperties\"", StringComparison.Ordinal));
                Assert.Contains("\"widthType\":\"FIXED_WIDTH\"", tableStyleBatchRequest.Body!);
                using (var tableBatchJson = JsonDocument.Parse(tableBatchRequest.Body!)) {
                    var inlineImageRequest = tableBatchJson.RootElement
                        .GetProperty("requests")
                        .EnumerateArray()
                        .First(request => request.TryGetProperty("insertInlineImage", out _))
                        .GetProperty("insertInlineImage");
                    var uri = inlineImageRequest.GetProperty("uri").GetString();
                    Assert.NotNull(uri);
                    Assert.Contains("img-table-123", uri!, StringComparison.Ordinal);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysFootnoteContent() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterFootnote.docx");

            try {
                using var document = BuildGoogleDocsFootnoteDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                int batchUpdateCount = 0;

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-footnote\",\"title\":\"Footnote Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footnote:batchUpdate") {
                        batchUpdateCount++;
                        if (body != null && body.Contains("\"createFootnote\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFootnote\":{\"footnoteId\":\"fn123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Footnote Export",
                });

                Assert.Equal("doc-footnote", result.DocumentId);
                Assert.Equal(3, recordedRequests.Count);
                Assert.Equal(2, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatchRequest = recordedRequests.First(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footnote:batchUpdate");
                Assert.Contains("\"createFootnote\"", initialBatchRequest.Body!);

                var footnoteBatchRequest = recordedRequests.Last(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footnote:batchUpdate");
                Assert.Contains("\"segmentId\":\"fn123\"", footnoteBatchRequest.Body!);
                Assert.Contains("\"text\":\"Footnote text\\n\"", footnoteBatchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsHighlightedRunBackgroundColor() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterHighlight.docx");

            try {
                using var document = BuildGoogleDocsHighlightDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-highlight\",\"title\":\"Highlight Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-highlight:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Highlight Export",
                });

                Assert.Equal("doc-highlight", result.DocumentId);
                Assert.Equal(2, recordedRequests.Count);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-highlight:batchUpdate");
                Assert.Contains("\"backgroundColor\"", batchRequest.Body!);
                Assert.Contains("\"red\":1", batchRequest.Body!);
                Assert.Contains("\"green\":1", batchRequest.Body!);
                Assert.Contains("\"blue\":0", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsBaselineOffsets() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterBaseline.docx");

            try {
                using var document = BuildGoogleDocsBaselineDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-baseline\",\"title\":\"Baseline Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-baseline:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Baseline Export",
                });

                Assert.Equal("doc-baseline", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-baseline:batchUpdate");
                Assert.Contains("\"baselineOffset\":\"SUPERSCRIPT\"", batchRequest.Body!);
                Assert.Contains("\"baselineOffset\":\"SUBSCRIPT\"", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsSmallCapsStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterSmallCaps.docx");

            try {
                using var document = BuildGoogleDocsSmallCapsDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-smallcaps\",\"title\":\"SmallCaps Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-smallcaps:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "SmallCaps Export",
                });

                Assert.Equal("doc-smallcaps", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-smallcaps:batchUpdate");
                Assert.Contains("\"smallCaps\":true", batchRequest.Body!);
                Assert.DoesNotContain("\"smallCaps\":false", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsWeightedFontFamilyStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterFontFamily.docx");

            try {
                using var document = BuildGoogleDocsFontFamilyDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-font\",\"title\":\"Font Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-font:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Font Export",
                });

                Assert.Equal("doc-font", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-font:batchUpdate");
                Assert.Contains("\"weightedFontFamily\":{\"fontFamily\":\"Consolas\"}", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsParagraphIndentationAndSpacing() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphLayout.docx");

            try {
                using var document = BuildGoogleDocsParagraphLayoutDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-layout\",\"title\":\"Paragraph Layout Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-layout:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Paragraph Layout Export",
                });

                Assert.Equal("doc-layout", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-layout:batchUpdate");
                Assert.Contains("\"indentStart\":{\"magnitude\":24", batchRequest.Body!);
                Assert.Contains("\"indentEnd\":{\"magnitude\":12", batchRequest.Body!);
                Assert.Contains("\"indentFirstLine\":{\"magnitude\":18", batchRequest.Body!);
                Assert.Contains("\"spaceAbove\":{\"magnitude\":6", batchRequest.Body!);
                Assert.Contains("\"spaceBelow\":{\"magnitude\":9", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsRightToLeftParagraphDirection() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphDirection.docx");

            try {
                using var document = BuildGoogleDocsRightToLeftParagraphDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-rtl\",\"title\":\"RTL Paragraph Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-rtl:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "RTL Paragraph Export",
                });

                Assert.Equal("doc-rtl", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-rtl:batchUpdate");
                Assert.Contains("\"direction\":\"RIGHT_TO_LEFT\"", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsAutoLineSpacing() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphLineSpacing.docx");

            try {
                using var document = BuildGoogleDocsAutoLineSpacingDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-linespacing\",\"title\":\"Line Spacing Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-linespacing:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Line Spacing Export",
                });

                Assert.Equal("doc-linespacing", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-linespacing:batchUpdate");
                Assert.Contains("\"lineSpacing\":150", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsParagraphShading() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphShading.docx");

            try {
                using var document = BuildGoogleDocsParagraphShadingDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-paragraph-shading\",\"title\":\"Paragraph Shading Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-shading:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Paragraph Shading Export",
                });

                Assert.Equal("doc-paragraph-shading", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-shading:batchUpdate");
                Assert.Contains("\"shading\":{\"backgroundColor\"", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsParagraphPaginationControls() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphPagination.docx");

            try {
                using var document = BuildGoogleDocsParagraphPaginationDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-paragraph-pagination\",\"title\":\"Paragraph Pagination Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-pagination:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Paragraph Pagination Export",
                });

                Assert.Equal("doc-paragraph-pagination", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-pagination:batchUpdate");
                Assert.Contains("\"keepWithNext\":true", batchRequest.Body!);
                Assert.Contains("\"keepLinesTogether\":true", batchRequest.Body!);
                Assert.Contains("\"avoidWidowAndOrphan\":true", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsParagraphTabStops() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphTabStops.docx");

            try {
                using var document = BuildGoogleDocsParagraphTabStopsDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-paragraph-tabstops\",\"title\":\"Paragraph Tab Stops Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-tabstops:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Paragraph Tab Stops Export",
                });

                Assert.Equal("doc-paragraph-tabstops", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-tabstops:batchUpdate");
                Assert.Contains("\"tabStops\":[", batchRequest.Body!);
                Assert.Contains("\"alignment\":\"START\"", batchRequest.Body!);
                Assert.Contains("\"alignment\":\"DECIMAL\"", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsSectionLayout() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterSectionLayout.docx");

            try {
                using var document = BuildGoogleDocsSectionLayoutDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-section-layout\",\"title\":\"Section Layout Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-section-layout:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Section Layout Export",
                });

                Assert.Equal("doc-section-layout", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-section-layout:batchUpdate");
                Assert.Contains("\"updateSectionStyle\":{", batchRequest.Body!);
                Assert.Contains("\"pageSize\":{", batchRequest.Body!);
                Assert.Contains("\"marginLeft\":{\"magnitude\":54", batchRequest.Body!);
                Assert.Contains("\"flipPageOrientation\":true", batchRequest.Body!);
                Assert.Contains("\"useFirstPageHeaderFooter\":true", batchRequest.Body!);
                Assert.Contains("\"pageNumberStart\":3", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsSectionColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterSectionColumns.docx");

            try {
                using var document = BuildGoogleDocsSectionColumnsDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-section-columns\",\"title\":\"Section Columns Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-section-columns:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Section Columns Export",
                });

                Assert.Equal("doc-section-columns", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-section-columns:batchUpdate");
                Assert.Contains("\"columnProperties\":[", batchRequest.Body!);
                Assert.Contains("\"columnSeparatorStyle\":\"BETWEEN_EACH_COLUMN\"", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_EmitsParagraphBorders() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterParagraphBorders.docx");

            try {
                using var document = BuildGoogleDocsParagraphBorderDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-paragraph-borders\",\"title\":\"Paragraph Border Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-borders:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Paragraph Border Export",
                });

                Assert.Equal("doc-paragraph-borders", result.DocumentId);
                var batchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-paragraph-borders:batchUpdate");
                Assert.Contains("\"borderTop\":{", batchRequest.Body!);
                Assert.Contains("\"borderLeft\":{", batchRequest.Body!);
                Assert.Contains("\"padding\":{\"magnitude\":6", batchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesNamedRangesForBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterBookmark.docx");

            try {
                using var document = BuildGoogleDocsBookmarkDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-bookmark\",\"title\":\"Bookmark Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-bookmark:batchUpdate") {
                        if (body != null && body.Contains("\"createNamedRange\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createNamedRange\":{\"namedRangeId\":\"nr123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Bookmark Export",
                });

                Assert.Equal("doc-bookmark", result.DocumentId);
                Assert.Equal(2, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatchRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-bookmark:batchUpdate");
                Assert.Contains("\"createNamedRange\"", initialBatchRequest.Body!);
                Assert.Contains("\"name\":\"TargetBookmark\"", initialBatchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesNamedRangesForHeaderBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterHeaderBookmark.docx");

            try {
                using var document = BuildGoogleDocsHeaderBookmarkDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-header-bookmark\",\"title\":\"Header Bookmark Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-header-bookmark") {
                        return CreateJsonResponse("{\"documentId\":\"doc-header-bookmark\",\"title\":\"Header Bookmark Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-header-bookmark:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"headerBookmark123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Header Bookmark Export",
                });

                Assert.Equal("doc-header-bookmark", result.DocumentId);
                var headerWrite = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"headerBookmark123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"createNamedRange\"", StringComparison.Ordinal));
                Assert.Contains("\"name\":\"HeaderBookmark\"", headerWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesNamedRangesForFootnoteBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterFootnoteBookmark.docx");

            try {
                using var document = BuildGoogleDocsFootnoteBookmarkDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-footnote-bookmark\",\"title\":\"Footnote Bookmark Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footnote-bookmark:batchUpdate") {
                        if (body != null && body.Contains("\"createFootnote\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFootnote\":{\"footnoteId\":\"fn-bookmark-123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Footnote Bookmark Export",
                });

                Assert.Equal("doc-footnote-bookmark", result.DocumentId);
                var footnoteBatchRequest = recordedRequests.Last(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footnote-bookmark:batchUpdate");
                Assert.Contains("\"segmentId\":\"fn-bookmark-123\"", footnoteBatchRequest.Body!);
                Assert.Contains("\"createNamedRange\"", footnoteBatchRequest.Body!);
                Assert.Contains("\"name\":\"FootnoteBookmark\"", footnoteBatchRequest.Body!);
                Assert.Contains("\"text\":\"Footnote text\\n\"", footnoteBatchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysBodyTableFootnotes() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterTableFootnote.docx");

            try {
                using var document = BuildGoogleDocsTableFootnoteDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                int batchUpdateCount = 0;

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-table-footnote\",\"title\":\"Table Footnote Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-footnote:batchUpdate") {
                        batchUpdateCount++;
                        if (body != null && body.Contains("\"createFootnote\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFootnote\":{\"footnoteId\":\"fn-table-123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-footnote") {
                        return CreateJsonResponse(CreateSingleCellBodyTableDocumentStateJson("doc-table-footnote", "Table Footnote Export"));
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Table Footnote Export",
                });

                Assert.Equal("doc-table-footnote", result.DocumentId);
                Assert.Equal(6, recordedRequests.Count);
                Assert.Equal(4, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var tableBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-footnote:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"createFootnote\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"Cell text\\n\"", tableBatchRequest.Body!);

                var footnoteBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-footnote:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"segmentId\":\"fn-table-123\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"Table footnote\\n\"", footnoteBatchRequest.Body!);
                var tableStyleBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-footnote:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"updateTableColumnProperties\"", StringComparison.Ordinal));
                Assert.Contains("\"widthType\":\"FIXED_WIDTH\"", tableStyleBatchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleDocsApiPayloadBuilder_EmitsMergeTableCellsRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsMergedTable.docx");

            try {
                using var document = BuildGoogleDocsMergedTableDocument(filePath);
                var batch = document.CreateGoogleDocsBatch(new GoogleDocsSaveOptions {
                    Title = "Merged Table Export"
                });

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleDocsInsertTableRequest>());
                Assert.Equal(2, tableRequest.Table.Rows[0].Cells[0].ColumnSpan);

                var documentState = JsonSerializer.Deserialize<GoogleDocsApiDocumentResponse>(
                    CreateBodyTableDocumentStateJson("doc-merge-table", "Merged Table Export"));
                Assert.NotNull(documentState);

                var mergePayload = GoogleDocsApiPayloadBuilder.BuildTableMergeBatchUpdatePayload(batch, documentState!);
                var mergeRequest = Assert.Single(mergePayload.Requests, request => request.MergeTableCells != null);
                Assert.Equal(1, mergeRequest.MergeTableCells!.TableRange.TableCellLocation.TableStartLocation.Index);
                Assert.Equal(0, mergeRequest.MergeTableCells.TableRange.TableCellLocation.RowIndex);
                Assert.Equal(0, mergeRequest.MergeTableCells.TableRange.TableCellLocation.ColumnIndex);
                Assert.Equal(1, mergeRequest.MergeTableCells.TableRange.RowSpan);
                Assert.Equal(2, mergeRequest.MergeTableCells.TableRange.ColumnSpan);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysMergedTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterMergedTable.docx");

            try {
                using var document = BuildGoogleDocsMergedTableDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                int batchUpdateCount = 0;

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-merge-table\",\"title\":\"Merged Table Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-merge-table:batchUpdate") {
                        batchUpdateCount++;
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-merge-table") {
                        return CreateJsonResponse(CreateBodyTableDocumentStateJson("doc-merge-table", "Merged Table Export"));
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Merged Table Export",
                });

                Assert.Equal("doc-merge-table", result.DocumentId);
                Assert.Equal(6, recordedRequests.Count);
                Assert.Equal(4, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var mergeBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-merge-table:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"mergeTableCells\"", StringComparison.Ordinal));
                Assert.Contains("\"mergeTableCells\"", mergeBatchRequest.Body!);
                Assert.Contains("\"rowSpan\":1", mergeBatchRequest.Body!);
                Assert.Contains("\"columnSpan\":2", mergeBatchRequest.Body!);
                var tableStyleBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-merge-table:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"updateTableColumnProperties\"", StringComparison.Ordinal));
                Assert.Contains("\"widthType\":\"FIXED_WIDTH\"", tableStyleBatchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_UsesConfiguredHttpPipeline_ForCreateFlow() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterCreate.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            try {
                using var document = BuildGoogleDocsSampleDocument(filePath, imagePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                int batchUpdateCount = 0;

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc123\",\"title\":\"Create Export\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id") {
                        return CreateJsonResponse("{\"id\":\"img123\"}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/img123/permissions?supportsAllDrives=true") {
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc123:batchUpdate") {
                        batchUpdateCount++;
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc123") {
                        return CreateJsonResponse("{\"documentId\":\"doc123\",\"title\":\"Create Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}},{\"startIndex\":20,\"endIndex\":60,\"table\":{\"tableRows\":[{\"tableCells\":[{\"content\":[{\"startIndex\":25,\"endIndex\":26,\"paragraph\":{}}]},{\"content\":[{\"startIndex\":30,\"endIndex\":31,\"paragraph\":{}}]}]},{\"tableCells\":[{\"content\":[{\"startIndex\":35,\"endIndex\":36,\"paragraph\":{}}]},{\"content\":[{\"startIndex\":40,\"endIndex\":41,\"paragraph\":{}}]}]}]}},{\"startIndex\":60,\"endIndex\":80,\"paragraph\":{}}]}}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Create Export",
                });

                Assert.Equal("doc123", result.DocumentId);
                Assert.Equal("https://docs.google.com/document/d/doc123/edit", result.WebViewLink);
                Assert.Equal(7, recordedRequests.Count);
                Assert.Equal(3, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var createRequest = Assert.Single(recordedRequests, request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents");
                using (var json = JsonDocument.Parse(createRequest.Body!)) {
                    Assert.Equal("Create Export", json.RootElement.GetProperty("title").GetString());
                }

                Assert.Contains(recordedRequests, request => request.Uri.AbsoluteUri == "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id");
                Assert.Contains(recordedRequests, request => request.Uri.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/img123/permissions?supportsAllDrives=true");

                var initialBatchRequest = recordedRequests.First(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc123:batchUpdate" && request.Body!.Contains("insertTable", StringComparison.Ordinal));
                Assert.Contains("Intro Bold Portal", initialBatchRequest.Body!);
                Assert.Contains("Image \\n", initialBatchRequest.Body!);
                Assert.Contains("https://example.com/", initialBatchRequest.Body!);
                Assert.Contains("insertInlineImage", initialBatchRequest.Body!);
                using (var initialBatchJson = JsonDocument.Parse(initialBatchRequest.Body!)) {
                    var inlineImageRequest = initialBatchJson.RootElement
                        .GetProperty("requests")
                        .EnumerateArray()
                        .First(request => request.TryGetProperty("insertInlineImage", out _))
                        .GetProperty("insertInlineImage");
                    Assert.Equal("https://drive.google.com/uc?export=download&id=img123", inlineImageRequest.GetProperty("uri").GetString());
                }

                var tableBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc123:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"text\":\"Name\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"Name\\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"text\":\"Value\\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"text\":\"Alpha\\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"text\":\"42\\n\"", tableBatchRequest.Body!);

                var tableStyleBatchRequest = recordedRequests.Single(request =>
                    request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc123:batchUpdate"
                    && request.Body != null
                    && request.Body.Contains("\"pinTableHeaderRows\"", StringComparison.Ordinal));
                Assert.Contains("\"pinnedHeaderRowsCount\":1", tableStyleBatchRequest.Body!);
                Assert.Contains("\"updateTableColumnProperties\"", tableStyleBatchRequest.Body!);
                Assert.Contains("\"widthType\":\"FIXED_WIDTH\"", tableStyleBatchRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesDefaultHeaderAndFooterSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterHeaderFooter.docx");

            try {
                using var document = BuildGoogleDocsHeaderFooterDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-hf\",\"title\":\"Header Footer Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-hf") {
                        return CreateJsonResponse("{\"documentId\":\"doc-hf\",\"title\":\"Header Footer Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-hf:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"header123\"}}]}");
                        }

                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"footer123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Header Footer Export",
                });

                Assert.Equal("doc-hf", result.DocumentId);
                Assert.Equal(7, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var headerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createHeader\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"DEFAULT\"", headerCreate.Body!);

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"DEFAULT\"", footerCreate.Body!);

                var headerWrite = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("header123", StringComparison.Ordinal) && request.Body.Contains("Header text", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"header123\"", headerWrite.Body!);

                var footerWrite = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("footer123", StringComparison.Ordinal) && request.Body.Contains("Footer text", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"footer123\"", footerWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysDefaultHeaderTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterHeaderTable.docx");

            try {
                using var document = BuildGoogleDocsHeaderTableDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-header-table\",\"title\":\"Header Table Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-header-table") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-header-table\",\"title\":\"Header Table Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateHeaderTableDocumentStateJson("headerTable123", "doc-header-table", "Header Table Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-header-table:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"headerTable123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Header Table Export",
                });

                Assert.Equal("doc-header-table", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var headerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createHeader\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"DEFAULT\"", headerCreate.Body!);

                var headerInsertTable = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"headerTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"insertTable\"", StringComparison.Ordinal));
                Assert.Contains("\"rows\":2", headerInsertTable.Body!);
                Assert.Contains("\"columns\":2", headerInsertTable.Body!);

                var headerTableReplay = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"headerTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"H1\\n\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"H4\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"H2\\n\"", headerTableReplay.Body!);
                Assert.Contains("\"text\":\"H3\\n\"", headerTableReplay.Body!);

                var headerTableStyle = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"headerTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"updateTableCellStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"backgroundColor\"", headerTableStyle.Body!);
                Assert.Contains("\"borderRight\"", headerTableStyle.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysDefaultFooterTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterFooterTable.docx");

            try {
                using var document = BuildGoogleDocsFooterTableDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-footer-table\",\"title\":\"Footer Table Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footer-table") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-footer-table\",\"title\":\"Footer Table Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateFooterTableDocumentStateJson("footerTable123", "doc-footer-table", "Footer Table Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-footer-table:batchUpdate") {
                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"footerTable123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Footer Table Export",
                });

                Assert.Equal("doc-footer-table", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"DEFAULT\"", footerCreate.Body!);

                var footerInsertTable = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"footerTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"insertTable\"", StringComparison.Ordinal));
                Assert.Contains("\"rows\":2", footerInsertTable.Body!);
                Assert.Contains("\"columns\":2", footerInsertTable.Body!);

                var footerTableReplay = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"footerTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F1\\n\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F4\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"F2\\n\"", footerTableReplay.Body!);
                Assert.Contains("\"text\":\"F3\\n\"", footerTableReplay.Body!);

                var footerTableStyle = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"footerTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"updateTableCellStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"backgroundColor\"", footerTableStyle.Body!);
                Assert.Contains("\"borderRight\"", footerTableStyle.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysEvenHeaderTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterEvenHeaderTable.docx");

            try {
                using var document = BuildGoogleDocsEvenHeaderTableDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-even-header-table\",\"title\":\"Even Header Table Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-header-table") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-even-header-table\",\"title\":\"Even Header Table Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateHeaderTableDocumentStateJson("evenHeaderTable123", "doc-even-header-table", "Even Header Table Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-header-table:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"evenHeaderTable123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Even Header Table Export",
                });

                Assert.Equal("doc-even-header-table", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatch = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"updateDocumentStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"useEvenPageHeaderFooter\":true", initialBatch.Body!);

                var headerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createHeader\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"EVEN_PAGE\"", headerCreate.Body!);

                var headerInsertTable = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenHeaderTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"insertTable\"", StringComparison.Ordinal));
                Assert.Contains("\"rows\":2", headerInsertTable.Body!);
                Assert.Contains("\"columns\":2", headerInsertTable.Body!);

                var headerTableReplay = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenHeaderTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"H1\\n\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"H4\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"H2\\n\"", headerTableReplay.Body!);
                Assert.Contains("\"text\":\"H3\\n\"", headerTableReplay.Body!);

                var headerTableStyle = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenHeaderTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"updateTableCellStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"backgroundColor\"", headerTableStyle.Body!);
                Assert.Contains("\"borderRight\"", headerTableStyle.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysFirstFooterTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterFirstFooterTable.docx");

            try {
                using var document = BuildGoogleDocsFirstFooterTableDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-first-footer-table\",\"title\":\"First Footer Table Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-first-footer-table") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-first-footer-table\",\"title\":\"First Footer Table Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateFooterTableDocumentStateJson("firstFooterTable123", "doc-first-footer-table", "First Footer Table Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-first-footer-table:batchUpdate") {
                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"firstFooterTable123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "First Footer Table Export",
                });

                Assert.Equal("doc-first-footer-table", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"FIRST_PAGE\"", footerCreate.Body!);

                var footerInsertTable = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"firstFooterTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"insertTable\"", StringComparison.Ordinal));
                Assert.Contains("\"rows\":2", footerInsertTable.Body!);

                var footerTableReplay = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"firstFooterTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F1\\n\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F4\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"F2\\n\"", footerTableReplay.Body!);

                var footerTableStyle = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"firstFooterTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"updateTableCellStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"backgroundColor\"", footerTableStyle.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaysEvenFooterTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterEvenFooterTable.docx");

            try {
                using var document = BuildGoogleDocsEvenFooterTableDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-even-footer-table\",\"title\":\"Even Footer Table Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-footer-table") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-even-footer-table\",\"title\":\"Even Footer Table Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateFooterTableDocumentStateJson("evenFooterTable123", "doc-even-footer-table", "Even Footer Table Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-footer-table:batchUpdate") {
                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"evenFooterTable123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Even Footer Table Export",
                });

                Assert.Equal("doc-even-footer-table", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatch = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"updateDocumentStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"useEvenPageHeaderFooter\":true", initialBatch.Body!);

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"EVEN_PAGE\"", footerCreate.Body!);

                var footerInsertTable = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenFooterTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"insertTable\"", StringComparison.Ordinal));
                Assert.Contains("\"rows\":2", footerInsertTable.Body!);

                var footerTableReplay = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenFooterTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F1\\n\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F4\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"text\":\"F2\\n\"", footerTableReplay.Body!);

                var footerTableStyle = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenFooterTable123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"updateTableCellStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"backgroundColor\"", footerTableStyle.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesFirstPageHeaderAndFooterSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterFirstPageHeaderFooter.docx");

            try {
                using var document = BuildGoogleDocsFirstPageHeaderFooterDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-first-hf\",\"title\":\"First Page Header Footer Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-first-hf") {
                        return CreateJsonResponse("{\"documentId\":\"doc-first-hf\",\"title\":\"First Page Header Footer Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-first-hf:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"header-first-123\"}}]}");
                        }

                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"footer-first-123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "First Page Header Footer Export",
                });

                Assert.Equal("doc-first-hf", result.DocumentId);
                Assert.Equal(7, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var headerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createHeader\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"FIRST_PAGE\"", headerCreate.Body!);

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"FIRST_PAGE\"", footerCreate.Body!);

                var headerWrite = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("header-first-123", StringComparison.Ordinal) && request.Body.Contains("First header text", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"header-first-123\"", headerWrite.Body!);

                var footerWrite = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("footer-first-123", StringComparison.Ordinal) && request.Body.Contains("First footer text", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"footer-first-123\"", footerWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesEvenPageHeaderAndFooterSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterEvenPageHeaderFooter.docx");

            try {
                using var document = BuildGoogleDocsEvenPageHeaderFooterDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-even-hf\",\"title\":\"Even Page Header Footer Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-hf") {
                        return CreateJsonResponse("{\"documentId\":\"doc-even-hf\",\"title\":\"Even Page Header Footer Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-hf:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"header-even-123\"}}]}");
                        }

                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"footer-even-123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Even Page Header Footer Export",
                });

                Assert.Equal("doc-even-hf", result.DocumentId);
                Assert.Equal(7, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatch = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"updateDocumentStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"useEvenPageHeaderFooter\":true", initialBatch.Body!);

                var headerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createHeader\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"EVEN_PAGE\"", headerCreate.Body!);

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"EVEN_PAGE\"", footerCreate.Body!);

                var headerWrite = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("header-even-123", StringComparison.Ordinal) && request.Body.Contains("Even header text", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"header-even-123\"", headerWrite.Body!);

                var footerWrite = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("footer-even-123", StringComparison.Ordinal) && request.Body.Contains("Even footer text", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"footer-even-123\"", footerWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesNamedRangesForHeaderTableBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterHeaderTableBookmark.docx");

            try {
                using var document = BuildGoogleDocsHeaderTableBookmarkDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-header-table-bookmark\",\"title\":\"Header Table Bookmark Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-header-table-bookmark") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-header-table-bookmark\",\"title\":\"Header Table Bookmark Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateHeaderTableDocumentStateJson("headerTableBookmark123", "doc-header-table-bookmark", "Header Table Bookmark Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-header-table-bookmark:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"headerTableBookmark123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Header Table Bookmark Export",
                });

                Assert.Equal("doc-header-table-bookmark", result.DocumentId);
                var headerTableWrite = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"headerTableBookmark123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"createNamedRange\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"name\":\"HeaderCellBookmark\"", StringComparison.Ordinal));
                Assert.Contains("\"insertText\"", headerTableWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesNamedRangesForEvenHeaderTableBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterEvenHeaderTableBookmark.docx");

            try {
                using var document = BuildGoogleDocsEvenHeaderTableBookmarkDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-even-header-table-bookmark\",\"title\":\"Even Header Table Bookmark Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-header-table-bookmark") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-even-header-table-bookmark\",\"title\":\"Even Header Table Bookmark Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateHeaderTableDocumentStateJson("evenHeaderTableBookmark123", "doc-even-header-table-bookmark", "Even Header Table Bookmark Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-header-table-bookmark:batchUpdate") {
                        if (body != null && body.Contains("\"createHeader\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"evenHeaderTableBookmark123\"}}]}");
                        }

                        if (body != null && body.Contains("\"createNamedRange\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createNamedRange\":{\"namedRangeId\":\"nr-even-header-cell-123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Even Header Table Bookmark Export",
                });

                Assert.Equal("doc-even-header-table-bookmark", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatch = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"updateDocumentStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"useEvenPageHeaderFooter\":true", initialBatch.Body!);

                var headerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createHeader\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"EVEN_PAGE\"", headerCreate.Body!);

                var headerTableWrite = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenHeaderTableBookmark123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"H1\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"insertText\"", headerTableWrite.Body!);

                var namedRangeWrite = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenHeaderTableBookmark123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"createNamedRange\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"name\":\"HeaderCellBookmark\"", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"evenHeaderTableBookmark123\"", namedRangeWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesNamedRangesForEvenFooterTableBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsExporterEvenFooterTableBookmark.docx");

            try {
                using var document = BuildGoogleDocsEvenFooterTableBookmarkDocument(filePath);
                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-even-footer-table-bookmark\",\"title\":\"Even Footer Table Bookmark Export\"}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-footer-table-bookmark") {
                        int getCount = recordedRequests.Count(entry => entry.Uri.AbsoluteUri == request.RequestUri.AbsoluteUri && entry.Method == HttpMethod.Get.Method);
                        if (getCount == 1) {
                            return CreateJsonResponse("{\"documentId\":\"doc-even-footer-table-bookmark\",\"title\":\"Even Footer Table Bookmark Export\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":20,\"paragraph\":{}}]}}");
                        }

                        return CreateJsonResponse(CreateFooterTableDocumentStateJson("evenFooterTableBookmark123", "doc-even-footer-table-bookmark", "Even Footer Table Bookmark Export"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-even-footer-table-bookmark:batchUpdate") {
                        if (body != null && body.Contains("\"createFooter\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createFooter\":{\"footerId\":\"evenFooterTableBookmark123\"}}]}");
                        }

                        if (body != null && body.Contains("\"createNamedRange\"", StringComparison.Ordinal)) {
                            return CreateJsonResponse("{\"replies\":[{\"createNamedRange\":{\"namedRangeId\":\"nr-even-footer-cell-123\"}}]}");
                        }

                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Even Footer Table Bookmark Export",
                });

                Assert.Equal("doc-even-footer-table-bookmark", result.DocumentId);
                Assert.Equal(8, recordedRequests.Count);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var initialBatch = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"updateDocumentStyle\"", StringComparison.Ordinal));
                Assert.Contains("\"useEvenPageHeaderFooter\":true", initialBatch.Body!);

                var footerCreate = Assert.Single(recordedRequests, request => request.Body != null && request.Body.Contains("\"createFooter\"", StringComparison.Ordinal));
                Assert.Contains("\"type\":\"EVEN_PAGE\"", footerCreate.Body!);

                var footerTableWrite = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenFooterTableBookmark123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"text\":\"F1\\n\"", StringComparison.Ordinal));
                Assert.Contains("\"insertText\"", footerTableWrite.Body!);

                var namedRangeWrite = Assert.Single(recordedRequests, request =>
                    request.Body != null
                    && request.Body.Contains("\"segmentId\":\"evenFooterTableBookmark123\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"createNamedRange\"", StringComparison.Ordinal)
                    && request.Body.Contains("\"name\":\"FooterCellBookmark\"", StringComparison.Ordinal));
                Assert.Contains("\"segmentId\":\"evenFooterTableBookmark123\"", namedRangeWrite.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private WordDocument BuildGoogleDocsSampleDocument(string filePath, string imagePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Snapshot";

            var intro = document.AddParagraph("Intro ");
            intro.Style = WordParagraphStyles.Heading1;
            intro.ParagraphAlignment = JustificationValues.Center;
            intro.AddFormattedText("Bold", bold: true);
            intro.AddHyperLink(" Portal", new Uri("https://example.com"));

            var imageParagraph = document.AddParagraph("Image ");
            var image = imageParagraph.InsertImage(imagePath, width: 2, height: 1, description: "Logo");
            image.Title = "Brand";

            var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            table.Title = "Summary";
            table.Description = "Demo table";
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Alpha";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "42";

            document.AddSection(SectionMarkValues.NextPage);
            document.AddParagraph("Second section");

            return document;
        }

        private WordDocument BuildGoogleDocsListDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Lists";

            var bullets = document.AddList(WordListStyle.Bulleted);
            bullets.AddItem("First bullet");
            bullets.AddItem("Nested bullet", level: 1);

            var numbered = document.AddList(WordListStyle.Numbered);
            numbered.AddItem("First step");
            numbered.AddItem("Nested step", level: 1);

            return document;
        }

        private WordDocument BuildGoogleDocsHighlightDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Highlights";

            var paragraph = document.AddParagraph("Plain ");
            paragraph.AddText("Highlighted").SetHighlight(HighlightColorValues.Yellow);
            paragraph.AddText(" tail");

            return document;
        }

        private WordDocument BuildGoogleDocsBaselineDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Baseline";

            var paragraph = document.AddParagraph("Base ");
            paragraph.AddText("Super").SetVerticalTextAlignment(VerticalPositionValues.Superscript);
            paragraph.AddText(" and ");
            paragraph.AddText("Sub").SetVerticalTextAlignment(VerticalPositionValues.Subscript);

            return document;
        }

        private WordDocument BuildGoogleDocsSmallCapsDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs SmallCaps";

            var paragraph = document.AddParagraph("Base ");
            paragraph.AddText("Small").SetSmallCaps();
            paragraph.AddText(" and ");
            paragraph.AddText("Caps").SetCapsStyle(CapsStyle.Caps);

            return document;
        }

        private WordDocument BuildGoogleDocsFontFamilyDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Font Family";

            var paragraph = document.AddParagraph();
            paragraph.AddText("Mono").SetFontFamily("Consolas");
            paragraph.AddText(" plain");

            return document;
        }

        private WordDocument BuildGoogleDocsParagraphLayoutDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Paragraph Layout";

            var paragraph = document.AddParagraph("Indented paragraph");
            paragraph.IndentationBeforePoints = 24;
            paragraph.IndentationAfterPoints = 12;
            paragraph.IndentationFirstLinePoints = 18;
            paragraph.LineSpacingBeforePoints = 6;
            paragraph.LineSpacingAfterPoints = 9;

            return document;
        }

        private WordDocument BuildGoogleDocsRightToLeftParagraphDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs RTL Paragraph";

            var paragraph = document.AddParagraph("RTL paragraph");
            paragraph.BiDi = true;

            return document;
        }

        private WordDocument BuildGoogleDocsAutoLineSpacingDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Auto Line Spacing";

            var paragraph = document.AddParagraph("Auto spaced paragraph");
            paragraph.LineSpacing = 360;
            paragraph.LineSpacingRule = LineSpacingRuleValues.Auto;

            return document;
        }

        private WordDocument BuildGoogleDocsParagraphPaginationDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Paragraph Pagination";

            var paragraph = document.AddParagraph("Paragraph with pagination controls");
            paragraph.KeepWithNext = true;
            paragraph.KeepLinesTogether = true;
            paragraph.AvoidWidowAndOrphan = true;

            return document;
        }

        private WordDocument BuildGoogleDocsParagraphTabStopsDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Paragraph Tab Stops";

            var paragraph = document.AddParagraph("Label");
            paragraph.AddTabStop(1440, TabStopValues.Left, TabStopLeaderCharValues.None);
            paragraph.AddText("\tValue");
            paragraph.AddTabStop(2880, TabStopValues.Decimal, TabStopLeaderCharValues.Dot);
            paragraph.AddText("\t12.34");

            return document;
        }

        private WordDocument BuildGoogleDocsSectionLayoutDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Section Layout";

            var section = document.Sections[0];
            section.PageSettings.PageSize = WordPageSize.A5;
            section.PageOrientation = PageOrientationValues.Landscape;
            section.Margins.Top = 720;
            section.Margins.Bottom = 900;
            section.Margins.Left = 1080U;
            section.Margins.Right = 1260U;
            section.Margins.HeaderDistance = 360U;
            section.Margins.FooterDistance = 540U;
            section.DifferentFirstPage = true;
            section.PageNumberType.Start = 3;

            document.AddParagraph("Section layout paragraph");
            return document;
        }

        private WordDocument BuildGoogleDocsSectionColumnsDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Section Columns";

            var section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 360;
            section.HasColumnSeparator = true;

            document.AddParagraph("Section columns paragraph");
            return document;
        }

        private WordDocument BuildGoogleDocsExactLineSpacingDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Exact Line Spacing";

            var paragraph = document.AddParagraph("Exact spaced paragraph");
            paragraph.LineSpacing = 360;
            paragraph.LineSpacingRule = LineSpacingRuleValues.Exact;

            return document;
        }

        private WordDocument BuildGoogleDocsAtLeastLineSpacingDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs AtLeast Line Spacing";

            var paragraph = document.AddParagraph("AtLeast spaced paragraph");
            paragraph.LineSpacing = 360;
            paragraph.LineSpacingRule = LineSpacingRuleValues.AtLeast;

            return document;
        }

        private WordDocument BuildGoogleDocsParagraphShadingDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Paragraph Shading";

            var paragraph = document.AddParagraph("Shaded paragraph");
            paragraph.ShadingFillColorHex = "D9EAF7";

            return document;
        }

        private WordDocument BuildGoogleDocsParagraphBorderDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Paragraph Borders";

            var paragraph = document.AddParagraph("Bordered paragraph");
            paragraph.Borders.TopStyle = BorderValues.Single;
            paragraph.Borders.TopColorHex = "336699";
            paragraph.Borders.TopSize = 8;
            paragraph.Borders.TopSpace = 6;
            paragraph.Borders.LeftStyle = BorderValues.Single;
            paragraph.Borders.LeftColorHex = "CC3300";
            paragraph.Borders.LeftSize = 12;
            paragraph.Borders.LeftSpace = 4;

            return document;
        }

        private WordDocument BuildGoogleDocsPageBreakDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Page Break";

            document.AddParagraph("Intro paragraph");
            var nextPageParagraph = document.AddParagraph("Starts on next page");
            nextPageParagraph.PageBreakBefore = true;

            return document;
        }

        private WordDocument BuildGoogleDocsFootnoteDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Footnotes";
            document.AddParagraph("Body text").AddFootNote("Footnote text");
            return document;
        }

        private WordDocument BuildGoogleDocsFootnoteBookmarkDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Footnote Bookmarks";
            var footnoteReferenceParagraph = document.AddParagraph("Body text").AddFootNote("Footnote text");
            footnoteReferenceParagraph.FootNote!.Paragraphs!.Last().AddBookmark("FootnoteBookmark");
            return document;
        }

        private WordDocument BuildGoogleDocsBookmarkDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Bookmarks";
            document.AddParagraph("Jump to ").AddHyperLink("target", "TargetBookmark");
            document.AddParagraph("Target paragraph").AddBookmark("TargetBookmark");
            return document;
        }

        private WordDocument BuildGoogleDocsTableBookmarkDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Table Bookmark";

            var table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Bookmarked cell";
            table.Rows[0].Cells[0].Paragraphs[0].AddBookmark("CellBookmark");

            return document;
        }

        private WordDocument BuildGoogleDocsHeaderBookmarkDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Header Bookmark";
            document.AddParagraph("Body text");
            document.AddHeadersAndFooters();
            document.Sections[0].Header.Default!.AddParagraph("Header bookmark").AddBookmark("HeaderBookmark");
            return document;
        }

        private WordDocument BuildGoogleDocsHeaderFooterDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Header Footer";
            document.AddParagraph("Body text");
            document.AddHeadersAndFooters();
            document.Sections[0].Header.Default!.AddParagraph("Header text");
            document.Sections[0].Footer.Default!.AddParagraph("Footer text");
            return document;
        }

        private WordDocument BuildGoogleDocsFirstPageHeaderFooterDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs First Page Header Footer";
            document.AddParagraph("Body text");
            document.Sections[0].DifferentFirstPage = true;
            document.Sections[0].Header.First!.AddParagraph("First header text");
            document.Sections[0].Footer.First!.AddParagraph("First footer text");
            return document;
        }

        private WordDocument BuildGoogleDocsEvenPageHeaderFooterDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Even Page Header Footer";
            document.AddParagraph("Body text");
            document.Sections[0].DifferentOddAndEvenPages = true;
            document.Sections[0].Header.Even!.AddParagraph("Even header text");
            document.Sections[0].Footer.Even!.AddParagraph("Even footer text");
            return document;
        }

        private WordDocument BuildGoogleDocsStyledTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Styled Table";

            var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
            table.ColumnWidth = new List<int> { 1440, 2880 };
            var firstCellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
            firstCellParagraph.Text = "Cell ";
            firstCellParagraph.AddFormattedText("Bold", bold: true);
            firstCellParagraph.AddHyperLink(" Link", new Uri("https://example.com"));
            table.Rows[0].Cells[0].ShadingFillColorHex = "FFCC00";
            table.Rows[0].Cells[0].Borders.LeftStyle = BorderValues.Single;
            table.Rows[0].Cells[0].Borders.LeftColorHex = "ff0000";
            table.Rows[0].Cells[0].Borders.LeftSize = (UInt32Value)8U;
            table.Rows[0].Cells[0].Borders.TopStyle = BorderValues.Dashed;
            table.Rows[0].Cells[0].Borders.TopColorHex = "0000ff";
            table.Rows[0].Cells[0].Borders.TopSize = (UInt32Value)12U;
            table.Rows[0].Cells[0].Borders.BottomStyle = BorderValues.Dotted;
            table.Rows[0].Cells[0].Borders.BottomColorHex = "00aa00";
            table.Rows[0].Cells[0].Borders.BottomSize = (UInt32Value)16U;

            table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Alpha";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "42";

            return document;
        }

        private WordDocument BuildGoogleDocsTableImageDocument(string filePath, string imagePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Table Image";

            var table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            var paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Cell ";
            var image = paragraph.InsertImage(imagePath, width: 1, height: 1, description: "CellLogo");
            image.Title = "CellBrand";

            return document;
        }

        private WordDocument BuildGoogleDocsTableFootnoteDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Table Footnote";

            var table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            var paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Cell text";
            paragraph.AddFootNote("Table footnote");

            return document;
        }

        private WordDocument BuildGoogleDocsMergedTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Merged Table";

            var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Merged";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Hidden";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Alpha";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "42";
            table.Rows[0].Cells[0].MergeHorizontally(1);

            return document;
        }

        private WordDocument BuildGoogleDocsLeadingTableSectionDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Leading Table Section";
            document.AddParagraph("First section");
            document.AddSection(SectionMarkValues.NextPage);

            var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Alpha";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "42";

            return document;
        }

        private WordDocument BuildGoogleDocsHeaderTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Header Table";
            document.AddParagraph("Body text");
            document.AddHeadersAndFooters();

            var headerTable = document.Sections[0].Header.Default!.AddTable(2, 2, WordTableStyle.TableGrid);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            headerTable.Rows[0].Cells[0].ShadingFillColorHex = "D9EAF7";
            headerTable.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            headerTable.Rows[0].Cells[0].Borders.RightColorHex = "336699";
            headerTable.Rows[0].Cells[0].Borders.RightSize = (UInt32Value)8U;
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            headerTable.Rows[1].Cells[0].Paragraphs[0].Text = "H3";
            headerTable.Rows[1].Cells[1].Paragraphs[0].Text = "H4";

            return document;
        }

        private WordDocument BuildGoogleDocsHeaderTableBookmarkDocument(string filePath) {
            var document = BuildGoogleDocsHeaderTableDocument(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Header Table Bookmark";
            document.Sections[0].Header.Default!.Tables[0].Rows[0].Cells[0].Paragraphs[0].AddBookmark("HeaderCellBookmark");
            return document;
        }

        private WordDocument BuildGoogleDocsFooterTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Footer Table";
            document.AddParagraph("Body text");
            document.AddHeadersAndFooters();

            var footerTable = document.Sections[0].Footer.Default!.AddTable(2, 2, WordTableStyle.TableGrid);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "F1";
            footerTable.Rows[0].Cells[0].ShadingFillColorHex = "D9EAF7";
            footerTable.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            footerTable.Rows[0].Cells[0].Borders.RightColorHex = "336699";
            footerTable.Rows[0].Cells[0].Borders.RightSize = (UInt32Value)8U;
            footerTable.Rows[0].Cells[1].Paragraphs[0].Text = "F2";
            footerTable.Rows[1].Cells[0].Paragraphs[0].Text = "F3";
            footerTable.Rows[1].Cells[1].Paragraphs[0].Text = "F4";

            return document;
        }

        private WordDocument BuildGoogleDocsFirstFooterTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs First Footer Table";
            document.AddParagraph("Body text");
            document.Sections[0].DifferentFirstPage = true;

            var footerTable = document.Sections[0].Footer.First!.AddTable(2, 2, WordTableStyle.TableGrid);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "F1";
            footerTable.Rows[0].Cells[0].ShadingFillColorHex = "D9EAF7";
            footerTable.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            footerTable.Rows[0].Cells[0].Borders.RightColorHex = "336699";
            footerTable.Rows[0].Cells[0].Borders.RightSize = (UInt32Value)8U;
            footerTable.Rows[0].Cells[1].Paragraphs[0].Text = "F2";
            footerTable.Rows[1].Cells[0].Paragraphs[0].Text = "F3";
            footerTable.Rows[1].Cells[1].Paragraphs[0].Text = "F4";

            return document;
        }

        private WordDocument BuildGoogleDocsEvenFooterTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Even Footer Table";
            document.AddParagraph("Body text");
            document.Sections[0].DifferentOddAndEvenPages = true;

            var footerTable = document.Sections[0].Footer.Even!.AddTable(2, 2, WordTableStyle.TableGrid);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "F1";
            footerTable.Rows[0].Cells[0].ShadingFillColorHex = "D9EAF7";
            footerTable.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            footerTable.Rows[0].Cells[0].Borders.RightColorHex = "336699";
            footerTable.Rows[0].Cells[0].Borders.RightSize = (UInt32Value)8U;
            footerTable.Rows[0].Cells[1].Paragraphs[0].Text = "F2";
            footerTable.Rows[1].Cells[0].Paragraphs[0].Text = "F3";
            footerTable.Rows[1].Cells[1].Paragraphs[0].Text = "F4";

            return document;
        }

        private WordDocument BuildGoogleDocsEvenFooterTableBookmarkDocument(string filePath) {
            var document = BuildGoogleDocsEvenFooterTableDocument(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Even Footer Table Bookmark";
            document.Sections[0].Footer.Even!.Tables[0].Rows[0].Cells[0].Paragraphs[0].AddBookmark("FooterCellBookmark");
            return document;
        }

        private WordDocument BuildGoogleDocsEvenHeaderTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Even Header Table";
            document.AddParagraph("Body text");
            document.Sections[0].DifferentOddAndEvenPages = true;

            var headerTable = document.Sections[0].Header.Even!.AddTable(2, 2, WordTableStyle.TableGrid);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            headerTable.Rows[0].Cells[0].ShadingFillColorHex = "D9EAF7";
            headerTable.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            headerTable.Rows[0].Cells[0].Borders.RightColorHex = "336699";
            headerTable.Rows[0].Cells[0].Borders.RightSize = (UInt32Value)8U;
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            headerTable.Rows[1].Cells[0].Paragraphs[0].Text = "H3";
            headerTable.Rows[1].Cells[1].Paragraphs[0].Text = "H4";

            return document;
        }

        private WordDocument BuildGoogleDocsEvenHeaderTableBookmarkDocument(string filePath) {
            var document = BuildGoogleDocsEvenHeaderTableDocument(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Even Header Table Bookmark";
            document.Sections[0].Header.Even!.Tables[0].Rows[0].Cells[0].Paragraphs[0].AddBookmark("HeaderCellBookmark");
            return document;
        }

        private static string CreateHeaderTableDocumentStateJson(
            string headerId,
            string documentId = "doc-header-table",
            string title = "Header Table Export") {
            return JsonSerializer.Serialize(new {
                documentId,
                title,
                body = new {
                    content = new object[] {
                        new {
                            startIndex = 1,
                            endIndex = 20,
                            paragraph = new { }
                        }
                    }
                },
                headers = new Dictionary<string, object> {
                    [headerId] = new {
                        content = new object[] {
                            new {
                                startIndex = 1,
                                endIndex = 40,
                                table = new {
                                    tableRows = new object[] {
                                        new {
                                            tableCells = new object[] {
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 5,
                                                            endIndex = 6,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                },
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 10,
                                                            endIndex = 11,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                }
                                            }
                                        },
                                        new {
                                            tableCells = new object[] {
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 15,
                                                            endIndex = 16,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                },
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 20,
                                                            endIndex = 21,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private static string CreateFooterTableDocumentStateJson(
            string footerId,
            string documentId = "doc-footer-table",
            string title = "Footer Table Export") {
            return JsonSerializer.Serialize(new {
                documentId,
                title,
                body = new {
                    content = new object[] {
                        new {
                            startIndex = 1,
                            endIndex = 20,
                            paragraph = new { }
                        }
                    }
                },
                footers = new Dictionary<string, object> {
                    [footerId] = new {
                        content = new object[] {
                            new {
                                startIndex = 1,
                                endIndex = 40,
                                table = new {
                                    tableRows = new object[] {
                                        new {
                                            tableCells = new object[] {
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 5,
                                                            endIndex = 6,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                },
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 10,
                                                            endIndex = 11,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                }
                                            }
                                        },
                                        new {
                                            tableCells = new object[] {
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 15,
                                                            endIndex = 16,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                },
                                                new {
                                                    content = new object[] {
                                                        new {
                                                            startIndex = 20,
                                                            endIndex = 21,
                                                            paragraph = new { }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private static string CreateBodyTableDocumentStateJson(
            string documentId = "doc-table-style",
            string title = "Styled Table Cell Export") {
            return JsonSerializer.Serialize(new {
                documentId,
                title,
                body = new {
                    content = new object[] {
                        new {
                            startIndex = 1,
                            endIndex = 60,
                            table = new {
                                tableRows = new object[] {
                                    new {
                                        tableCells = new object[] {
                                            new {
                                                content = new object[] {
                                                    new {
                                                        startIndex = 25,
                                                        endIndex = 26,
                                                        paragraph = new { }
                                                    }
                                                }
                                            },
                                            new {
                                                content = new object[] {
                                                    new {
                                                        startIndex = 30,
                                                        endIndex = 31,
                                                        paragraph = new { }
                                                    }
                                                }
                                            }
                                        }
                                    },
                                    new {
                                        tableCells = new object[] {
                                            new {
                                                content = new object[] {
                                                    new {
                                                        startIndex = 35,
                                                        endIndex = 36,
                                                        paragraph = new { }
                                                    }
                                                }
                                            },
                                            new {
                                                content = new object[] {
                                                    new {
                                                        startIndex = 40,
                                                        endIndex = 41,
                                                        paragraph = new { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private static string CreateSingleCellBodyTableDocumentStateJson(
            string documentId = "doc-table-footnote",
            string title = "Table Footnote Export") {
            return JsonSerializer.Serialize(new {
                documentId,
                title,
                body = new {
                    content = new object[] {
                        new {
                            startIndex = 1,
                            endIndex = 30,
                            table = new {
                                tableRows = new object[] {
                                    new {
                                        tableCells = new object[] {
                                            new {
                                                content = new object[] {
                                                    new {
                                                        startIndex = 11,
                                                        endIndex = 12,
                                                        paragraph = new { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private static HttpResponseMessage CreateJsonResponse(string json) {
            return new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private sealed class FakeGoogleWorkspaceCredentialSource : IGoogleWorkspaceCredentialSource {
            public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(IEnumerable<string> scopes, CancellationToken cancellationToken = default) {
                return Task.FromResult(new GoogleWorkspaceAccessToken(
                    "fake-access-token",
                    DateTimeOffset.UtcNow.AddHours(1),
                    scopes.ToList()));
            }
        }

        private sealed class FakeHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public FakeHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
            }
        }
    }
}
