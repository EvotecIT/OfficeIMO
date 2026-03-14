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
                Assert.Equal(6, recordedRequests.Count);
                Assert.Equal(2, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                Assert.Contains(recordedRequests, request => request.Uri.AbsoluteUri == "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id");
                Assert.Contains(recordedRequests, request => request.Uri.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/img-table-123/permissions?supportsAllDrives=true");

                var tableBatchRequest = recordedRequests.Last(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-table-image:batchUpdate");
                Assert.Contains("\"text\":\"Cell \\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"insertInlineImage\"", tableBatchRequest.Body!);
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
                Assert.Equal(5, recordedRequests.Count);
                Assert.Equal(3, batchUpdateCount);
                Assert.All(recordedRequests, request => Assert.Equal("Bearer fake-access-token", request.Authorization));

                var mergeBatchRequest = recordedRequests.Last(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-merge-table:batchUpdate");
                Assert.Contains("\"mergeTableCells\"", mergeBatchRequest.Body!);
                Assert.Contains("\"rowSpan\":1", mergeBatchRequest.Body!);
                Assert.Contains("\"columnSpan\":2", mergeBatchRequest.Body!);
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
                Assert.Equal(6, recordedRequests.Count);
                Assert.Equal(2, batchUpdateCount);
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

                var tableBatchRequest = recordedRequests.Last(request => request.Uri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc123:batchUpdate");
                Assert.Contains("\"text\":\"Name\\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"text\":\"Value\\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"text\":\"Alpha\\n\"", tableBatchRequest.Body!);
                Assert.Contains("\"text\":\"42\\n\"", tableBatchRequest.Body!);
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
                Assert.Equal(7, recordedRequests.Count);
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

        private WordDocument BuildGoogleDocsHeaderFooterDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Header Footer";
            document.AddParagraph("Body text");
            document.AddHeadersAndFooters();
            document.Sections[0].Header.Default!.AddParagraph("Header text");
            document.Sections[0].Footer.Default!.AddParagraph("Footer text");
            return document;
        }

        private WordDocument BuildGoogleDocsStyledTableDocument(string filePath) {
            var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Google Docs Styled Table";

            var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            var firstCellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
            firstCellParagraph.Text = "Cell ";
            firstCellParagraph.AddFormattedText("Bold", bold: true);
            firstCellParagraph.AddHyperLink(" Link", new Uri("https://example.com"));

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
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            headerTable.Rows[1].Cells[0].Paragraphs[0].Text = "H3";
            headerTable.Rows[1].Cells[1].Paragraphs[0].Text = "H4";

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
