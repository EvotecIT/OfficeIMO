using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>Imports Google documents through Drive conversion or a tab-aware native projection.</summary>
    public sealed class GoogleDocsImporter : IGoogleDocsImporter {
        public async Task<GoogleDocsImportResult> ImportAsync(
            string documentId,
            GoogleWorkspaceSession session,
            GoogleDocsImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(documentId)) throw new ArgumentException("Document ID is required.", nameof(documentId));
            if (session == null) throw new ArgumentNullException(nameof(session));
            GoogleDocsImportOptions effective = options ?? new GoogleDocsImportOptions();
            ValidateOptions(effective);
            return effective.Mode == GoogleDocsImportMode.DriveExport
                ? await ImportDriveAsync(documentId, session, effective, cancellationToken).ConfigureAwait(false)
                : await ImportNativeAsync(documentId, session, effective, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<GoogleDocsImportResult> ImportDriveAsync(
            string documentId,
            GoogleWorkspaceSession session,
            GoogleDocsImportOptions options,
            CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(documentId, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsureDocument(source, documentId);
            EnsureDownloadable(source, documentId);
            byte[] bytes = await drive.ExportAsync(documentId, GoogleDriveMimeTypes.MicrosoftWord,
                options.Progress, report, cancellationToken, options.MaxResponseBytes).ConfigureAwait(false);
            var stream = new MemoryStream(bytes, writable: true);
            WordDocument document;
            try {
                document = WordDocument.Load(stream, options.LoadOptions);
            } catch {
                stream.Dispose();
                throw;
            }
            report.Add(TranslationSeverity.Info, "DriveExportImport", "The Google document was exported to DOCX through Drive and loaded by OfficeIMO.",
                code: "DOCS.IMPORT.DRIVE_EXPORT", action: TranslationAction.Preserve, targetId: documentId);
            return new GoogleDocsImportResult(document, BuildReference(source, documentId, report), report);
        }

        private static async Task<GoogleDocsImportResult> ImportNativeAsync(
            string documentId,
            GoogleWorkspaceSession session,
            GoogleDocsImportOptions options,
            CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(documentId, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsureDocument(source, documentId);
            EnsureDownloadable(source, documentId);
            GoogleWorkspaceAccessToken token = await session.AcquireAccessTokenAsync(new[] { GoogleWorkspaceScopeCatalog.DocumentsReadonly }, cancellationToken).ConfigureAwait(false);
            string uri = $"https://docs.googleapis.com/v1/documents/{Uri.EscapeDataString(documentId)}?includeTabsContent=true&suggestionsViewMode={MapSuggestions(options.Suggestions)}";
            GoogleDocsApiDocumentResponse response;
            using (var transport = new GoogleWorkspaceHttpTransport(session.Options)) {
                response = await transport.SendJsonAsync<GoogleDocsApiDocumentResponse>(token.AccessToken, HttpMethod.Get, uri, null,
                    GoogleWorkspaceRequestSafety.Safe, "Google Docs API", report,
                    GoogleDocsJsonSerializerContext.Default.GoogleDocsApiDocumentResponse, cancellationToken,
                    options.MaxResponseBytes).ConfigureAwait(false);
            }

            ValidateNativeResponse(response, options);
            WordDocument document = Project(response, options, report);
            report.Add(TranslationSeverity.Info, "NativeImport", "Tab-aware text, core run styles, headings, lists, and simple tables were projected into OfficeIMO.",
                code: "DOCS.IMPORT.NATIVE", action: TranslationAction.Preserve, targetId: documentId);
            return new GoogleDocsImportResult(document, BuildReference(source, documentId, report, response), report);
        }

        private static WordDocument Project(GoogleDocsApiDocumentResponse response, GoogleDocsImportOptions options, TranslationReport report) {
            var stream = new MemoryStream();
            WordDocument document = WordDocument.Create(stream);
            try {
                document.BuiltinDocumentProperties.Title = response.Title;
                IReadOnlyList<GoogleDocsApiTabResponse> tabs = SelectImportTabs(response, options);
                bool addHeadings = options.TabMode == GoogleDocsImportTabMode.FlattenWithHeadings && tabs.Count > 1;
                foreach (GoogleDocsApiTabResponse tab in tabs) {
                    if (addHeadings) {
                        WordParagraph heading = document.AddParagraph(tab.Properties.Title ?? "Tab");
                        heading.Style = WordParagraphStyles.Heading1;
                    }
                    ProjectContent(document, tab.DocumentTab?.Body?.Content, report, tab.Properties.TabId);
                    ProjectSegments(document, tab, report);
                }
                document.Save();
                return document;
            } catch {
                document.Dispose();
                stream.Dispose();
                throw;
            }
        }

        private static void ProjectContent(
            WordDocument document,
            IReadOnlyList<GoogleDocsApiStructuralElementResponse>? content,
            TranslationReport report,
            string? tabId) {
            if (content == null) return;
            foreach (GoogleDocsApiStructuralElementResponse element in content) {
                if (element.Paragraph != null) {
                    ProjectParagraph(document, element.Paragraph, report, tabId);
                } else if (element.Table != null) {
                    ProjectTable(document, element.Table, report, tabId);
                }
            }
        }

        private static void ProjectParagraph(WordDocument document, GoogleDocsApiParagraphElementResponse source, TranslationReport report, string? tabId) {
            WordParagraph paragraph = document.AddParagraph();
            if (source.Bullet != null) paragraph.AddText(new string(' ', Math.Max(0, source.Bullet.NestingLevel) * 2) + "• ");
            ApplyParagraphStyle(paragraph, source.ParagraphStyle);
            foreach (GoogleDocsApiParagraphInlineElementResponse element in source.Elements) {
                if (element.TextRun?.Content is string text) {
                    text = text.EndsWith("\n", StringComparison.Ordinal) ? text.Substring(0, text.Length - 1) : text;
                    if (text.Length == 0) continue;
                    WordParagraph run = paragraph.AddText(text);
                    ApplyTextStyle(run, element.TextRun.TextStyle, report, tabId);
                } else if (element.InlineObjectElement != null) {
                    paragraph.AddText("[Google Docs inline object]");
                    report.AddUnique(TranslationSeverity.Warning, "InlineObjects", "Native inline objects require Drive-export import for binary preservation.",
                        path: tabId ?? string.Empty, code: "DOCS.IMPORT.INLINE_OBJECT_FALLBACK", action: TranslationAction.Flatten);
                } else if (element.FootnoteReference != null) {
                    paragraph.AddText("[footnote]");
                }
            }
        }

        private static void ProjectTable(WordDocument document, GoogleDocsApiTableResponse source, TranslationReport report, string? tabId) {
            int rows = source.Rows.Count;
            int columns = source.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(0).Max();
            if (rows == 0 || columns == 0) return;
            WordTable table = document.AddTable(rows, columns, WordTableStyle.TableGrid);
            for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
                for (int columnIndex = 0; columnIndex < source.Rows[rowIndex].Cells.Count; columnIndex++) {
                    string text = ExtractPlainText(source.Rows[rowIndex].Cells[columnIndex].Content);
                    table.Rows[rowIndex].Cells[columnIndex].AddParagraph(text, removeExistingParagraphs: true);
                }
            }
            report.AddUnique(TranslationSeverity.Info, "Tables", "Native simple table cells were projected; merged-cell and advanced table styling use the Drive-export fallback.",
                path: tabId ?? string.Empty, code: "DOCS.IMPORT.TABLE.PARTIAL", action: TranslationAction.Preserve);
        }

        private static string ExtractPlainText(IReadOnlyList<GoogleDocsApiStructuralElementResponse>? content) {
            if (content == null) return string.Empty;
            var text = new StringBuilder();
            foreach (GoogleDocsApiStructuralElementResponse element in content) {
                if (element.Paragraph != null) {
                    foreach (GoogleDocsApiParagraphInlineElementResponse inline in element.Paragraph.Elements) text.Append(inline.TextRun?.Content);
                }
            }
            return text.ToString().TrimEnd('\r', '\n');
        }

        private static void ApplyParagraphStyle(WordParagraph paragraph, GoogleDocsApiParagraphStylePayload? style) {
            if (style == null) return;
            paragraph.Style = style.NamedStyleType switch {
                "TITLE" => WordParagraphStyles.Heading1,
                "SUBTITLE" => WordParagraphStyles.Heading2,
                "HEADING_1" => WordParagraphStyles.Heading1,
                "HEADING_2" => WordParagraphStyles.Heading2,
                "HEADING_3" => WordParagraphStyles.Heading3,
                "HEADING_4" => WordParagraphStyles.Heading4,
                "HEADING_5" => WordParagraphStyles.Heading5,
                "HEADING_6" => WordParagraphStyles.Heading6,
                _ => paragraph.Style,
            };
            paragraph.ParagraphAlignment = style.Alignment switch {
                "CENTER" => JustificationValues.Center,
                "END" => JustificationValues.Right,
                "JUSTIFIED" => JustificationValues.Both,
                _ => paragraph.ParagraphAlignment,
            };
        }

        private static void ApplyTextStyle(WordParagraph run, GoogleDocsApiTextStylePayload? style, TranslationReport report, string? tabId) {
            if (style == null) return;
            if (style.Bold == true) run.SetBold();
            if (style.Italic == true) run.SetItalic();
            if (style.Underline == true) run.SetUnderline(UnderlineValues.Single);
            if (style.Strikethrough == true) run.SetStrike();
            if (style.FontSize?.Magnitude is double size && size > 0) run.SetFontSize((int)Math.Round(size));
            if (!string.IsNullOrWhiteSpace(style.WeightedFontFamily?.FontFamily)) run.SetFontFamily(style.WeightedFontFamily!.FontFamily);
            if (style.ForegroundColor?.Color.RgbColor is GoogleDocsApiRgbColorPayload color) run.SetColorHex(ToHex(color));
            if (style.SmallCaps == true) run.CapsStyle = CapsStyle.SmallCaps;
            if (style.Link != null) {
                report.AddUnique(TranslationSeverity.Warning, "Links", "Native links are detected, but exact tab/bookmark hyperlink reconstruction currently uses the Drive-export fallback.",
                    path: tabId ?? string.Empty, code: "DOCS.IMPORT.LINK_FALLBACK", action: TranslationAction.Flatten);
            }
        }

        private static void ProjectSegments(WordDocument document, GoogleDocsApiTabResponse tab, TranslationReport report) {
            int count = (tab.DocumentTab?.Headers?.Count ?? 0) + (tab.DocumentTab?.Footers?.Count ?? 0) + (tab.DocumentTab?.Footnotes?.Count ?? 0);
            if (count > 0) report.Add(TranslationSeverity.Warning, "Segments", $"Native import detected {count} header/footer/footnote segment(s); Drive-export import preserves their exact placement.",
                path: tab.Properties.TabId ?? string.Empty, code: "DOCS.IMPORT.SEGMENT_FALLBACK", action: TranslationAction.Flatten, count: count);
        }

        private static IReadOnlyList<GoogleDocsApiTabResponse> SelectImportTabs(GoogleDocsApiDocumentResponse response, GoogleDocsImportOptions options) {
            var all = GoogleDocsApiPayloadBuilder.FlattenTabs(response.Tabs).ToList();
            if (all.Count == 0) {
                return new[] { new GoogleDocsApiTabResponse { Properties = new GoogleDocsApiTabPropertiesResponse { Title = response.Title }, DocumentTab = new GoogleDocsApiDocumentTabResponse { Body = response.Body, Headers = response.Headers, Footers = response.Footers, Footnotes = response.Footnotes } } };
            }
            return options.TabMode switch {
                GoogleDocsImportTabMode.FirstTab => new[] { all[0] },
                GoogleDocsImportTabMode.SelectedTab => new[] { all.FirstOrDefault(tab => string.Equals(tab.Properties.TabId, options.TabId, StringComparison.Ordinal))
                    ?? throw new InvalidOperationException($"Google document does not contain tab '{options.TabId}'.") },
                _ => all,
            };
        }

        private static string MapSuggestions(GoogleDocsSuggestionsMode mode) => mode switch {
            GoogleDocsSuggestionsMode.Default => "DEFAULT_FOR_CURRENT_ACCESS",
            GoogleDocsSuggestionsMode.Accepted => "PREVIEW_SUGGESTIONS_ACCEPTED",
            GoogleDocsSuggestionsMode.Inline => "SUGGESTIONS_INLINE",
            GoogleDocsSuggestionsMode.Rejected => "PREVIEW_WITHOUT_SUGGESTIONS",
            _ => throw new ArgumentOutOfRangeException(nameof(mode), mode, "Unsupported Google Docs suggestions mode."),
        };

        private static string ToHex(GoogleDocsApiRgbColorPayload color) => $"{ToByte(color.Red):X2}{ToByte(color.Green):X2}{ToByte(color.Blue):X2}";
        private static int ToByte(double value) => Math.Max(0, Math.Min(255, (int)Math.Round(value * 255d)));

        private static void ValidateOptions(GoogleDocsImportOptions options) {
            if (options.MaxResponseBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxResponseBytes));
            if (options.MaxTabs <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTabs));
            if (options.MaxStructuralElements <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxStructuralElements));
            if (options.MaxTableCells <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTableCells));
            if (options.MaxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTextCharacters));
        }

        private static void ValidateNativeResponse(
            GoogleDocsApiDocumentResponse response,
            GoogleDocsImportOptions options) {
            List<GoogleDocsApiTabResponse> tabs = GoogleDocsApiPayloadBuilder.FlattenTabs(response.Tabs).ToList();
            if (tabs.Count > options.MaxTabs) {
                throw new InvalidDataException($"Google Docs native import exceeded the configured {options.MaxTabs} tab limit.");
            }

            int structuralElements = 0;
            int tableCells = 0;
            long textCharacters = 0;
            if (tabs.Count == 0) {
                CountNativeContent(response.Body?.Content, options,
                    ref structuralElements, ref tableCells, ref textCharacters);
                return;
            }

            foreach (GoogleDocsApiTabResponse tab in tabs) {
                CountNativeContent(tab.DocumentTab?.Body?.Content, options,
                    ref structuralElements, ref tableCells, ref textCharacters);
            }
        }

        private static void CountNativeContent(
            IReadOnlyList<GoogleDocsApiStructuralElementResponse>? content,
            GoogleDocsImportOptions options,
            ref int structuralElements,
            ref int tableCells,
            ref long textCharacters) {
            if (content == null) return;
            foreach (GoogleDocsApiStructuralElementResponse element in content) {
                if (structuralElements >= options.MaxStructuralElements) {
                    throw new InvalidDataException($"Google Docs native import exceeded the configured {options.MaxStructuralElements} structural-element limit.");
                }
                structuralElements++;
                if (element.Paragraph != null) {
                    foreach (GoogleDocsApiParagraphInlineElementResponse inline in element.Paragraph.Elements) {
                        int length = inline.TextRun?.Content?.Length ?? 0;
                        if (length > options.MaxTextCharacters - textCharacters) {
                            throw new InvalidDataException($"Google Docs native import exceeded the configured {options.MaxTextCharacters} text-character limit.");
                        }
                        textCharacters += length;
                    }
                }
                if (element.Table == null) continue;
                foreach (GoogleDocsApiTableRowResponse row in element.Table.Rows) {
                    foreach (GoogleDocsApiTableCellResponse cell in row.Cells) {
                        if (tableCells >= options.MaxTableCells) {
                            throw new InvalidDataException($"Google Docs native import exceeded the configured {options.MaxTableCells} table-cell limit.");
                        }
                        tableCells++;
                        CountNativeContent(cell.Content, options,
                            ref structuralElements, ref tableCells, ref textCharacters);
                    }
                }
            }
        }

        private static void EnsureDocument(GoogleDriveFile file, string id) {
            if (!string.Equals(file.MimeType, GoogleDriveMimeTypes.Document, StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Drive file '{id}' is not a Google document (mimeType: '{file.MimeType}').");
            }
        }

        private static void EnsureDownloadable(GoogleDriveFile file, string id) {
            if (file.Capabilities != null && !file.Capabilities.CanDownload) throw new InvalidOperationException($"Drive file '{id}' cannot be exported by the current principal.");
        }

        private static GoogleDocumentReference BuildReference(GoogleDriveFile file, string id, TranslationReport report, GoogleDocsApiDocumentResponse? native = null) {
            return new GoogleDocumentReference {
                DocumentId = native?.DocumentId ?? file.Id ?? id,
                FileId = file.Id ?? id,
                Name = native?.Title ?? file.Name,
                MimeType = file.MimeType,
                WebViewLink = file.WebViewLink,
                RevisionId = native?.RevisionId,
                DriveVersion = file.Version,
                ModifiedTime = file.ModifiedTime,
                Report = report,
            };
        }
    }
}
