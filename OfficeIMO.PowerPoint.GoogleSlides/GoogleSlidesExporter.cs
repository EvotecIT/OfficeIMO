using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public sealed class GoogleSlidesExporter : IGoogleSlidesExporter {
        private const int RequestsPerBatch = 250;

        public GoogleSlidesTranslationPlan BuildPlan(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) => BuildBatch(presentation, options).Plan;

        public GoogleSlidesBatch BuildBatch(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return GoogleSlidesBatchCompiler.Build(presentation, options ?? new GoogleSlidesSaveOptions());
        }

        public async Task<GooglePresentationReference> ExportAsync(PowerPointPresentation presentation, GoogleWorkspaceSession session, GoogleSlidesSaveOptions? options = null, CancellationToken cancellationToken = default) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (session == null) throw new ArgumentNullException(nameof(session));
            GoogleSlidesSaveOptions effective = options ?? new GoogleSlidesSaveOptions();
            GoogleSlidesBatch batch = BuildBatch(presentation, effective);
            GoogleWorkspacePreflight.Validate(batch.Plan.Report, effective.FidelityPolicy);
            GoogleDriveFileLocation location = session.ResolveLocationDefaults(effective.Location);
            GoogleWorkspaceAccessToken token = await session.AcquireAccessTokenAsync(GoogleWorkspaceScopeCatalog.SlidesAuthoring, cancellationToken).ConfigureAwait(false);
            using var transport = new GoogleWorkspaceHttpTransport(session.Options);
            using var drive = new GoogleDriveClient(session);
            var leases = new List<GoogleDriveTemporaryContentLease>();
            try {
                string presentationId;
                bool copiedTemplate = !string.IsNullOrWhiteSpace(effective.TemplatePresentationId);
                if (copiedTemplate) {
                    GoogleDriveFile copy = await drive.CopyFileAsync(effective.TemplatePresentationId!, batch.Title, location.FolderId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                    presentationId = copy.Id ?? throw new InvalidOperationException("Drive template copy did not return an id.");
                    batch.Plan.Report.Add(TranslationSeverity.Info, "Templates", "Copied the requested Google Slides template before applying the OfficeIMO batch.", code: "SLIDES.TEMPLATE.COPIED", action: TranslationAction.Preserve);
                } else if (!string.IsNullOrWhiteSpace(location.ExistingFileId)) {
                    presentationId = location.ExistingFileId!;
                } else {
                    GoogleSlidesApiPresentationResponse created = await transport.SendJsonAsync<GoogleSlidesApiPresentationResponse>(token.AccessToken, HttpMethod.Post,
                        "https://slides.googleapis.com/v1/presentations", new { title = batch.Title }, GoogleWorkspaceRequestSafety.NonIdempotent, "Google Slides API", batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                    presentationId = created.PresentationId ?? throw new InvalidOperationException("Google Slides create response did not return a presentationId.");
                }

                GoogleSlidesApiPresentationResponse current = await GetPresentationAsync(transport, token.AccessToken, presentationId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                string? revision = ResolveRevision(effective, current, copiedTemplate || string.IsNullOrWhiteSpace(location.ExistingFileId), batch.Plan.Report);
                IReadOnlyDictionary<string, string> imageUrls = await CreateImageLeasesAsync(drive, batch, leases, cancellationToken).ConfigureAwait(false);
                List<object> requests = BuildRequests(batch, current, imageUrls);
                revision = await SendRequestsAsync(transport, token.AccessToken, presentationId, requests, revision, batch.Plan.Report, cancellationToken).ConfigureAwait(false);

                if (batch.Slides.Any(slide => !string.IsNullOrWhiteSpace(slide.SpeakerNotes))) {
                    GoogleSlidesApiPresentationResponse withNotes = await GetPresentationAsync(transport, token.AccessToken, presentationId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                    revision = withNotes.RevisionId ?? revision;
                    List<object> noteRequests = BuildSpeakerNotesRequests(batch, withNotes);
                    revision = await SendRequestsAsync(transport, token.AccessToken, presentationId, noteRequests, revision, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                }

                if (!copiedTemplate && !string.IsNullOrWhiteSpace(location.FolderId)) {
                    await drive.MoveFileAsync(presentationId, location.FolderId!, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                }
                GoogleDriveFile metadata = await drive.GetFileAsync(presentationId, report: batch.Plan.Report, cancellationToken: cancellationToken).ConfigureAwait(false);
                return new GooglePresentationReference {
                    PresentationId = presentationId, FileId = presentationId, Name = metadata.Name ?? batch.Title,
                    MimeType = metadata.MimeType ?? GoogleDriveMimeTypes.Presentation,
                    WebViewLink = metadata.WebViewLink ?? $"https://docs.google.com/presentation/d/{presentationId}/edit",
                    Location = location, RevisionId = revision, DriveVersion = metadata.Version, ModifiedTime = metadata.ModifiedTime, Report = batch.Plan.Report,
                };
            } catch (GoogleWorkspaceApiException ex) when (
                effective.Replace.ConflictMode == GoogleSlidesRevisionConflictMode.RequireRevision
                && !string.IsNullOrWhiteSpace(effective.Replace.ExpectedRevisionId)
                && ex.ResponseStatusCode == System.Net.HttpStatusCode.BadRequest) {
                throw new GoogleWorkspaceConflictException("Google presentation changed before the batch could be applied.", effective.Location.ExistingFileId ?? effective.TemplatePresentationId ?? "presentation", effective.Replace.ExpectedRevisionId, null, batch.Plan.Report);
            } finally {
                foreach (GoogleDriveTemporaryContentLease lease in leases.AsEnumerable().Reverse()) await lease.CleanupAsync(CancellationToken.None).ConfigureAwait(false);
            }
        }

        private static string? ResolveRevision(GoogleSlidesSaveOptions options, GoogleSlidesApiPresentationResponse current, bool ownsNewCopy, TranslationReport report) {
            if (ownsNewCopy) return current.RevisionId;
            if (options.Replace.ConflictMode == GoogleSlidesRevisionConflictMode.OverwriteLatest) return null;
            if (string.IsNullOrWhiteSpace(options.Replace.ExpectedRevisionId)) {
                report.Add(TranslationSeverity.Error, "ReplaceConflict", "Replacing a Google presentation requires the revision observed by a prior read/import.", code: "SLIDES.REPLACE.EXPECTED_REVISION_REQUIRED", action: TranslationAction.Fail);
                throw new GoogleWorkspacePreflightException("Google Slides replacement requires Replace.ExpectedRevisionId unless OverwriteLatest is selected.", report, report.Notices.Where(notice => notice.Code == "SLIDES.REPLACE.EXPECTED_REVISION_REQUIRED").ToArray());
            }
            if (!string.Equals(options.Replace.ExpectedRevisionId, current.RevisionId, StringComparison.Ordinal)) {
                throw new GoogleWorkspaceConflictException("Google presentation changed after it was read.", current.PresentationId ?? "presentation", options.Replace.ExpectedRevisionId, current.RevisionId, report);
            }
            return current.RevisionId;
        }

        private static async Task<GoogleSlidesApiPresentationResponse> GetPresentationAsync(GoogleWorkspaceHttpTransport transport, string token, string id, TranslationReport report, CancellationToken cancellationToken) =>
            await transport.SendJsonAsync<GoogleSlidesApiPresentationResponse>(token, HttpMethod.Get, $"https://slides.googleapis.com/v1/presentations/{Uri.EscapeDataString(id)}", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Slides API", report, cancellationToken).ConfigureAwait(false);

        private static async Task<IReadOnlyDictionary<string, string>> CreateImageLeasesAsync(GoogleDriveClient drive, GoogleSlidesBatch batch, IList<GoogleDriveTemporaryContentLease> leases, CancellationToken cancellationToken) {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (GoogleSlidesImage image in batch.Slides.SelectMany(slide => slide.Elements).OfType<GoogleSlidesImage>()) {
                GoogleDriveTemporaryContentLease lease = await GoogleDriveTemporaryContentLease.CreatePublicReadLeaseAsync(drive, image.Bytes,
                    new GoogleDriveUploadOptions { Name = image.FileName, ContentType = image.ContentType }, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                leases.Add(lease); result[image.ObjectId] = lease.PublicUri;
            }
            return result;
        }

        private static List<object> BuildRequests(GoogleSlidesBatch batch, GoogleSlidesApiPresentationResponse current, IReadOnlyDictionary<string, string> imageUrls) {
            ResolvePagePlacement(batch, current, out double scale, out double offsetX, out double offsetY);
            var requests = current.Slides.Where(slide => !string.IsNullOrWhiteSpace(slide.ObjectId)).Select(slide => (object)new { deleteObject = new { objectId = slide.ObjectId } }).ToList();
            foreach (GoogleSlidesSlide slide in batch.Slides) {
                requests.Add(new { createSlide = new { objectId = slide.ObjectId, insertionIndex = slide.Index, slideLayoutReference = new { predefinedLayout = "BLANK" } } });
                if (!string.IsNullOrWhiteSpace(slide.BackgroundColorHex)) requests.Add(new { updateSlideProperties = new { objectId = slide.ObjectId, slideProperties = new { background = new { solidFill = new { color = new { rgbColor = Rgb(slide.BackgroundColorHex!) } } } }, fields = "background" } });
                foreach (GoogleSlidesElement element in slide.Elements) AddElementRequests(requests, slide.ObjectId, element, imageUrls, scale, offsetX, offsetY);
            }
            return requests;
        }

        private static void AddElementRequests(
            ICollection<object> requests,
            string slideId,
            GoogleSlidesElement element,
            IReadOnlyDictionary<string, string> imageUrls,
            double scale,
            double offsetX,
            double offsetY) {
            object properties = ElementProperties(slideId, element, scale, offsetX, offsetY);
            switch (element) {
                case GoogleSlidesTextBox text:
                    requests.Add(new { createShape = new { objectId = text.ObjectId, shapeType = "TEXT_BOX", elementProperties = properties } });
                    if (!string.IsNullOrEmpty(text.Text)) requests.Add(new { insertText = new { objectId = text.ObjectId, text = text.Text } });
                    var style = new Dictionary<string, object?>();
                    if (text.Bold) style["bold"] = true; if (text.Italic) style["italic"] = true; if (text.Underline) style["underline"] = true;
                    if (text.FontSize.HasValue) style["fontSize"] = new { magnitude = Math.Max(1, text.FontSize.Value * scale), unit = "PT" };
                    if (!string.IsNullOrWhiteSpace(text.FontFamily)) style["fontFamily"] = text.FontFamily;
                    if (!string.IsNullOrWhiteSpace(text.ForegroundColorHex)) style["foregroundColor"] = new { opaqueColor = new { rgbColor = Rgb(text.ForegroundColorHex!) } };
                    if (!string.IsNullOrWhiteSpace(text.Hyperlink)) style["link"] = new { url = text.Hyperlink };
                    if (style.Count > 0 && text.Text.Length > 0) requests.Add(new { updateTextStyle = new { objectId = text.ObjectId, textRange = new { type = "ALL" }, style, fields = string.Join(",", style.Keys) } });
                    break;
                case GoogleSlidesTable table:
                    int rows = table.Cells.Count; int columns = table.Cells.Select(row => row.Count).DefaultIfEmpty(0).Max();
                    if (rows == 0 || columns == 0) break;
                    requests.Add(new { createTable = new { objectId = table.ObjectId, rows, columns, elementProperties = properties } });
                    for (int row = 0; row < rows; row++) for (int column = 0; column < table.Cells[row].Count; column++) if (!string.IsNullOrEmpty(table.Cells[row][column]))
                        requests.Add(new { insertText = new { objectId = table.ObjectId, cellLocation = new { rowIndex = row, columnIndex = column }, text = table.Cells[row][column] } });
                    break;
                case GoogleSlidesImage image:
                    requests.Add(new { createImage = new { objectId = image.ObjectId, url = imageUrls[image.ObjectId], elementProperties = properties } });
                    break;
                case GoogleSlidesShape shape:
                    requests.Add(new { createShape = new { objectId = shape.ObjectId, shapeType = shape.ShapeType, elementProperties = properties } });
                    break;
            }
        }

        private static object ElementProperties(string slideId, GoogleSlidesElement element, double scale, double offsetX, double offsetY) => new {
            pageObjectId = slideId,
            size = new { width = new { magnitude = Math.Max(1, element.WidthPoints * scale), unit = "PT" }, height = new { magnitude = Math.Max(1, element.HeightPoints * scale), unit = "PT" } },
            transform = new { scaleX = 1, scaleY = 1, translateX = offsetX + (element.LeftPoints * scale), translateY = offsetY + (element.TopPoints * scale), unit = "PT" },
        };

        private static void ResolvePagePlacement(
            GoogleSlidesBatch batch,
            GoogleSlidesApiPresentationResponse current,
            out double scale,
            out double offsetX,
            out double offsetY) {
            scale = 1;
            offsetX = 0;
            offsetY = 0;
            double targetWidth = ToPoints(current.PageSize?.Width);
            double targetHeight = ToPoints(current.PageSize?.Height);
            if (batch.WidthPoints <= 0 || batch.HeightPoints <= 0 || targetWidth <= 0 || targetHeight <= 0) return;

            scale = Math.Min(targetWidth / batch.WidthPoints, targetHeight / batch.HeightPoints);
            offsetX = (targetWidth - (batch.WidthPoints * scale)) / 2d;
            offsetY = (targetHeight - (batch.HeightPoints * scale)) / 2d;
            if (Math.Abs(targetWidth - batch.WidthPoints) < 0.01 && Math.Abs(targetHeight - batch.HeightPoints) < 0.01) return;

            batch.Plan.Report.AddUnique(
                TranslationSeverity.Info,
                "PageSize",
                $"Google Slides does not expose presentation page-size updates; elements were proportionally scaled and centered from {batch.WidthPoints:0.##}x{batch.HeightPoints:0.##} pt to {targetWidth:0.##}x{targetHeight:0.##} pt.",
                code: "SLIDES.PAGE_SIZE.SCALED",
                action: TranslationAction.Preserve);
        }

        private static double ToPoints(GoogleSlidesApiDimension? dimension) {
            if (dimension == null) return 0;
            return string.Equals(dimension.Unit, "EMU", StringComparison.OrdinalIgnoreCase)
                ? dimension.Magnitude / 12700d
                : dimension.Magnitude;
        }

        private static Dictionary<string, double> Rgb(string hex) {
            string value = hex.TrimStart('#'); if (value.Length >= 6) value = value.Substring(value.Length - 6);
            if (value.Length != 6) return new Dictionary<string, double>();
            return new Dictionary<string, double> { ["red"] = Convert.ToInt32(value.Substring(0, 2), 16) / 255d, ["green"] = Convert.ToInt32(value.Substring(2, 2), 16) / 255d, ["blue"] = Convert.ToInt32(value.Substring(4, 2), 16) / 255d };
        }

        private static List<object> BuildSpeakerNotesRequests(GoogleSlidesBatch batch, GoogleSlidesApiPresentationResponse current) {
            var requests = new List<object>();
            foreach (GoogleSlidesSlide slide in batch.Slides.Where(slide => !string.IsNullOrWhiteSpace(slide.SpeakerNotes))) {
                GoogleSlidesApiPage? page = current.Slides.FirstOrDefault(candidate => string.Equals(candidate.ObjectId, slide.ObjectId, StringComparison.Ordinal));
                string? notesId = page?.SlideProperties?.NotesPage?.NotesProperties?.SpeakerNotesObjectId;
                if (string.IsNullOrWhiteSpace(notesId)) continue;
                requests.Add(new { deleteText = new { objectId = notesId, textRange = new { type = "ALL" } } });
                requests.Add(new { insertText = new { objectId = notesId, text = slide.SpeakerNotes } });
            }
            return requests;
        }

        private static async Task<string?> SendRequestsAsync(GoogleWorkspaceHttpTransport transport, string token, string id, IReadOnlyList<object> requests, string? revision, TranslationReport report, CancellationToken cancellationToken) {
            for (int offset = 0; offset < requests.Count; offset += RequestsPerBatch) {
                object[] chunk = requests.Skip(offset).Take(RequestsPerBatch).ToArray();
                var payload = new Dictionary<string, object?> { ["requests"] = chunk };
                if (!string.IsNullOrWhiteSpace(revision)) payload["writeControl"] = new { requiredRevisionId = revision };
                GoogleSlidesApiBatchResponse response = await transport.SendJsonAsync<GoogleSlidesApiBatchResponse>(token, HttpMethod.Post,
                    $"https://slides.googleapis.com/v1/presentations/{Uri.EscapeDataString(id)}:batchUpdate", payload, GoogleWorkspaceRequestSafety.NonIdempotent, "Google Slides API", report, cancellationToken).ConfigureAwait(false);
                revision = response.WriteControl?.RequiredRevisionId ?? revision;
            }
            return revision;
        }
    }
}
