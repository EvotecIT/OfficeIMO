using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.PowerPoint;
using System.Text.Json.Nodes;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public sealed class GoogleSlidesExporter : IGoogleSlidesExporter {
        private const int RequestsPerBatch = 250;

        public GoogleSlidesTranslationPlan BuildPlan(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return GoogleSlidesBatchCompiler.BuildPlan(presentation, options ?? new GoogleSlidesSaveOptions());
        }

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
            using var drive = new GoogleDriveClient(session, GoogleDriveClientOptions.ForFileAuthoring());
            var leases = new List<GoogleDriveTemporaryContentLease>();
            string? presentationId = null;
            try {
                if (!string.IsNullOrWhiteSpace(location.FolderId)) {
                    await drive.ResolveFolderAsync(location.FolderId!, location.DriveId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                }

                bool copiedTemplate = !string.IsNullOrWhiteSpace(effective.TemplatePresentationId);
                if (copiedTemplate) {
                    GoogleDriveFile copy = await drive.CopyFileAsync(effective.TemplatePresentationId!, batch.Title, location.FolderId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                    presentationId = copy.Id ?? throw new InvalidOperationException("Drive template copy did not return an id.");
                    batch.Plan.Report.Add(TranslationSeverity.Info, "Templates", "Copied the requested Google Slides template before applying the OfficeIMO batch.", code: "SLIDES.TEMPLATE.COPIED", action: TranslationAction.Preserve);
                } else if (!string.IsNullOrWhiteSpace(location.ExistingFileId)) {
                    presentationId = location.ExistingFileId!;
                    await ValidateExistingPresentationDriveAccessAsync(
                        drive,
                        presentationId,
                        batch.Plan.Report,
                        cancellationToken).ConfigureAwait(false);
                } else {
                    GoogleSlidesApiPresentationResponse created = await transport.SendJsonAsync(
                        token.AccessToken,
                        HttpMethod.Post,
                        "https://slides.googleapis.com/v1/presentations",
                        Obj(("title", batch.Title)),
                        GoogleWorkspaceRequestSafety.NonIdempotent,
                        "Google Slides API",
                        batch.Plan.Report,
                        GoogleSlidesJsonSerializerContext.Default.GoogleSlidesApiPresentationResponse,
                        cancellationToken).ConfigureAwait(false);
                    presentationId = created.PresentationId ?? throw new InvalidOperationException("Google Slides create response did not return a presentationId.");
                }

                GoogleSlidesApiPresentationResponse current = await GetPresentationAsync(transport, token.AccessToken, presentationId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                bool ownsNewCopy = copiedTemplate || string.IsNullOrWhiteSpace(location.ExistingFileId);
                bool overwritingExisting = !ownsNewCopy && effective.Replace.ConflictMode == GoogleSlidesRevisionConflictMode.OverwriteLatest;
                bool classifyRevisionConflicts = !ownsNewCopy && effective.Replace.ConflictMode == GoogleSlidesRevisionConflictMode.RequireRevision;
                string? revision = ResolveRevision(effective, current, ownsNewCopy, batch.Plan.Report);
                IReadOnlyDictionary<string, string> imageUrls = await CreateImageLeasesAsync(drive, batch, leases, cancellationToken).ConfigureAwait(false);
                List<JsonObject> requests = BuildRequests(batch, current, imageUrls);
                revision = await SendRequestsAsync(transport, token.AccessToken, presentationId, requests, revision, classifyRevisionConflicts, batch.Plan.Report, cancellationToken).ConfigureAwait(false);

                if (batch.Slides.Any(slide => !string.IsNullOrWhiteSpace(slide.SpeakerNotes))) {
                    GoogleSlidesApiPresentationResponse withNotes = await GetPresentationAsync(transport, token.AccessToken, presentationId, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                    revision = withNotes.RevisionId ?? revision;
                    List<JsonObject> noteRequests = BuildSpeakerNotesRequests(batch, withNotes);
                    revision = await SendRequestsAsync(transport, token.AccessToken, presentationId, noteRequests, revision, classifyRevisionConflicts, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                }

                if (overwritingExisting || string.IsNullOrWhiteSpace(revision)) {
                    GoogleSlidesApiPresentationResponse refreshed = await GetPresentationAsync(
                        transport,
                        token.AccessToken,
                        presentationId,
                        batch.Plan.Report,
                        cancellationToken).ConfigureAwait(false);
                    revision = refreshed.RevisionId;
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

        private static async Task ValidateExistingPresentationDriveAccessAsync(
            GoogleDriveClient drive,
            string presentationId,
            TranslationReport report,
            CancellationToken cancellationToken) {
            GoogleDriveFile metadata;
            try {
                metadata = await drive.GetFileAsync(
                    presentationId,
                    report: report,
                    cancellationToken: cancellationToken).ConfigureAwait(false);
            } catch (GoogleWorkspaceApiException exception) when (
                exception.ResponseStatusCode == System.Net.HttpStatusCode.Forbidden
                || exception.ResponseStatusCode == System.Net.HttpStatusCode.NotFound) {
                report.Add(
                    TranslationSeverity.Error,
                    "ExistingPresentation",
                    "The existing Google presentation is not available through the configured Drive authoring grant. Open or create it through this app, or provide credentials with a Drive scope that covers the target before replacing it.",
                    code: "SLIDES.REPLACE.DRIVE_ACCESS_REQUIRED",
                    action: TranslationAction.Fail,
                    targetId: presentationId);
                throw new GoogleWorkspacePreflightException(
                    $"Google Slides replacement was blocked before mutation because Drive metadata for '{presentationId}' is not accessible.",
                    report,
                    report.Notices.Where(notice => notice.Code == "SLIDES.REPLACE.DRIVE_ACCESS_REQUIRED").ToArray());
            }

            if (metadata.Capabilities != null && !metadata.Capabilities.CanEdit) {
                report.Add(
                    TranslationSeverity.Error,
                    "ExistingPresentation",
                    "The configured Google identity cannot edit the existing Drive file.",
                    code: "SLIDES.REPLACE.DRIVE_EDIT_REQUIRED",
                    action: TranslationAction.Fail,
                    targetId: presentationId);
                throw new GoogleWorkspacePreflightException(
                    $"Google Slides replacement was blocked before mutation because '{presentationId}' is not editable.",
                    report,
                    report.Notices.Where(notice => notice.Code == "SLIDES.REPLACE.DRIVE_EDIT_REQUIRED").ToArray());
            }
        }

        private static async Task<GoogleSlidesApiPresentationResponse> GetPresentationAsync(GoogleWorkspaceHttpTransport transport, string token, string id, TranslationReport report, CancellationToken cancellationToken) =>
            await transport.SendJsonAsync<GoogleSlidesApiPresentationResponse>(token, HttpMethod.Get, $"https://slides.googleapis.com/v1/presentations/{Uri.EscapeDataString(id)}", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Slides API", report,
                GoogleSlidesJsonSerializerContext.Default.GoogleSlidesApiPresentationResponse, cancellationToken).ConfigureAwait(false);

        private static async Task<string?> TryGetPresentationRevisionAsync(
            GoogleWorkspaceHttpTransport transport,
            string token,
            string? presentationId,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(presentationId)) return null;
            try {
                GoogleSlidesApiPresentationResponse presentation = await GetPresentationAsync(
                    transport,
                    token,
                    presentationId!,
                    report,
                    cancellationToken).ConfigureAwait(false);
                return presentation.RevisionId;
            } catch (GoogleWorkspaceApiException) {
                return null;
            }
        }

        private static async Task<IReadOnlyDictionary<string, string>> CreateImageLeasesAsync(GoogleDriveClient drive, GoogleSlidesBatch batch, IList<GoogleDriveTemporaryContentLease> leases, CancellationToken cancellationToken) {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            IEnumerable<GoogleSlidesImage> images = batch.Slides
                .SelectMany(slide => slide.Elements)
                .OfType<GoogleSlidesImage>()
                .Concat(batch.Slides.Select(slide => slide.BackgroundImage).OfType<GoogleSlidesImage>());
            foreach (GoogleSlidesImage image in images) {
                GoogleDriveTemporaryContentLease lease = await GoogleDriveTemporaryContentLease.CreatePublicReadLeaseAsync(drive, image.Bytes,
                    new GoogleDriveUploadOptions { Name = image.FileName, ContentType = image.ContentType }, batch.Plan.Report, cancellationToken).ConfigureAwait(false);
                leases.Add(lease); result[image.ObjectId] = lease.PublicUri;
            }
            return result;
        }

        private static List<JsonObject> BuildRequests(GoogleSlidesBatch batch, GoogleSlidesApiPresentationResponse current, IReadOnlyDictionary<string, string> imageUrls) {
            ResolvePagePlacement(batch, current, out double scale, out double offsetX, out double offsetY);
            var existingSlideIds = current.Slides
                .Where(slide => !string.IsNullOrWhiteSpace(slide.ObjectId))
                .Select(slide => slide.ObjectId!)
                .ToList();
            var requests = new List<JsonObject>();
            string? keeperSlideId = null;
            if (existingSlideIds.Count > 0) {
                var occupiedIds = new HashSet<string>(existingSlideIds.Concat(batch.Slides.Select(slide => slide.ObjectId)), StringComparer.Ordinal);
                keeperSlideId = "officeimo_replacement_keeper";
                for (int suffix = 2; occupiedIds.Contains(keeperSlideId); suffix++) {
                    keeperSlideId = "officeimo_replacement_keeper_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                requests.Add(Obj(("createSlide", Obj(
                    ("objectId", keeperSlideId),
                    ("insertionIndex", existingSlideIds.Count),
                    ("slideLayoutReference", Obj(("predefinedLayout", "BLANK")))))));
            }

            foreach (string existingSlideId in existingSlideIds) {
                requests.Add(Obj(("deleteObject", Obj(("objectId", existingSlideId)))));
            }

            foreach (GoogleSlidesSlide slide in batch.Slides) {
                requests.Add(Obj(("createSlide", Obj(
                    ("objectId", slide.ObjectId),
                    ("insertionIndex", slide.Index),
                    ("slideLayoutReference", Obj(("predefinedLayout", "BLANK")))))));
                if (slide.IsSkipped) {
                    requests.Add(Obj(("updateSlideProperties", Obj(
                        ("objectId", slide.ObjectId),
                        ("slideProperties", Obj(("isSkipped", true))),
                        ("fields", "isSkipped")))));
                }
                if (!string.IsNullOrWhiteSpace(slide.BackgroundColorHex)) {
                    JsonObject pageBackgroundFill = Obj(("solidFill", Obj(
                        ("color", Obj(("rgbColor", Rgb(slide.BackgroundColorHex!)))))));
                    requests.Add(Obj(("updatePageProperties", Obj(
                        ("objectId", slide.ObjectId),
                        ("pageProperties", Obj(("pageBackgroundFill", pageBackgroundFill))),
                        ("fields", "pageBackgroundFill.solidFill.color")))));
                } else if (slide.BackgroundImage != null) {
                    requests.Add(Obj(("updatePageProperties", Obj(
                        ("objectId", slide.ObjectId),
                        ("pageProperties", Obj(("pageBackgroundFill", Obj(("stretchedPictureFill", Obj(("contentUrl", imageUrls[slide.BackgroundImage.ObjectId]))))))),
                        ("fields", "pageBackgroundFill.stretchedPictureFill.contentUrl")))));
                }
                foreach (GoogleSlidesElement element in slide.Elements) AddElementRequests(requests, slide.ObjectId, element, imageUrls, scale, offsetX, offsetY);
            }
            if (keeperSlideId != null && batch.Slides.Count > 0) {
                requests.Add(Obj(("deleteObject", Obj(("objectId", keeperSlideId)))));
            }
            return requests;
        }

        private static void AddElementRequests(
            ICollection<JsonObject> requests,
            string slideId,
            GoogleSlidesElement element,
            IReadOnlyDictionary<string, string> imageUrls,
            double scale,
            double offsetX,
            double offsetY) {
            JsonObject properties = ElementProperties(slideId, element, scale, offsetX, offsetY);
            switch (element) {
                case GoogleSlidesTextBox text:
                    requests.Add(Obj(("createShape", Obj(("objectId", text.ObjectId), ("shapeType", text.ShapeType), ("elementProperties", properties)))));
                    AddShapeStyleRequest(requests, text.ObjectId, text.Style, scale);
                    if (!string.IsNullOrEmpty(text.Text)) requests.Add(Obj(("insertText", Obj(("objectId", text.ObjectId), ("text", text.Text)))));
                    if (text.TextRuns.Count > 0) {
                        foreach (GoogleSlidesTextStyleRun run in text.TextRuns) {
                            JsonObject style = BuildTextStyle(run.Bold, run.Italic, run.Underline, run.FontSize, run.FontFamily, run.ForegroundColorHex, run.Hyperlink, scale);
                            if (style.Count > 0 && run.EndIndex > run.StartIndex) {
                                requests.Add(Obj(("updateTextStyle", Obj(
                                    ("objectId", text.ObjectId),
                                    ("textRange", Obj(("type", "FIXED_RANGE"), ("startIndex", run.StartIndex), ("endIndex", run.EndIndex))),
                                    ("style", style),
                                    ("fields", string.Join(",", style.Select(pair => pair.Key)))))));
                            }
                        }
                    } else {
                        JsonObject style = BuildTextStyle(text.Bold, text.Italic, text.Underline, text.FontSize, text.FontFamily, text.ForegroundColorHex, text.Hyperlink, scale);
                        if (style.Count > 0 && text.Text.Length > 0) {
                            requests.Add(Obj(("updateTextStyle", Obj(
                                ("objectId", text.ObjectId),
                                ("textRange", Obj(("type", "ALL"))),
                                ("style", style),
                                ("fields", string.Join(",", style.Select(pair => pair.Key)))))));
                        }
                    }
                    break;
                case GoogleSlidesTable table:
                    int rows = table.Cells.Count; int columns = table.Cells.Select(row => row.Count).DefaultIfEmpty(0).Max();
                    if (rows == 0 || columns == 0) break;
                    requests.Add(Obj(("createTable", Obj(("objectId", table.ObjectId), ("rows", rows), ("columns", columns), ("elementProperties", properties)))));
                    for (int row = 0; row < rows; row++) for (int column = 0; column < table.Cells[row].Count; column++) if (!string.IsNullOrEmpty(table.Cells[row][column]))
                        requests.Add(Obj(("insertText", Obj(
                            ("objectId", table.ObjectId),
                            ("cellLocation", Obj(("rowIndex", row), ("columnIndex", column))),
                            ("text", table.Cells[row][column])))));
                    break;
                case GoogleSlidesImage image:
                    requests.Add(Obj(("createImage", Obj(("objectId", image.ObjectId), ("url", imageUrls[image.ObjectId]), ("elementProperties", properties)))));
                    break;
                case GoogleSlidesShape shape:
                    requests.Add(Obj(("createShape", Obj(("objectId", shape.ObjectId), ("shapeType", shape.ShapeType), ("elementProperties", properties)))));
                    AddShapeStyleRequest(requests, shape.ObjectId, shape.Style, scale);
                    break;
            }
        }

        private static JsonObject BuildTextStyle(
            bool bold,
            bool italic,
            bool underline,
            int? fontSize,
            string? fontFamily,
            string? foregroundColorHex,
            string? hyperlink,
            double scale) {
            var style = new JsonObject();
            if (bold) style["bold"] = true;
            if (italic) style["italic"] = true;
            if (underline) style["underline"] = true;
            if (fontSize.HasValue) style["fontSize"] = Obj(("magnitude", Math.Max(1, fontSize.Value * scale)), ("unit", "PT"));
            if (!string.IsNullOrWhiteSpace(fontFamily)) style["fontFamily"] = fontFamily;
            if (!string.IsNullOrWhiteSpace(foregroundColorHex)) style["foregroundColor"] = Obj(("opaqueColor", Obj(("rgbColor", Rgb(foregroundColorHex!)))));
            if (!string.IsNullOrWhiteSpace(hyperlink)) style["link"] = Obj(("url", hyperlink));
            return style;
        }

        private static void AddShapeStyleRequest(ICollection<JsonObject> requests, string objectId, GoogleSlidesShapeStyle style, double scale) {
            var shapeProperties = new JsonObject();
            var fields = new List<string>();

            if (!string.IsNullOrWhiteSpace(style.FillColorHex)) {
                var solidFill = Obj(("color", Obj(("rgbColor", Rgb(style.FillColorHex!)))));
                fields.Add("shapeBackgroundFill.solidFill.color");
                if (style.FillTransparencyPercent.HasValue) {
                    int transparency = Math.Min(100, Math.Max(0, style.FillTransparencyPercent.Value));
                    solidFill["alpha"] = (100d - transparency) / 100d;
                    fields.Add("shapeBackgroundFill.solidFill.alpha");
                }
                shapeProperties["shapeBackgroundFill"] = Obj(("solidFill", solidFill));
            }

            var outline = new JsonObject();
            if (!string.IsNullOrWhiteSpace(style.OutlineColorHex)) {
                outline["outlineFill"] = Obj(("solidFill", Obj(("color", Obj(("rgbColor", Rgb(style.OutlineColorHex!)))))));
                fields.Add("outline.outlineFill.solidFill.color");
            }
            if (style.OutlineWidthPoints.HasValue) {
                outline["weight"] = Obj(("magnitude", Math.Max(0, style.OutlineWidthPoints.Value * scale)), ("unit", "PT"));
                fields.Add("outline.weight");
            }
            if (outline.Count > 0) shapeProperties["outline"] = outline;

            if (fields.Count > 0) {
                requests.Add(Obj(("updateShapeProperties", Obj(
                    ("objectId", objectId),
                    ("shapeProperties", shapeProperties),
                    ("fields", string.Join(",", fields))))));
            }
        }

        private static JsonObject ElementProperties(string slideId, GoogleSlidesElement element, double scale, double offsetX, double offsetY) {
            double width = Math.Max(1, element.WidthPoints * scale);
            double height = Math.Max(1, element.HeightPoints * scale);
            double left = offsetX + (element.LeftPoints * scale);
            double top = offsetY + (element.TopPoints * scale);
            double radians = element.RotationDegrees * (Math.PI / 180d);
            double cosine = NormalizeTransformComponent(Math.Cos(radians));
            double sine = NormalizeTransformComponent(Math.Sin(radians));
            double horizontalReflection = element.HorizontalFlip ? -1d : 1d;
            double verticalReflection = element.VerticalFlip ? -1d : 1d;
            double scaleX = NormalizeTransformComponent(cosine * horizontalReflection);
            double shearX = NormalizeTransformComponent(-sine * verticalReflection);
            double shearY = NormalizeTransformComponent(sine * horizontalReflection);
            double scaleY = NormalizeTransformComponent(cosine * verticalReflection);
            double translateX = left + (width / 2d) - (scaleX * width / 2d) - (shearX * height / 2d);
            double translateY = top + (height / 2d) - (shearY * width / 2d) - (scaleY * height / 2d);

            return Obj(
                ("pageObjectId", slideId),
                ("size", Obj(
                    ("width", Obj(("magnitude", width), ("unit", "PT"))),
                    ("height", Obj(("magnitude", height), ("unit", "PT"))))),
                ("transform", Obj(
                    ("scaleX", scaleX),
                    ("scaleY", scaleY),
                    ("shearX", shearX),
                    ("shearY", shearY),
                    ("translateX", translateX),
                    ("translateY", translateY),
                    ("unit", "PT"))));
        }

        private static double NormalizeTransformComponent(double value) => Math.Abs(value) < 0.000000000001d ? 0d : value;

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

        private static JsonObject Rgb(string hex) {
            string value = hex.TrimStart('#'); if (value.Length >= 6) value = value.Substring(value.Length - 6);
            if (value.Length != 6) return new JsonObject();
            return Obj(
                ("red", Convert.ToInt32(value.Substring(0, 2), 16) / 255d),
                ("green", Convert.ToInt32(value.Substring(2, 2), 16) / 255d),
                ("blue", Convert.ToInt32(value.Substring(4, 2), 16) / 255d));
        }

        private static List<JsonObject> BuildSpeakerNotesRequests(GoogleSlidesBatch batch, GoogleSlidesApiPresentationResponse current) {
            var requests = new List<JsonObject>();
            foreach (GoogleSlidesSlide slide in batch.Slides.Where(slide => !string.IsNullOrWhiteSpace(slide.SpeakerNotes))) {
                GoogleSlidesApiPage? page = current.Slides.FirstOrDefault(candidate => string.Equals(candidate.ObjectId, slide.ObjectId, StringComparison.Ordinal));
                GoogleSlidesApiPage? notesPage = page?.SlideProperties?.NotesPage;
                string? notesId = notesPage?.NotesProperties?.SpeakerNotesObjectId;
                if (string.IsNullOrWhiteSpace(notesId)) continue;
                if (HasDeletableSpeakerNotes(notesPage!, notesId!)) {
                    requests.Add(Obj(("deleteText", Obj(("objectId", notesId), ("textRange", Obj(("type", "ALL")))))));
                }
                requests.Add(Obj(("insertText", Obj(("objectId", notesId), ("text", slide.SpeakerNotes)))));
            }
            return requests;
        }

        private static bool HasDeletableSpeakerNotes(GoogleSlidesApiPage notesPage, string notesId) {
            GoogleSlidesApiPageElement? notesShape = notesPage.PageElements.FirstOrDefault(element =>
                string.Equals(element.ObjectId, notesId, StringComparison.Ordinal));
            return notesShape?.Shape?.Text?.TextElements.Any(element =>
                element.TextRun?.Content?.Any(character => character != '\r' && character != '\n') == true) == true;
        }

        private static async Task<string?> SendRequestsAsync(
            GoogleWorkspaceHttpTransport transport,
            string token,
            string id,
            IReadOnlyList<JsonObject> requests,
            string? revision,
            bool classifyRevisionConflicts,
            TranslationReport report,
            CancellationToken cancellationToken) {
            for (int offset = 0; offset < requests.Count; offset += RequestsPerBatch) {
                JsonObject[] chunk = requests.Skip(offset).Take(RequestsPerBatch).ToArray();
                var payload = Obj(("requests", new JsonArray(chunk.Select(request => (JsonNode?)request).ToArray())));
                if (!string.IsNullOrWhiteSpace(revision)) payload["writeControl"] = Obj(("requiredRevisionId", revision));
                string? attemptedRevision = revision;
                try {
                    GoogleSlidesApiBatchResponse response = await transport.SendJsonAsync(
                        token,
                        HttpMethod.Post,
                        $"https://slides.googleapis.com/v1/presentations/{Uri.EscapeDataString(id)}:batchUpdate",
                        payload,
                        GoogleWorkspaceRequestSafety.NonIdempotent,
                        "Google Slides API",
                        report,
                        GoogleSlidesJsonSerializerContext.Default.GoogleSlidesApiBatchResponse,
                        cancellationToken).ConfigureAwait(false);
                    revision = response.WriteControl?.RequiredRevisionId ?? revision;
                } catch (GoogleWorkspaceApiException ex) when (
                    classifyRevisionConflicts
                    && !string.IsNullOrWhiteSpace(attemptedRevision)
                    && ex.ResponseStatusCode == System.Net.HttpStatusCode.BadRequest) {
                    string? latestRevision = await TryGetPresentationRevisionAsync(
                        transport,
                        token,
                        id,
                        report,
                        cancellationToken).ConfigureAwait(false);
                    if (!string.IsNullOrWhiteSpace(latestRevision)
                        && !string.Equals(attemptedRevision, latestRevision, StringComparison.Ordinal)) {
                        throw new GoogleWorkspaceConflictException(
                            "Google presentation changed before the batch could be applied.",
                            id,
                            attemptedRevision,
                            latestRevision,
                            report);
                    }

                    throw;
                }
            }
            return revision;
        }

        private static JsonObject Obj(params (string Name, JsonNode? Value)[] properties) {
            var result = new JsonObject();
            foreach ((string name, JsonNode? value) in properties) result[name] = value;
            return result;
        }
    }
}
