using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public sealed class GoogleSlidesImporter : IGoogleSlidesImporter {
        public async Task<GoogleSlidesImportResult> ImportAsync(string presentationId, GoogleWorkspaceSession session, GoogleSlidesImportOptions? options = null, CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(presentationId)) throw new ArgumentException("Presentation ID is required.", nameof(presentationId));
            if (session == null) throw new ArgumentNullException(nameof(session));
            GoogleSlidesImportOptions effective = options ?? new GoogleSlidesImportOptions();
            return effective.Mode == GoogleSlidesImportMode.DriveExport
                ? await ImportDriveAsync(presentationId, session, effective, cancellationToken).ConfigureAwait(false)
                : await ImportNativeAsync(presentationId, session, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<GoogleSlidesImportResult> ImportDriveAsync(string id, GoogleWorkspaceSession session, GoogleSlidesImportOptions options, CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(id, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsurePresentation(source, id);
            EnsureDownloadable(source, id);
            byte[] bytes = await drive.ExportAsync(id, GoogleDriveMimeTypes.MicrosoftPowerPoint, options.Progress, report, cancellationToken).ConfigureAwait(false);
            var stream = new MemoryStream(bytes, writable: true);
            PowerPointPresentation presentation;
            try { presentation = PowerPointPresentation.Load(stream, options.LoadOptions); } catch { stream.Dispose(); throw; }
            report.Add(TranslationSeverity.Info, "DriveExportImport", "The Google presentation was exported to PPTX through Drive and loaded by OfficeIMO.", code: "SLIDES.IMPORT.DRIVE_EXPORT", action: TranslationAction.Preserve);
            return new GoogleSlidesImportResult(presentation, Reference(source, id, report), report);
        }

        private static async Task<GoogleSlidesImportResult> ImportNativeAsync(string id, GoogleWorkspaceSession session, CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(id, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsurePresentation(source, id);
            GoogleWorkspaceAccessToken token = await session.AcquireAccessTokenAsync(new[] { GoogleWorkspaceScopeCatalog.PresentationsReadonly }, cancellationToken).ConfigureAwait(false);
            using var transport = new GoogleWorkspaceHttpTransport(session.Options);
            GoogleSlidesApiPresentationResponse response = await transport.SendJsonAsync<GoogleSlidesApiPresentationResponse>(token.AccessToken, HttpMethod.Get,
                $"https://slides.googleapis.com/v1/presentations/{Uri.EscapeDataString(id)}", null, GoogleWorkspaceRequestSafety.Safe, "Google Slides API", report, cancellationToken).ConfigureAwait(false);
            PowerPointPresentation presentation = await ProjectAsync(response, transport, token.AccessToken, report, cancellationToken).ConfigureAwait(false);
            report.Add(TranslationSeverity.Info, "NativeImport", "Slides, text boxes, core text styles, tables, images, geometry, and speaker-note text were projected into OfficeIMO.", code: "SLIDES.IMPORT.NATIVE", action: TranslationAction.Preserve);
            return new GoogleSlidesImportResult(presentation, Reference(source, id, report, response), report);
        }

        private static async Task<PowerPointPresentation> ProjectAsync(GoogleSlidesApiPresentationResponse source, GoogleWorkspaceHttpTransport transport, string token, TranslationReport report, CancellationToken cancellationToken) {
            var stream = new MemoryStream();
            PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            try {
                if (source.PageSize?.Width != null && source.PageSize.Height != null) presentation.SlideSize.SetSizePoints(ToPoints(source.PageSize.Width), ToPoints(source.PageSize.Height));
                foreach (GoogleSlidesApiPage sourceSlide in source.Slides) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Hidden = sourceSlide.SlideProperties?.IsSkipped == true;
                    GoogleSlidesApiRgbColor? backgroundColor = sourceSlide.PageProperties?.PageBackgroundFill?.SolidFill?.Color?.RgbColor;
                    if (backgroundColor != null) {
                        slide.BackgroundColor = ToHex(backgroundColor);
                    }
                    foreach (GoogleSlidesApiPageElement element in sourceSlide.PageElements) {
                        double left = ToPoints(element.Transform?.TranslateX ?? 0, element.Transform?.Unit);
                        double top = ToPoints(element.Transform?.TranslateY ?? 0, element.Transform?.Unit);
                        double width = ToPoints(element.Size?.Width) * (element.Transform?.ScaleX ?? 1);
                        double height = ToPoints(element.Size?.Height) * (element.Transform?.ScaleY ?? 1);
                        if (element.Shape?.Text != null) {
                            string text = ExtractText(element.Shape.Text);
                            PowerPointTextBox box;
                            if (TryMapShapeType(element.Shape.ShapeType, out A.ShapeTypeValues textShapeType)) {
                                box = slide.AddTextShapePoints(textShapeType, text, left, top, Math.Max(1, width), Math.Max(1, height), element.ObjectId);
                            } else {
                                box = slide.AddTextBoxPoints(text, left, top, Math.Max(1, width), Math.Max(1, height));
                                if (!string.Equals(element.Shape.ShapeType, "TEXT_BOX", StringComparison.OrdinalIgnoreCase)) {
                                    report.Add(
                                        TranslationSeverity.Warning,
                                        "Shapes",
                                        $"Native Google Slides text shape '{element.Shape.ShapeType ?? "unspecified"}' was imported as a plain PowerPoint text box because its geometry is unsupported.",
                                        code: "SLIDES.IMPORT.TEXT_SHAPE_GEOMETRY_UNSUPPORTED",
                                        action: TranslationAction.Flatten);
                                }
                            }
                            GoogleSlidesApiTextRun? style = element.Shape.Text.TextElements.Select(item => item.TextRun).FirstOrDefault(run => run?.Style != null);
                            PowerPointTextRun? run = box.Paragraphs.SelectMany(paragraph => paragraph.Runs).FirstOrDefault();
                            if (style?.Style != null && run != null) {
                                run.Bold = style.Style.Bold == true; run.Italic = style.Style.Italic == true; run.Underline = style.Style.Underline == true;
                                run.FontSize = style.Style.FontSize == null ? null : (int?)Math.Round(ToPoints(style.Style.FontSize)); run.FontName = style.Style.FontFamily;
                                if (Uri.TryCreate(style.Style.Link?.Url, UriKind.Absolute, out Uri? link)) run.Hyperlink = link;
                            }
                        } else if (element.Shape != null) {
                            if (TryMapShapeType(element.Shape.ShapeType, out A.ShapeTypeValues shapeType)) {
                                slide.AddShapePoints(shapeType, left, top, Math.Max(1, width), Math.Max(1, height), element.ObjectId);
                            } else {
                                report.Add(
                                    TranslationSeverity.Warning,
                                    "Shapes",
                                    $"Native Google Slides shape '{element.Shape.ShapeType ?? "unspecified"}' could not be mapped to PowerPoint geometry.",
                                    code: "SLIDES.IMPORT.SHAPE_UNSUPPORTED",
                                    action: TranslationAction.Skip);
                            }
                        } else if (element.Table != null && element.Table.Rows > 0 && element.Table.Columns > 0) {
                            PowerPointTable table = slide.AddTablePoints(element.Table.Rows, element.Table.Columns, left, top, Math.Max(1, width), Math.Max(1, height));
                            for (int row = 0; row < Math.Min(table.RowItems.Count, element.Table.TableRows.Count); row++) {
                                for (int column = 0; column < Math.Min(table.RowItems[row].Cells.Count, element.Table.TableRows[row].TableCells.Count); column++) {
                                    table.RowItems[row].Cells[column].Text = ExtractText(element.Table.TableRows[row].TableCells[column].Text);
                                }
                            }
                        } else if (element.Image?.ContentUrl is string url && !string.IsNullOrWhiteSpace(url)) {
                            try {
                                byte[] bytes = await transport.SendBytesAsync(
                                    token,
                                    HttpMethod.Get,
                                    url,
                                    GoogleWorkspaceRequestSafety.Safe,
                                    "Google Slides image",
                                    report,
                                    cancellationToken,
                                    preserveRequestUri: true).ConfigureAwait(false);
                                using var image = new MemoryStream(bytes, writable: false);
                                slide.AddPicturePoints(image, DetectImageType(bytes), left, top, Math.Max(1, width), Math.Max(1, height));
                            } catch (Exception ex) when (!(ex is OperationCanceledException)) {
                                report.Add(TranslationSeverity.Warning, "Images", "A native Google Slides image could not be downloaded; Drive-export import remains the broad fallback.", code: "SLIDES.IMPORT.IMAGE_FALLBACK", action: TranslationAction.Skip);
                            }
                        }
                    }
                    string? notesId = sourceSlide.SlideProperties?.NotesPage?.NotesProperties?.SpeakerNotesObjectId;
                    GoogleSlidesApiPageElement? notesShape = sourceSlide.SlideProperties?.NotesPage?.PageElements.FirstOrDefault(element => string.Equals(element.ObjectId, notesId, StringComparison.Ordinal));
                    string notes = ExtractText(notesShape?.Shape?.Text);
                    if (!string.IsNullOrWhiteSpace(notes)) slide.Notes.Text = notes;
                }
                presentation.Save();
                return presentation;
            } catch { presentation.Dispose(); stream.Dispose(); throw; }
        }

        private static string ExtractText(GoogleSlidesApiTextContent? text) => text == null ? string.Empty : string.Concat(text.TextElements.Select(element => element.TextRun?.Content));
        private static double ToPoints(GoogleSlidesApiDimension? dimension) => dimension == null ? 0 : ToPoints(dimension.Magnitude, dimension.Unit);
        private static double ToPoints(double value, string? unit) => string.Equals(unit, "EMU", StringComparison.OrdinalIgnoreCase) ? value / 12700d : value;
        private static string ToHex(GoogleSlidesApiRgbColor color) => $"{ToByte(color.Red):X2}{ToByte(color.Green):X2}{ToByte(color.Blue):X2}";
        private static int ToByte(double component) => Math.Max(0, Math.Min(255, (int)Math.Round(component * 255d)));
        private static bool TryMapShapeType(string? shapeType, out A.ShapeTypeValues mapped) {
            switch (shapeType) {
                case "RECTANGLE": mapped = A.ShapeTypeValues.Rectangle; return true;
                case "ROUND_RECTANGLE": mapped = A.ShapeTypeValues.RoundRectangle; return true;
                case "ELLIPSE": mapped = A.ShapeTypeValues.Ellipse; return true;
                case "TRIANGLE": mapped = A.ShapeTypeValues.Triangle; return true;
                case "RIGHT_TRIANGLE": mapped = A.ShapeTypeValues.RightTriangle; return true;
                case "PARALLELOGRAM": mapped = A.ShapeTypeValues.Parallelogram; return true;
                case "TRAPEZOID": mapped = A.ShapeTypeValues.Trapezoid; return true;
                case "DIAMOND": mapped = A.ShapeTypeValues.Diamond; return true;
                case "RIGHT_ARROW": mapped = A.ShapeTypeValues.RightArrow; return true;
                default: mapped = default; return false;
            }
        }

        private static ImagePartType DetectImageType(byte[] bytes) {
            if (bytes.Length >= 8
                && bytes[0] == 0x89
                && bytes[1] == 0x50
                && bytes[2] == 0x4E
                && bytes[3] == 0x47) {
                return ImagePartType.Png;
            }
            if (bytes.Length >= 6
                && bytes[0] == 0x47
                && bytes[1] == 0x49
                && bytes[2] == 0x46
                && bytes[3] == 0x38
                && (bytes[4] == 0x37 || bytes[4] == 0x39)
                && bytes[5] == 0x61) {
                return ImagePartType.Gif;
            }
            return ImagePartType.Jpeg;
        }

        private static void EnsurePresentation(GoogleDriveFile file, string id) {
            if (!string.Equals(file.MimeType, GoogleDriveMimeTypes.Presentation, StringComparison.Ordinal)) throw new InvalidOperationException($"Drive file '{id}' is not a Google presentation.");
        }

        private static void EnsureDownloadable(GoogleDriveFile file, string id) {
            if (file.Capabilities != null && !file.Capabilities.CanDownload) throw new InvalidOperationException($"Drive file '{id}' cannot be exported by the current principal.");
        }

        private static GooglePresentationReference Reference(GoogleDriveFile file, string id, TranslationReport report, GoogleSlidesApiPresentationResponse? native = null) => new GooglePresentationReference {
            PresentationId = native?.PresentationId ?? file.Id ?? id, FileId = file.Id ?? id, Name = native?.Title ?? file.Name, MimeType = file.MimeType,
            WebViewLink = file.WebViewLink, RevisionId = native?.RevisionId, DriveVersion = file.Version, ModifiedTime = file.ModifiedTime, Report = report,
        };
    }
}
