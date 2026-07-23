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
            if (effective.MaxImageBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options),
                    "MaxImageBytes must be greater than zero.");
            }
            return effective.Mode == GoogleSlidesImportMode.DriveExport
                ? await ImportDriveAsync(presentationId, session, effective, cancellationToken).ConfigureAwait(false)
                : await ImportNativeAsync(presentationId, session, effective,
                    cancellationToken).ConfigureAwait(false);
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

        private static async Task<GoogleSlidesImportResult> ImportNativeAsync(string id,
            GoogleWorkspaceSession session, GoogleSlidesImportOptions options,
            CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(id, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsurePresentation(source, id);
            GoogleWorkspaceAccessToken token = await session.AcquireAccessTokenAsync(new[] { GoogleWorkspaceScopeCatalog.PresentationsReadonly }, cancellationToken).ConfigureAwait(false);
            using var transport = new GoogleWorkspaceHttpTransport(session.Options);
            GoogleSlidesApiPresentationResponse response = await transport.SendJsonAsync<GoogleSlidesApiPresentationResponse>(token.AccessToken, HttpMethod.Get,
                $"https://slides.googleapis.com/v1/presentations/{Uri.EscapeDataString(id)}", null, GoogleWorkspaceRequestSafety.Safe, "Google Slides API", report,
                GoogleSlidesJsonSerializerContext.Default.GoogleSlidesApiPresentationResponse, cancellationToken).ConfigureAwait(false);
            PowerPointPresentation presentation = await ProjectAsync(response,
                transport, token.AccessToken, options, report,
                cancellationToken).ConfigureAwait(false);
            report.Add(TranslationSeverity.Info, "NativeImport", "Slides, text boxes, core text styles, tables, images, geometry, and speaker-note text were projected into OfficeIMO.", code: "SLIDES.IMPORT.NATIVE", action: TranslationAction.Preserve);
            return new GoogleSlidesImportResult(presentation, Reference(source, id, report, response), report);
        }

        private static async Task<PowerPointPresentation> ProjectAsync(
            GoogleSlidesApiPresentationResponse source,
            GoogleWorkspaceHttpTransport transport,
            string token,
            GoogleSlidesImportOptions options,
            TranslationReport report,
            CancellationToken cancellationToken) {
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
                    } else if (sourceSlide.PageProperties?.PageBackgroundFill?.StretchedPictureFill?.ContentUrl is string backgroundUrl
                        && !string.IsNullOrWhiteSpace(backgroundUrl)) {
                        try {
                            Uri trustedUrl = GetTrustedImageUrl(backgroundUrl);
                            byte[] bytes = await transport.SendBytesAsync(
                                token,
                                HttpMethod.Get,
                                trustedUrl.AbsoluteUri,
                                GoogleWorkspaceRequestSafety.Safe,
                                "Google Slides background image",
                                report,
                                cancellationToken,
                                preserveRequestUri: true,
                                includeAuthorization: false,
                                maxResponseBytes: options.MaxImageBytes).ConfigureAwait(false);
                            using var image = new MemoryStream(bytes, writable: false);
                            slide.SetBackgroundImage(image, DetectImageType(bytes));
                        } catch (Exception ex) when (!(ex is OperationCanceledException)) {
                            report.Add(
                                TranslationSeverity.Warning,
                                "Backgrounds",
                                $"The image background on slide '{sourceSlide.ObjectId ?? "unspecified"}' could not be downloaded; Drive-export import preserves the original presentation package.",
                                code: "SLIDES.IMPORT.BACKGROUND_IMAGE_FALLBACK",
                                action: TranslationAction.Skip);
                        }
                    }
                    foreach (GoogleSlidesApiPageElement element in sourceSlide.PageElements) {
                        ElementGeometry geometry = ProjectGeometry(element, report);
                        double left = geometry.Left;
                        double top = geometry.Top;
                        double width = geometry.Width;
                        double height = geometry.Height;
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
                            ApplyTransform(box, geometry);
                            ApplyShapeStyle(box, element.Shape, report, element.ObjectId);
                            ApplyTextRuns(box, element.Shape.Text);
                        } else if (element.Shape != null) {
                            if (TryMapShapeType(element.Shape.ShapeType, out A.ShapeTypeValues shapeType)) {
                                PowerPointAutoShape shape = slide.AddShapePoints(shapeType, left, top, Math.Max(1, width), Math.Max(1, height), element.ObjectId);
                                ApplyTransform(shape, geometry);
                                ApplyShapeStyle(shape, element.Shape, report, element.ObjectId);
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
                            ApplyTransform(table, geometry);
                            for (int row = 0; row < Math.Min(table.RowItems.Count, element.Table.TableRows.Count); row++) {
                                for (int column = 0; column < Math.Min(table.RowItems[row].Cells.Count, element.Table.TableRows[row].TableCells.Count); column++) {
                                    table.RowItems[row].Cells[column].Text = ExtractText(element.Table.TableRows[row].TableCells[column].Text);
                                }
                            }
                        } else if (element.Image?.ContentUrl is string url && !string.IsNullOrWhiteSpace(url)) {
                            try {
                                Uri trustedUrl = GetTrustedImageUrl(url);
                                byte[] bytes = await transport.SendBytesAsync(
                                    token,
                                    HttpMethod.Get,
                                    trustedUrl.AbsoluteUri,
                                    GoogleWorkspaceRequestSafety.Safe,
                                    "Google Slides image",
                                    report,
                                    cancellationToken,
                                    preserveRequestUri: true,
                                    includeAuthorization: false,
                                    maxResponseBytes: options.MaxImageBytes).ConfigureAwait(false);
                                using var image = new MemoryStream(bytes, writable: false);
                                PowerPointPicture picture = slide.AddPicturePoints(image, DetectImageType(bytes), left, top, Math.Max(1, width), Math.Max(1, height));
                                ApplyTransform(picture, geometry);
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

        private static string ExtractText(GoogleSlidesApiTextContent? text) {
            if (text == null) return string.Empty;
            string value = string.Concat(text.TextElements.Select(element => element.TextRun?.Content));
            return value.EndsWith("\n", StringComparison.Ordinal) ? value.Substring(0, value.Length - 1) : value;
        }

        private static void ApplyTextRuns(PowerPointTextBox box, GoogleSlidesApiTextContent text) {
            List<GoogleSlidesApiTextRun> sourceRuns = text.TextElements
                .Select(element => element.TextRun)
                .Where(run => run != null)
                .Cast<GoogleSlidesApiTextRun>()
                .ToList();
            if (sourceRuns.Count == 0) return;

            int lastContentRun = sourceRuns.FindLastIndex(run => !string.IsNullOrEmpty(run.Content));
            PowerPointParagraph paragraph = box.Paragraphs.FirstOrDefault() ?? box.AddParagraph();
            PowerPointTextRun targetRun = paragraph.Runs.FirstOrDefault() ?? paragraph.AddRun(string.Empty);
            for (int index = 0; index < sourceRuns.Count; index++) {
                GoogleSlidesApiTextRun sourceRun = sourceRuns[index];
                string content = sourceRun.Content ?? string.Empty;
                if (index == lastContentRun && content.EndsWith("\n", StringComparison.Ordinal)) {
                    content = content.Substring(0, content.Length - 1);
                }

                if (index == 0) {
                    targetRun.Text = content;
                } else {
                    targetRun = paragraph.AddRun(content);
                }
                ApplyTextRunStyle(targetRun, sourceRun.Style);
            }
        }

        private static void ApplyTextRunStyle(PowerPointTextRun run, GoogleSlidesApiTextStyle? style) {
            if (style == null) return;
            run.Bold = style.Bold == true;
            run.Italic = style.Italic == true;
            run.Underline = style.Underline == true;
            if (style.FontSize != null) run.FontSize = (int)Math.Round(ToPoints(style.FontSize));
            if (!string.IsNullOrWhiteSpace(style.FontFamily)) run.FontName = style.FontFamily;
            if (style.ForegroundColor?.OpaqueColor?.RgbColor is GoogleSlidesApiRgbColor textColor) run.Color = ToHex(textColor);
            if (Uri.TryCreate(style.Link?.Url, UriKind.Absolute, out Uri? link)) run.Hyperlink = link;
        }
        private static double ToPoints(GoogleSlidesApiDimension? dimension) => dimension == null ? 0 : ToPoints(dimension.Magnitude, dimension.Unit);
        private static double ToPoints(double value, string? unit) => string.Equals(unit, "EMU", StringComparison.OrdinalIgnoreCase) ? value / 12700d : value;
        private static ElementGeometry ProjectGeometry(GoogleSlidesApiPageElement element, TranslationReport report) {
            GoogleSlidesApiTransform? transform = element.Transform;
            double matrixScaleX = transform?.ScaleX ?? 1;
            double matrixScaleY = transform?.ScaleY ?? 1;
            double shearX = transform?.ShearX ?? 0;
            double shearY = transform?.ShearY ?? 0;
            double scaleX = Math.Sqrt((matrixScaleX * matrixScaleX) + (shearY * shearY));
            double scaleY = Math.Sqrt((shearX * shearX) + (matrixScaleY * matrixScaleY));
            double rotationRadians = scaleX > 0.000000001
                ? Math.Atan2(shearY, matrixScaleX)
                : Math.Atan2(-shearX, matrixScaleY);
            double determinant = (matrixScaleX * matrixScaleY) - (shearX * shearY);
            bool verticalFlip = determinant < 0;
            double cos = Math.Cos(rotationRadians);
            double sin = Math.Sin(rotationRadians);
            double normalizedScaleX = cos;
            double normalizedShearY = sin;
            double normalizedShearX = verticalFlip ? sin : -sin;
            double normalizedScaleY = verticalFlip ? -cos : cos;
            double tolerance = 0.0000001 * Math.Max(1, Math.Max(scaleX, scaleY));
            bool exact = scaleX > 0.000000001
                && scaleY > 0.000000001
                && Math.Abs(matrixScaleX - (normalizedScaleX * scaleX)) <= tolerance
                && Math.Abs(shearY - (normalizedShearY * scaleX)) <= tolerance
                && Math.Abs(shearX - (normalizedShearX * scaleY)) <= tolerance
                && Math.Abs(matrixScaleY - (normalizedScaleY * scaleY)) <= tolerance;

            double width = ToPoints(element.Size?.Width) * scaleX;
            double height = ToPoints(element.Size?.Height) * scaleY;
            double translateX = ToPoints(transform?.TranslateX ?? 0, transform?.Unit);
            double translateY = ToPoints(transform?.TranslateY ?? 0, transform?.Unit);
            double left = translateX + (normalizedScaleX * width / 2d) + (normalizedShearX * height / 2d) - (width / 2d);
            double top = translateY + (normalizedShearY * width / 2d) + (normalizedScaleY * height / 2d) - (height / 2d);

            if (!exact) {
                report.Add(
                    TranslationSeverity.Warning,
                    "PageElements",
                    $"Native Google Slides element '{element.ObjectId ?? "unspecified"}' uses a skewed or degenerate affine transform that PowerPoint cannot represent exactly; rotation, reflection, size, and position were approximated.",
                    code: "SLIDES.IMPORT.TRANSFORM_PARTIAL",
                    action: TranslationAction.Flatten);
            }

            return new ElementGeometry(left, top, width, height, rotationRadians * (180d / Math.PI), verticalFlip);
        }

        private static void ApplyTransform(PowerPointShape shape, ElementGeometry geometry) {
            if (Math.Abs(geometry.Rotation) > 0.0000001) shape.Rotation = geometry.Rotation;
            if (geometry.VerticalFlip) shape.VerticalFlip = true;
        }

        private static void ApplyShapeStyle(PowerPointShape target, GoogleSlidesApiShape source, TranslationReport report, string? objectId) {
            GoogleSlidesApiShapeBackgroundFill? background = source.ShapeProperties?.ShapeBackgroundFill;
            if (background != null) {
                if (string.Equals(background.PropertyState, "NOT_RENDERED", StringComparison.OrdinalIgnoreCase)) {
                    target.FillColor = "FFFFFF";
                    target.FillTransparency = 100;
                } else if (background.SolidFill?.Color?.RgbColor is GoogleSlidesApiRgbColor fillColor) {
                    target.FillColor = ToHex(fillColor);
                    if (background.SolidFill.Alpha.HasValue) {
                        target.FillTransparency = ToTransparency(background.SolidFill.Alpha.Value);
                    }
                } else if (background.SolidFill != null) {
                    AddShapeStyleDiagnostic(report, objectId, "fill uses a theme color that cannot be resolved without its inherited color scheme");
                }
            }

            GoogleSlidesApiOutline? outline = source.ShapeProperties?.Outline;
            if (outline == null) return;
            if (string.Equals(outline.PropertyState, "NOT_RENDERED", StringComparison.OrdinalIgnoreCase)) {
                target.OutlineColor = null;
                target.OutlineWidthPoints = 0;
                return;
            }

            GoogleSlidesApiSolidFill? outlineFill = outline.OutlineFill?.SolidFill;
            if (outlineFill?.Color?.RgbColor is GoogleSlidesApiRgbColor outlineColor) {
                target.OutlineColor = ToHex(outlineColor);
            } else if (outlineFill != null) {
                AddShapeStyleDiagnostic(report, objectId, "outline uses a theme color that cannot be resolved without its inherited color scheme");
            }
            if (outline.Weight != null) {
                target.OutlineWidthPoints = Math.Max(0, ToPoints(outline.Weight));
            }
            if (outlineFill?.Alpha is double alpha && Math.Abs(alpha - 1d) > 0.0000001) {
                AddShapeStyleDiagnostic(report, objectId, "outline transparency is not representable by the current PowerPoint shape API");
            }
            if (!string.IsNullOrWhiteSpace(outline.DashStyle)
                && !string.Equals(outline.DashStyle, "SOLID", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(outline.DashStyle, "DASH_STYLE_UNSPECIFIED", StringComparison.OrdinalIgnoreCase)) {
                AddShapeStyleDiagnostic(report, objectId, $"outline dash style '{outline.DashStyle}' is not preserved by the Google Slides round-trip model");
            }
        }

        private static void AddShapeStyleDiagnostic(TranslationReport report, string? objectId, string detail) {
            report.Add(
                TranslationSeverity.Warning,
                "Shapes",
                $"Native Google Slides shape '{objectId ?? "unspecified"}' {detail}; Drive-export import remains the fidelity fallback.",
                code: "SLIDES.IMPORT.SHAPE_STYLE_PARTIAL",
                action: TranslationAction.Flatten);
        }

        private readonly struct ElementGeometry {
            internal ElementGeometry(double left, double top, double width, double height, double rotation, bool verticalFlip) {
                Left = left;
                Top = top;
                Width = width;
                Height = height;
                Rotation = rotation;
                VerticalFlip = verticalFlip;
            }

            internal double Left { get; }
            internal double Top { get; }
            internal double Width { get; }
            internal double Height { get; }
            internal double Rotation { get; }
            internal bool VerticalFlip { get; }
        }

        private static string ToHex(GoogleSlidesApiRgbColor color) => $"{ToByte(color.Red):X2}{ToByte(color.Green):X2}{ToByte(color.Blue):X2}";
        private static int ToByte(double component) => Math.Max(0, Math.Min(255, (int)Math.Round(component * 255d)));
        private static int ToTransparency(double alpha) => (int)Math.Round((1d - Math.Max(0d, Math.Min(1d, alpha))) * 100d);
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

        private static Uri GetTrustedImageUrl(string value) {
            if (!Uri.TryCreate(value, UriKind.Absolute, out Uri? uri)
                || !string.Equals(uri.Scheme, Uri.UriSchemeHttps,
                    StringComparison.OrdinalIgnoreCase)
                || !uri.IsDefaultPort && uri.Port != 443
                || !IsGoogleContentHost(uri.Host)) {
                throw new InvalidDataException(
                    "Native Google Slides image URLs must use HTTPS on a Google content host.");
            }
            return uri;
        }

        private static bool IsGoogleContentHost(string host) =>
            string.Equals(host, "googleusercontent.com",
                StringComparison.OrdinalIgnoreCase)
            || host.EndsWith(".googleusercontent.com",
                StringComparison.OrdinalIgnoreCase);

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
