using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

internal static class PdfEditorCommandExecutor {
    internal static byte[] Apply(byte[] pdf, PdfEditorCommand command) {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(command);
        return command.Tool switch {
            PdfEditorTool.AddText => AddText(pdf, command),
            PdfEditorTool.AddImage => AddImage(pdf, command),
            _ => ApplyAnnotation(pdf, command).Bytes
        };
    }

    internal static PdfAnnotationEditResult ApplyAnnotation(byte[] pdf, PdfEditorCommand command) {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(command);
        return command.Tool switch {
            PdfEditorTool.Note => AddAnnotation(pdf, command, "Text", iconName: "Comment", createPopup: true),
            PdfEditorTool.FreeText => AddAnnotation(pdf, command, "FreeText"),
            PdfEditorTool.Highlight => AddMarkup(pdf, command, "Highlight"),
            PdfEditorTool.Underline => AddMarkup(pdf, command, "Underline"),
            PdfEditorTool.StrikeOut => AddMarkup(pdf, command, "StrikeOut"),
            PdfEditorTool.Squiggly => AddMarkup(pdf, command, "Squiggly"),
            PdfEditorTool.Polygon => AddPolygon(pdf, command),
            PdfEditorTool.Rectangle => AddAnnotation(pdf, command, "Square"),
            PdfEditorTool.Ellipse => AddAnnotation(pdf, command, "Circle"),
            PdfEditorTool.Line => AddLine(pdf, command),
            PdfEditorTool.Ink => AddInk(pdf, command),
            PdfEditorTool.Stamp => AddAnnotation(pdf, command, "Stamp", iconName: command.Properties.StampName),
            PdfEditorTool.Link => AddLink(pdf, command),
            PdfEditorTool.SignatureAppearance => AddAnnotation(pdf, command, "FreeText"),
            PdfEditorTool.Redact => throw new InvalidOperationException("Redaction requires the verified redaction workflow."),
            _ => throw new InvalidOperationException("The select tool does not create a PDF mutation.")
        };
    }

    internal static PdfVerifiedRedactionResult ApplyVerifiedRedaction(
        byte[] pdf,
        PdfEditorCommand command,
        string? removedTextMarker = null) {
        if (command.Tool != PdfEditorTool.Redact) throw new ArgumentException("A redaction command is required.", nameof(command));
        PdfRedactionPlan plan = PlanRedaction(pdf, command);
        return ApplyVerifiedRedaction(pdf, plan, removedTextMarker);
    }

    internal static PdfVerifiedRedactionResult ApplyVerifiedRedaction(
        byte[] pdf,
        PdfRedactionPlan plan,
        string? removedTextMarker = null,
        PdfSanitizationOptions? sanitizationOptions = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(plan);
        PdfDocument source = PdfDocument.Load(pdf);
        var applyOptions = new PdfRedactionApplyOptions {
            CancellationToken = cancellationToken,
            PaintUnmatchedAreas = true,
            UnsupportedImagePolicy = PdfRedactionUnsupportedImagePolicy.RemoveWholePlacement,
            RemoveIntersectingPaths = true
        };
        var verificationOptions = new PdfRedactionVerificationOptions {
            CancellationToken = cancellationToken,
            CheckManagedRendering = true,
            FailOnUndecodablePdfStreams = true,
            RequireCompleteStreamInspection = true
        };
        if (!string.IsNullOrWhiteSpace(removedTextMarker)) {
            verificationOptions.RequireRemovedText(removedTextMarker.Trim());
        }
        if (sanitizationOptions is not null) {
            PdfRedactionSharingResult sharing = source.Redactions.ApplyForSharing(plan, sanitizationOptions, applyOptions, verificationOptions);
            return new PdfVerifiedRedactionResult(sharing.ToBytes(), plan, sharing.Redaction, sharing.Summary);
        }
        PdfRedactionApplyResult applied = source.Redactions
            .ApplyWithEvidence(plan, applyOptions, verificationOptions)
            .ThrowIfUnverified();
        return new PdfVerifiedRedactionResult(applied.Pdf, plan, applied.Evidence, applied.Evidence.CreateShareableSummary());
    }

    internal static PdfRedactionPlan PlanRedaction(byte[] pdf, PdfEditorCommand command) {
        if (command.Tool != PdfEditorTool.Redact) throw new ArgumentException("A redaction command is required.", nameof(command));
        var area = new PdfRedactionArea(
            command.PageNumber,
            command.Bounds.Left,
            command.Bounds.Bottom,
            command.Bounds.Width,
            command.Bounds.Height,
            "OfficeIMO Studio area redaction");
        return PdfDocument.Load(pdf).Redactions.Plan(new[] { area });
    }

    private static PdfAnnotationEditResult AddAnnotation(
        byte[] pdf,
        PdfEditorCommand command,
        string subtype,
        string? iconName = null,
        bool createPopup = false) {
        PdfPageRectangle bounds = command.Bounds;
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = subtype,
            Rectangle = Rectangle(bounds),
            Contents = command.Properties.Text,
            Title = command.Properties.Author,
            Color = Color(command.Properties.Color),
            IconName = iconName,
            CreatePopup = createPopup,
            PopupOpen = false,
            GenerateAppearance = true
        });
        return result;
    }

    private static PdfAnnotationEditResult AddMarkup(byte[] pdf, PdfEditorCommand command, string subtype) {
        PdfPageRectangle bounds = command.Bounds;
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = subtype,
            Rectangle = Rectangle(bounds),
            QuadPoints = (command.TextQuads is { Count: > 0 } quads ? quads : new[] { bounds })
                .SelectMany(quad => new[] {
                    quad.Left, quad.Top,
                    quad.Right, quad.Top,
                    quad.Left, quad.Bottom,
                    quad.Right, quad.Bottom
                }).ToArray(),
            Contents = command.Properties.Text,
            Title = command.Properties.Author,
            Color = Color(command.Properties.Color),
            GenerateAppearance = true
        });
        return result;
    }

    private static PdfAnnotationEditResult AddLine(byte[] pdf, PdfEditorCommand command) {
        PdfPagePoint start = command.Path.Count > 0
            ? command.Path[0]
            : new PdfPagePoint(command.Bounds.Left, command.Bounds.Bottom);
        PdfPagePoint end = command.Path.Count > 1
            ? command.Path[^1]
            : new PdfPagePoint(command.Bounds.Right, command.Bounds.Top);
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = "Line",
            Rectangle = Rectangle(command.Bounds),
            Line = new[] { start.X, start.Y, end.X, end.Y },
            Contents = command.Properties.Text,
            Title = command.Properties.Author,
            Color = Color(command.Properties.Color),
            GenerateAppearance = true
        });
        return result;
    }

    // A freehand outline becomes a closed polygon; long paths are thinned so the vertex list stays small.
    private static PdfAnnotationEditResult AddPolygon(byte[] pdf, PdfEditorCommand command) {
        if (command.Path.Count < 3) throw new InvalidOperationException("Draw an outline with at least three points to create a polygon.");
        int step = Math.Max(1, command.Path.Count / 48);
        double[] vertices = command.Path.Where((_, index) => index % step == 0).SelectMany(point => new[] { point.X, point.Y }).ToArray();
        return PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = "Polygon",
            Rectangle = Rectangle(command.Bounds),
            Vertices = vertices,
            Contents = command.Properties.Text,
            Title = command.Properties.Author,
            Color = Color(command.Properties.Color),
            GenerateAppearance = true
        });
    }

    private static PdfAnnotationEditResult AddInk(byte[] pdf, PdfEditorCommand command) {
        if (command.Strokes is { Count: > 0 } strokes) {
            return PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
                PageNumber = command.PageNumber,
                Subtype = "Ink",
                Rectangle = Rectangle(command.Bounds),
                InkPaths = strokes.Where(stroke => stroke.Count > 1)
                    .Select(stroke => (IReadOnlyList<double>)stroke.SelectMany(point => new[] { point.X, point.Y }).ToArray()).ToArray(),
                Contents = command.Properties.Text,
                Title = command.Properties.Author,
                Color = Color(command.Properties.Color),
                GenerateAppearance = true
            });
        }
        if (command.Path.Count < 2) throw new InvalidOperationException("Ink requires a pointer path with at least two points.");
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = "Ink",
            Rectangle = Rectangle(command.Bounds),
            InkPaths = new[] { (IReadOnlyList<double>)command.Path.SelectMany(point => new[] { point.X, point.Y }).ToArray() },
            Contents = command.Properties.Text,
            Title = command.Properties.Author,
            Color = Color(command.Properties.Color),
            GenerateAppearance = true
        });
        return result;
    }

    private static byte[] AddText(byte[] pdf, PdfEditorCommand command) {
        PdfPageRectangle bounds = command.Bounds;
        return PdfDocument.Load(pdf).Stamp.Text(command.Properties.Text, new PdfTextStampOptions {
            PageNumbers = new[] { command.PageNumber },
            X = bounds.Left,
            Y = bounds.Bottom,
            FontSize = command.Properties.FontSize,
            Color = command.Properties.Color
        }).ToBytes();
    }

    private static byte[] AddImage(byte[] pdf, PdfEditorCommand command) {
        byte[] image = command.Properties.ImageBytes
            ?? throw new InvalidOperationException("Choose an image before drawing its placement.");
        PdfPageRectangle bounds = command.Bounds;
        return PdfDocument.Load(pdf).Stamp.Image(image, new PdfImageStampOptions {
            PageNumbers = new[] { command.PageNumber },
            X = bounds.Left,
            Y = bounds.Bottom,
            Width = bounds.Width,
            Height = bounds.Height
        }).ToBytes();
    }

    private static PdfAnnotationEditResult AddLink(byte[] pdf, PdfEditorCommand command) {
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = "Link",
            Rectangle = Rectangle(command.Bounds),
            Contents = command.Properties.Text,
            LinkUri = command.Properties.LinkPageNumber is null ? command.Properties.LinkUri : null,
            LinkPageNumber = command.Properties.LinkPageNumber,
            GenerateAppearance = false
        });
        return result;
    }

    private static double[] Rectangle(PdfPageRectangle bounds) =>
        new[] { bounds.Left, bounds.Bottom, bounds.Right, bounds.Top };

    private static double[] Color(PdfColor color) => new[] { color.R, color.G, color.B };
}
