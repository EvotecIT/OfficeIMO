using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

internal static class PdfEditorCommandExecutor {
    internal static byte[] Apply(byte[] pdf, PdfEditorCommand command) {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(command);
        return command.Tool switch {
            PdfEditorTool.Note => AddAnnotation(pdf, command, "Text", iconName: "Comment", createPopup: true),
            PdfEditorTool.FreeText => AddAnnotation(pdf, command, "FreeText"),
            PdfEditorTool.Highlight => AddMarkup(pdf, command, "Highlight"),
            PdfEditorTool.Underline => AddMarkup(pdf, command, "Underline"),
            PdfEditorTool.StrikeOut => AddMarkup(pdf, command, "StrikeOut"),
            PdfEditorTool.Rectangle => AddAnnotation(pdf, command, "Square"),
            PdfEditorTool.Ellipse => AddAnnotation(pdf, command, "Circle"),
            PdfEditorTool.Line => AddLine(pdf, command),
            PdfEditorTool.Ink => AddInk(pdf, command),
            PdfEditorTool.Stamp => AddAnnotation(pdf, command, "Stamp", iconName: command.Properties.StampName),
            PdfEditorTool.AddText => AddText(pdf, command),
            PdfEditorTool.AddImage => AddImage(pdf, command),
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
        string? removedTextMarker = null) {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(plan);
        PdfDocument source = PdfDocument.Load(pdf);
        PdfDocument redacted = source.Redactions.Apply(plan, new PdfRedactionApplyOptions {
            PaintUnmatchedAreas = true,
            UnsupportedImagePolicy = PdfRedactionUnsupportedImagePolicy.RemoveWholePlacement,
            RemoveIntersectingPaths = true
        });
        var verificationOptions = new PdfRedactionVerificationOptions {
            CheckManagedRendering = true,
            FailOnUndecodablePdfStreams = true,
            RequireCompleteStreamInspection = true
        };
        if (!string.IsNullOrWhiteSpace(removedTextMarker)) {
            verificationOptions.RequireRemovedText(removedTextMarker.Trim());
        }
        PdfRedactionVerificationReport verification = redacted.Redactions.AssertAppliedPlan(plan, verificationOptions);
        return new PdfVerifiedRedactionResult(redacted.ToBytes(), plan, verification);
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

    private static byte[] AddAnnotation(
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
        return result.Bytes;
    }

    private static byte[] AddMarkup(byte[] pdf, PdfEditorCommand command, string subtype) {
        PdfPageRectangle bounds = command.Bounds;
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = subtype,
            Rectangle = Rectangle(bounds),
            QuadPoints = new[] {
                bounds.Left, bounds.Top,
                bounds.Right, bounds.Top,
                bounds.Left, bounds.Bottom,
                bounds.Right, bounds.Bottom
            },
            Contents = command.Properties.Text,
            Title = command.Properties.Author,
            Color = Color(command.Properties.Color),
            GenerateAppearance = true
        });
        return result.Bytes;
    }

    private static byte[] AddLine(byte[] pdf, PdfEditorCommand command) {
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
        return result.Bytes;
    }

    private static byte[] AddInk(byte[] pdf, PdfEditorCommand command) {
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
        return result.Bytes;
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

    private static byte[] AddLink(byte[] pdf, PdfEditorCommand command) {
        PdfAnnotationEditResult result = PdfDocument.Load(pdf).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = command.PageNumber,
            Subtype = "Link",
            Rectangle = Rectangle(command.Bounds),
            Contents = command.Properties.Text,
            LinkUri = command.Properties.LinkUri,
            GenerateAppearance = false
        });
        return result.Bytes;
    }

    private static double[] Rectangle(PdfPageRectangle bounds) =>
        new[] { bounds.Left, bounds.Bottom, bounds.Right, bounds.Top };

    private static double[] Color(PdfColor color) => new[] { color.R, color.G, color.B };
}
