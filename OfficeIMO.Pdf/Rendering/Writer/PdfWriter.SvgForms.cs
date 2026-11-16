using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// One SVG drawing written once as a Form XObject in drawing units (y-up, origin at the drawing's bottom left)
    /// and referenced from every placement. It owns its paint resources, so any page can draw it.
    /// </summary>
    private sealed class SvgFormXObject {
        internal SvgFormXObject(int sequence, double width, double height, string content,
            IReadOnlyList<PageGraphicsState> graphicsStates, IReadOnlyList<PageShading> shadings) {
            ResourceName = "/Sv" + sequence.ToString(CultureInfo.InvariantCulture);
            Width = width;
            Height = height;
            Content = content;
            GraphicsStates = graphicsStates;
            Shadings = shadings;
        }

        internal string ResourceName { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal string Content { get; }
        internal IReadOnlyList<PageGraphicsState> GraphicsStates { get; }
        internal IReadOnlyList<PageShading> Shadings { get; }
    }

    /// <summary>The SVG Forms of one document, shared by its layout passes and running stories.</summary>
    private sealed class SvgFormRegistry {
        // Forms stay in memory until the document is written; past these bounds drawings are written inline.
        private const int MaximumForms = 1024;
        private const long MaximumContentLength = 32L * 1024L * 1024L;
        private readonly Dictionary<OfficeDrawing, SvgFormXObject?> _byDrawing = new();
        private readonly Dictionary<string, SvgFormXObject> _byContent = new(StringComparer.Ordinal);
        private long _contentLength;

        internal bool TryGet(OfficeDrawing drawing, out SvgFormXObject? form) => _byDrawing.TryGetValue(drawing, out form);

        /// <summary>Records the Form captured for a drawing, or that it has none, and returns the shared Form.</summary>
        internal SvgFormXObject? Register(OfficeDrawing drawing, double width, double height, string? content,
            IReadOnlyList<PageGraphicsState> graphicsStates, IReadOnlyList<PageShading> shadings) {
            SvgFormXObject? form = null;
            if (content != null) {
                // Equal SVG parsed from different arrays draws identically; share one Form for it.
                string key = BuildContentKey(width, height, content, graphicsStates, shadings);
                if (!_byContent.TryGetValue(key, out form) &&
                    _byContent.Count < MaximumForms && _contentLength + content.Length <= MaximumContentLength) {
                    form = new SvgFormXObject(_byContent.Count + 1, width, height, content, graphicsStates, shadings);
                    _byContent.Add(key, form);
                    _contentLength += content.Length;
                }
            }
            if (_byDrawing.Count < MaximumForms * 4) _byDrawing[drawing] = form;
            return form;
        }

        private static string BuildContentKey(double width, double height, string content,
            IReadOnlyList<PageGraphicsState> graphicsStates, IReadOnlyList<PageShading> shadings) {
            var key = new StringBuilder(content.Length + 64);
            key.Append(width.ToString("R", CultureInfo.InvariantCulture)).Append(' ')
                .Append(height.ToString("R", CultureInfo.InvariantCulture)).Append('\n');
            foreach (PageGraphicsState state in graphicsStates) {
                key.Append(state.Name).Append(' ').Append(state.FillOpacity.ToString("R", CultureInfo.InvariantCulture)).Append(' ')
                    .Append(state.StrokeOpacity.ToString("R", CultureInfo.InvariantCulture)).Append(' ').Append(state.BlendMode).Append('\n');
            }
            foreach (PageShading shading in shadings) {
                key.Append(shading.Name).Append(shading.IsRadial ? " R " : " A ").Append(shading.SpreadMode).Append(' ')
                    .Append(shading.ColorInterpolation).Append(shading.InterpolateAlphaSeparately ? " s" : " j")
                    .Append(shading.IsOpacityData ? " o " : " c ").Append(shading.OutsideColor?.ToString() ?? "-");
                foreach (double value in new[] { shading.X0, shading.Y0, shading.R0, shading.X1, shading.Y1, shading.R1,
                    shading.AlphaLeft, shading.AlphaBottom, shading.AlphaRight, shading.AlphaTop }) {
                    key.Append(' ').Append(value.ToString("R", CultureInfo.InvariantCulture));
                }
                foreach (OfficeGradientStop stop in shading.Stops) {
                    key.Append(' ').Append(stop.Offset.ToString("R", CultureInfo.InvariantCulture)).Append(':').Append(stop.Color.ToString());
                }
                key.Append('\n');
            }
            return key.Append(content).ToString();
        }
    }

    // Writer placeholders for page objects (inline images, effect groups) all start with this comment.
    private const string InlinePlaceholderMarker = "%OIMO_";

    /// <summary>Returns whether content written for an SVG drawing can stand alone in a Form XObject.</summary>
    private static bool IsSelfContainedSvgFormContent(string content) {
        // Only path, colour, state and shading operators may appear: no text, marked content,
        // XObjects or writer placeholders, whose resources and structure belong to a page.
        if (content.Contains(InlinePlaceholderMarker)) return false;
        int index = 0;
        while (index < content.Length) {
            while (index < content.Length && IsPdfWhiteSpace(content[index])) index++;
            int start = index;
            while (index < content.Length && !IsPdfWhiteSpace(content[index])) index++;
            int length = index - start;
            if (length == 0) continue;
            char first = content[start];
            if (first == '(' || first == '<') return false;
            if (length == 2 && (Matches(content, start, "BT") || Matches(content, start, "Do") || Matches(content, start, "BI") ||
                Matches(content, start, "MP") || Matches(content, start, "Tf") || Matches(content, start, "BX") ||
                Matches(content, start, "DP"))) return false;
            if (length == 3 && (Matches(content, start, "BDC") || Matches(content, start, "BMC") || Matches(content, start, "EMC"))) return false;
        }
        return true;

        static bool Matches(string text, int start, string value) => string.CompareOrdinal(text, start, value, 0, value.Length) == 0;
    }

    private static bool IsPdfWhiteSpace(char value) => value is ' ' or '\n' or '\r' or '\t' or '\f' or '\0';

    /// <summary>Returns whether a drawing consists only of shapes, which a self-contained Form can draw.</summary>
    private static bool IsSvgFormCandidate(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingShape) continue;
            if (element is OfficeDrawingGroup group && group.ActualText == null && IsSvgFormCandidate(group.InnerDrawing)) continue;
            return false;
        }
        return true;
    }

    /// <summary>Writes shading objects and their alpha and pad helpers, adding them to the given resource lists.</summary>
    private static void AddShadingResources(IList<byte[]> objects, IEnumerable<PageShading> pageShadings,
        PdfPrintColorTransform? shadingColorTransform, List<(string Name, int Id)> graphicsStates,
        List<(string Name, int Id)> xobjects, List<(string Name, int Id)> shadings, CancellationToken cancellationToken) {
        foreach (PageShading shading in pageShadings) {
            PdfPrintColorTransform? resourceColorTransform = shading.IsOpacityData ? null : shadingColorTransform;
            int shadingId;
            if (shading.SpreadMode != OfficeGradientSpreadMode.Pad) {
                shadingId = AddRadialSpreadShading(objects, shading, alphaOnly: false, resourceColorTransform);
            } else {
                string shadingObject = shading.IsRadial
                    ? PdfVisualResourceDictionaryBuilder.BuildRadialShadingObject(
                        shading.X0,
                        shading.Y0,
                        shading.R0,
                        shading.X1,
                        shading.Y1,
                        shading.R1,
                        shading.Stops,
                        resourceColorTransform, colorInterpolation: shading.ColorInterpolation,
                        interpolateAlphaSeparately: shading.InterpolateAlphaSeparately,
                        functionReferenceResolver: body => PdfSyntaxEscaper.IndirectReference(AddObject(objects, body)))
                    : PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(
                        shading.X0,
                        shading.Y0,
                        shading.X1,
                        shading.Y1,
                        shading.Stops,
                        resourceColorTransform, colorInterpolation: shading.ColorInterpolation,
                        interpolateAlphaSeparately: shading.InterpolateAlphaSeparately,
                        functionReferenceResolver: body => PdfSyntaxEscaper.IndirectReference(AddObject(objects, body)));
                shadingId = AddObject(objects, shadingObject);
            }
            shadings.Add(("/" + shading.Name, shadingId));
            AddGradientAlphaResources(objects, shading, graphicsStates);
            AddRadialPadResources(objects, shading, shadingId, xobjects, resourceColorTransform, cancellationToken);
        }
    }

    /// <summary>Writes an SVG Form with its own paint resources and returns its object number.</summary>
    private static int AddSvgFormXObject(IList<byte[]> objects, SvgFormXObject form,
        PdfPrintColorTransform? colorTransform, CancellationToken cancellationToken) {
        var graphicsStates = new List<(string Name, int Id)>();
        foreach (PageGraphicsState state in form.GraphicsStates) {
            int stateId = AddObject(objects, PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(state.FillOpacity, state.StrokeOpacity, state.BlendMode));
            graphicsStates.Add(("/" + state.Name, stateId));
        }
        var xobjects = new List<(string Name, int Id)>();
        var shadings = new List<(string Name, int Id)>();
        AddShadingResources(objects, form.Shadings, colorTransform, graphicsStates, xobjects, shadings, cancellationToken);
        string content = colorTransform != null ? colorTransform.NormalizeGeneratedContent(form.Content, cancellationToken) : form.Content;
        byte[] contentBytes = PdfEncoding.Latin1GetBytes(content);
        var resources = new StringBuilder();
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "XObject", FilterPdfResources(content, xobjects));
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "ExtGState", FilterPdfResources(content, graphicsStates));
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Shading", FilterPdfResources(content, shadings));
        string dictionary = "<< /Type /XObject /Subtype /Form /FormType 1 /BBox [0 0 " +
            PdfNumberFormatter.Precise(form.Width) + " " + PdfNumberFormatter.Precise(form.Height) +
            "] /Resources <<" + resources + " >> /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>";
        return AddStreamObject(objects, dictionary, contentBytes);
    }
}
