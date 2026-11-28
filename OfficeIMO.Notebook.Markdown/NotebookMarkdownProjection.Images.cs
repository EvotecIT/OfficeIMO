using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal sealed partial class NotebookMarkdownProjection {
    private const int MaxImagePixels = 16_000_000;
    private static readonly string[] PreferredImageTypes = { "image/svg+xml", "image/png", "image/jpeg" };

    private sealed class CellImages {
        private const string AttachmentPrefix = "attachment:";
        private readonly NotebookMarkdownProjection _projection;
        private readonly NotebookCell _cell;
        private readonly string _location;
        private readonly Dictionary<string, string?> _resolved = new(StringComparer.Ordinal);

        internal CellImages(NotebookMarkdownProjection projection, NotebookCell cell, string location) {
            _projection = projection;
            _cell = cell;
            _location = location;
        }

        internal string? Resolve(string source) {
            _projection._cancellationToken.ThrowIfCancellationRequested();
            if (_resolved.TryGetValue(source, out string? cached)) return cached;
            string? result = source.StartsWith(AttachmentPrefix, StringComparison.Ordinal)
                ? ResolveAttachment(source.Substring(AttachmentPrefix.Length))
                : _projection.RejectImage(_location);
            _resolved.Add(source, result);
            return result;
        }

        private string? ResolveAttachment(string name) {
            if (!_cell.Attachments.TryGetValue(name, out IReadOnlyDictionary<string, JsonElement>? representations)) {
                string unescaped = Uri.UnescapeDataString(name);
                _cell.Attachments.TryGetValue(unescaped, out representations);
            }
            if (representations != null) {
                foreach (string mime in PreferredImageTypes) {
                    if (!representations.TryGetValue(mime, out JsonElement value)) continue;
                    if (_projection.TryGetImageUri(mime, value, out string imageUri)) return imageUri;
                    _projection.AddUnsupported("notebook-attachment-representation", _location,
                        "The referenced " + mime + " attachment representation is invalid or exceeds the conversion limit.");
                }
            }
            _projection.AddUnsupported("notebook-attachment-unavailable", _location,
                "The referenced attachment has no supported image representation in this cell.");
            return null;
        }
    }

    private bool TryGetImageUri(string mime, JsonElement item, out string uri) {
        uri = string.Empty;
        if (mime == "image/svg+xml") {
            if (!TryReadMultiline(item, _options.MaxImageBytes, out string svg)
                || Encoding.UTF8.GetByteCount(svg) > _options.MaxImageBytes) return false;
            byte[] svgBytes = Encoding.UTF8.GetBytes(svg);
            var svgOptions = new OfficeSvgDrawingReaderOptions { CancellationToken = _cancellationToken };
            if (!OfficeSvgDrawingReader.TryRead(svgBytes, svgOptions, out OfficeDrawing? drawing, out int omissions)
                || drawing == null || drawing.Elements.Count == 0 || omissions != 0) return false;
            // Re-export the bounded drawing rather than letting its original XML flow into
            // downloadable HTML or office documents as a data-URI resource.
            byte[] safeSvg = OfficeDrawingSvgExporter.ToSvgBytes(drawing);
            if (safeSvg.Length > _options.MaxImageBytes) return false;
            uri = "data:image/svg+xml;base64," + Convert.ToBase64String(safeSvg);
            return true;
        }

        long base64Limit = (_options.MaxImageBytes + 2L) / 3L * 4L;
        if (!TryReadMultiline(item, (int)Math.Min(base64Limit, int.MaxValue), out string encoded)) return false;
        byte[] bytes;
        try {
            bytes = Convert.FromBase64String(encoded);
        } catch (FormatException) {
            return false;
        }
        if (bytes.Length > _options.MaxImageBytes
            || !OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo info)
            || info.Width < 1 || info.Height < 1 || (long)info.Width * info.Height > MaxImagePixels
            || !string.Equals(info.MimeType, mime, StringComparison.Ordinal)) return false;
        uri = "data:" + mime + ";base64," + (encoded.Any(char.IsWhiteSpace) ? Convert.ToBase64String(bytes) : encoded);
        return true;
    }
}
