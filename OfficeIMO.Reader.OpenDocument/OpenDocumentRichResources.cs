using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Shared bounded projection of native ODF image references and hyperlinks.</summary>
internal sealed class OpenDocumentRichResources {
    private readonly OdfDocument _document;
    private readonly ReaderOpenDocumentOptions _options;
    private readonly OfficeDocumentReadResult _result;
    private readonly Dictionary<string, byte[]> _payloads = new(StringComparer.Ordinal);
    private readonly List<OfficeDocumentLink> _links = new();
    private readonly List<OfficeDocumentAsset> _assets = new();
    private readonly CancellationToken _token;
    private long _retainedBytes;
    private long _linkCharacters;
    private readonly HashSet<string> _entries;

    internal OpenDocumentRichResources(OdfDocument document, ReaderOpenDocumentOptions options, OfficeDocumentReadResult result, CancellationToken token) {
        if (options.MaxImageAssets < 0 || options.MaxImageBytes < 0 || options.MaxTotalImageBytes < 0 || options.MaxLinks < 0 || options.MaxLinkCharacters < 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Image limits cannot be negative.");
        _document = document; _options = options; _result = result; _token = token;
        _entries = new HashSet<string>(document.PackageEntries, StringComparer.Ordinal);
    }

    internal void AddLink(string target, string text, ReaderLocation location) {
        _token.ThrowIfCancellationRequested();
        if (_links.Count >= _options.MaxLinks) throw new ReaderResourceLimitException("OpenDocument.MaxLinks", _options.MaxLinks);
        long characters = (long)target.Length + text.Length;
        if (characters > _options.MaxLinkCharacters - _linkCharacters) throw new ReaderResourceLimitException("OpenDocument.MaxLinkCharacters", _options.MaxLinkCharacters);
        _linkCharacters += characters;
        bool local = target.StartsWith("#", StringComparison.Ordinal);
        _links.Add(new OfficeDocumentLink {
            Id = "odf-link-" + _links.Count.ToString("D6", CultureInfo.InvariantCulture), Text = text,
            Kind = local ? "internal" : "uri", Uri = local ? null : target,
            DestinationName = local ? target.Substring(1) : null, Location = location
        });
    }

    internal void AddImage(string reference, ReaderLocation location, OfficeDocumentRegion? region = null) {
        _token.ThrowIfCancellationRequested();
        if (_assets.Count >= _options.MaxImageAssets) throw new ReaderResourceLimitException("OpenDocument.MaxImageAssets", _options.MaxImageAssets);
        byte[]? bytes = null; string? mediaType = null;
        string path = reference;
        if (!Uri.TryCreate(reference, UriKind.Absolute, out _)) {
            path = OdfDocument.NormalizePackageEntryPath(reference);
            if (!_entries.Contains(path)) {
                AddDiagnostic("ODF_IMAGE_RESOURCE_UNAVAILABLE", "An image reference has no loaded package resource; its descriptor was retained without a payload.", location);
            } else if (!_payloads.TryGetValue(path, out bytes)) {
                long length = _document.GetPackageEntryByteLength(path);
                if (length > _options.MaxImageBytes) throw new ReaderResourceLimitException("OpenDocument.MaxImageBytes", _options.MaxImageBytes);
                if (length > _options.MaxTotalImageBytes - _retainedBytes) throw new ReaderResourceLimitException("OpenDocument.MaxTotalImageBytes", _options.MaxTotalImageBytes);
                bytes = _document.GetPackageEntryBytes(path, _options.MaxImageBytes);
                _retainedBytes = checked(_retainedBytes + bytes.LongLength); _payloads.Add(path, bytes);
            }
            if (_entries.Contains(path)) mediaType = _document.GetPackageEntryMediaType(path);
        } else {
            AddDiagnostic("ODF_EXTERNAL_IMAGE_NOT_FETCHED", "External image resource was retained as metadata without accessing its target.", location);
        }
        var asset = new OfficeDocumentAsset {
            Id = "odf-image-" + _assets.Count.ToString("D6", CultureInfo.InvariantCulture), Kind = "image",
            FileName = Path.GetFileName(path), Extension = Path.GetExtension(path), MediaType = mediaType,
            SourceObjectId = reference, Location = location, Region = region, PayloadBytes = bytes,
            LengthBytes = bytes?.LongLength, PayloadHash = bytes == null ? null : OfficeDocumentAssetHash.ComputeSha256Hex(bytes)
        };
        if (bytes != null && OfficeDocumentImageDimensions.TryReadPixelDimensions(bytes, mediaType, out int width, out int height)) { asset.Width = width; asset.Height = height; }
        _assets.Add(asset);
    }

    internal void AddDiagnostic(string code, string message, ReaderLocation location) =>
        _result.Diagnostics = _result.Diagnostics.Concat(new[] { new OfficeDocumentDiagnostic {
            Code = code, Message = message, Category = OfficeDocumentDiagnosticCategory.Content,
            Source = "officeimo.reader.opendocument", Location = location
        }}).ToArray();

    internal void Complete() {
        _result.Links = _links; _result.Assets = _assets;
        foreach (OfficeDocumentPage page in _result.Pages) {
            page.Links = _links.Where(link => Matches(link.Location, page.Location)).ToArray();
            page.Assets = _assets.Where(asset => Matches(asset.Location, page.Location)).ToArray();
        }
    }

    internal static OfficeDocumentRegion Region(OdfRect bounds) => new() {
        X = bounds.X.ToPoints(), Y = bounds.Y.ToPoints(), Width = bounds.Width.ToPoints(), Height = bounds.Height.ToPoints()
    };

    private static bool Matches(ReaderLocation left, ReaderLocation right) =>
        right.Slide.HasValue ? left.Slide == right.Slide : right.Page.HasValue ? left.Page == right.Page :
        !string.IsNullOrWhiteSpace(right.Sheet) && string.Equals(left.Sheet, right.Sheet, StringComparison.Ordinal);
}
