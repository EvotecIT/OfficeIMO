using OfficeIMO.Drawing;
using OfficeIMO.Publisher;

namespace OfficeIMO.Reader.Publisher;

internal sealed class PublisherReadProjection {
    private readonly PublisherDocument _source;
    private readonly string _path;
    private readonly ReaderOptions _settings;
    private readonly ReaderPublisherOptions _options;
    private readonly CancellationToken _token;
    private readonly List<OfficeDocumentBlock> _blocks = new();
    private readonly List<ReaderChunk> _chunks = new();
    private readonly List<OfficeDocumentAsset> _assets = new();
    private readonly List<OfficeDocumentDiagnostic> _diagnostics = new();
    private long _items, _characters;

    internal PublisherReadProjection(PublisherDocument source, string path, ReaderOptions settings,
        ReaderPublisherOptions options, CancellationToken token) {
        _source = source; _path = path; _settings = settings; _options = options; _token = token;
    }

    internal OfficeDocumentReadResult Build() {
        foreach (OfficeConversionFidelityDiagnostic finding in _source.ReadReport.FidelityDiagnostics) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Code = finding.Code, Message = finding.Message, Source = finding.Source,
                Severity = finding.LossKind == OfficeConversionLossKind.None ? OfficeDocumentDiagnosticSeverity.Information : OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Content, Location = Location(),
                Attributes = new Dictionary<string, string> {
                    ["lossKind"] = finding.LossKind.ToString(), ["sourceLocation"] = finding.Location ?? string.Empty
                }
            });
        }
        _diagnostics.Add(new OfficeDocumentDiagnostic {
            Code = "PUB_READER_LAYOUT_OMITTED", Source = "OfficeIMO.Reader.Publisher",
            Message = "Reader retains complete stories once in native story order. Page placement, frame flow, table structure and typography remain in the Publisher model and are not reconstructed in this semantic projection.",
            Category = OfficeDocumentDiagnosticCategory.Content, Location = Location(),
            Attributes = new Dictionary<string, string> { ["lossKind"] = OfficeConversionLossKind.Omission.ToString() }
        });
        foreach (PublisherTextStory story in _source.TextStories) AddStory(story);
        foreach (PublisherImage image in _source.Images) {
            Item();
            byte[]? payload = _options.IncludeImagePayloads ? image.GetBytes() : null;
            var asset = new OfficeDocumentAsset {
                Id = "publisher-image-" + image.Id.ToString(CultureInfo.InvariantCulture), Kind = "image",
                SourceObjectId = image.Id.ToString(CultureInfo.InvariantCulture), MediaType = image.ContentType,
                LengthBytes = image.ByteCount, PayloadBytes = payload,
                PayloadHash = payload != null && _settings.ComputeHashes ? OfficeDocumentAssetHash.ComputeSha256Hex(payload) : null,
                Location = Location("publisher-image-store")
            };
            _assets.Add(asset);
        }
        string[] warnings = _diagnostics.Where(diagnostic => diagnostic.Severity != OfficeDocumentDiagnosticSeverity.Information)
            .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message).ToArray();
        if (_chunks.Count == 0) {
            Item(); _chunks.Add(new ReaderChunk { Id = "publisher-diagnostics", Kind = ReaderInputKind.Publisher, Location = Location() });
        }
        _chunks[0].Warnings = warnings;
        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Publisher, Source = new OfficeDocumentSource { Path = _path },
            CapabilitiesUsed = new[] { "officeimo.reader.publisher", "officeimo.publisher.native-recovery" },
            Blocks = _blocks.ToArray(), Chunks = _chunks.ToArray(), Assets = _assets.ToArray(), Diagnostics = _diagnostics.ToArray(),
            Markdown = string.Concat(_chunks.Select(chunk => chunk.Markdown)),
            Pages = _source.Pages.Select((page, index) => new OfficeDocumentPage {
                Number = index + 1, Name = page.Name, Width = page.Width, Height = page.Height,
                Location = new ReaderLocation { Path = _path, Page = index + 1, SourceBlockKind = "publisher-page" }
            }).ToArray()
        };
    }

    private void AddStory(PublisherTextStory story) {
        int offset = 0;
        string anchor = "publisher-story-" + story.Id.ToString(CultureInfo.InvariantCulture);
        for (int index = 0; index < story.Paragraphs.Count; index++) {
            _token.ThrowIfCancellationRequested();
            OfficeRichTextParagraph paragraph = story.Paragraphs[index];
            string body = string.Concat(paragraph.Runs.Select(run => run.Text));
            if (body.Length > story.Text.Length - offset || string.CompareOrdinal(story.Text, offset, body, 0, body.Length) != 0)
                throw new InvalidDataException("Publisher paragraph projection does not match its complete source story.");
            int length = body.Length;
            if (offset + length < story.Text.Length && story.Text[offset + length] == '\n') length++;
            string text = story.Text.Substring(offset, length);
            string id = anchor + "-p" + index.ToString("D4", CultureInfo.InvariantCulture);
            string? marker = paragraph.Label?.Run.Text;
            Item();
            ReaderLocation blockLocation = Location("publisher-story-paragraph", anchor, index);
            blockLocation.LogicalOrder = _blocks.Count;
            _blocks.Add(new OfficeDocumentBlock {
                Id = id, Kind = marker == null ? "paragraph" : "list-item", Text = text, Marker = marker,
                Location = blockLocation
            });
            // Leave room for a list marker and paragraph separator after literal escaping,
            // including character references that preserve source indentation.
            int maximum = Math.Max(1, (_settings.MaxChars - 3) / ReaderMarkdownEscaping.MaximumExpansion);
            IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(text, maximum);
            if (parts.Count == 0) parts = new[] { string.Empty };
            for (int part = 0; part < parts.Count; part++) {
                Item();
                string markdown = (marker != null && part == 0 ? "- " : string.Empty)
                    + ReaderMarkdownEscaping.EscapeLiteral(parts[part], _token);
                if (part == parts.Count - 1) markdown += "\n";
                _characters = checked(_characters + parts[part].Length + markdown.Length);
                if (_characters > _options.ReadOptions!.Limits.MaxTextCharacters)
                    throw new InvalidDataException("Publisher Reader projection text limit exceeded.");
                var chunk = new ReaderChunk {
                    Id = id + "-" + part.ToString("D4", CultureInfo.InvariantCulture), Kind = ReaderInputKind.Publisher,
                    Text = parts[part], Markdown = markdown, ContinuesPreviousChunk = part > 0,
                    Location = Location("publisher-story-paragraph", id, index)
                };
                chunk.Location.BlockIndex = _chunks.Count;
                chunk.Location.LogicalOrder = blockLocation.LogicalOrder;
                ReaderReadScope.Current?.Budget?.AddChunk(chunk);
                _chunks.Add(chunk);
            }
            offset += length;
        }
        if (offset != story.Text.Length) throw new InvalidDataException("Publisher Reader projection did not consume its complete source story.");
    }

    private void Item() {
        _token.ThrowIfCancellationRequested();
        if (++_items > _options.ReadOptions!.Limits.MaxItems)
            throw new InvalidDataException("Publisher Reader projection item limit exceeded.");
    }
    private ReaderLocation Location(string? kind = null, string? anchor = null, int? sourceIndex = null) => new() {
        Path = _path, SourceBlockKind = kind, BlockAnchor = anchor, SourceBlockIndex = sourceIndex
    };
}
