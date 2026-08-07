using CoreAsset = OfficeIMO.OfficeDocumentModelAsset;
using CoreBlock = OfficeIMO.OfficeDocumentModelBlock;
using CoreDiagnostic = OfficeIMO.OfficeDocumentModelDiagnostic;
using CoreForm = OfficeIMO.OfficeDocumentModelFormField;
using CoreLink = OfficeIMO.OfficeDocumentModelLink;
using CoreLocation = OfficeIMO.OfficeDocumentModelLocation;
using CoreMetadata = OfficeIMO.OfficeDocumentModelMetadataEntry;
using CoreModel = OfficeIMO.OfficeDocumentModel;
using CorePage = OfficeIMO.OfficeDocumentModelPage;
using CoreRegion = OfficeIMO.OfficeDocumentModelRegion;
using CoreTable = OfficeIMO.OfficeDocumentModelTable;
using ReaderAsset = OfficeIMO.Reader.OfficeDocumentAsset;
using ReaderBlock = OfficeIMO.Reader.OfficeDocumentBlock;
using ReaderDiagnostic = OfficeIMO.Reader.OfficeDocumentDiagnostic;
using ReaderDiagnosticCategory = OfficeIMO.Reader.OfficeDocumentDiagnosticCategory;
using ReaderDiagnosticSeverity = OfficeIMO.Reader.OfficeDocumentDiagnosticSeverity;
using ReaderForm = OfficeIMO.Reader.OfficeDocumentFormField;
using ReaderLink = OfficeIMO.Reader.OfficeDocumentLink;
using ReaderMetadata = OfficeIMO.Reader.OfficeDocumentMetadataEntry;
using ReaderPage = OfficeIMO.Reader.OfficeDocumentPage;
using ReaderRegion = OfficeIMO.Reader.OfficeDocumentRegion;
using ReaderSource = OfficeIMO.Reader.OfficeDocumentSource;

namespace OfficeIMO.Reader.Visio;

internal static partial class VisioReaderAdapter {
    private static OfficeDocumentReadResult BuildReaderResult(
        CoreModel model,
        SourceMetadata source,
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<ReaderVisual> visuals) {
        var blocks = model.Blocks.ToDictionary(static item => item, MapBlock);
        var tables = model.Tables.ToDictionary(static item => item, MapTable);
        var assets = model.Assets.ToDictionary(static item => item, MapAsset);
        var links = model.Links.ToDictionary(static item => item, MapLink);
        var forms = model.Forms.ToDictionary(static item => item, MapForm);

        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Visio,
            Source = new ReaderSource {
                Path = source.Path,
                SourceId = source.SourceId,
                SourceHash = source.SourceHash,
                LastWriteUtc = source.LastWriteUtc,
                LengthBytes = source.LengthBytes,
                Title = model.Source.Title,
                Author = model.Source.Author,
                Subject = model.Source.Subject,
                Keywords = model.Source.Keywords
            },
            CapabilitiesUsed = BuildReaderCapabilities(model.CapabilitiesUsed),
            Markdown = chunks.Count == 0
                ? null
                : string.Join(Environment.NewLine + Environment.NewLine, chunks.Select(static chunk => chunk.Markdown ?? chunk.Text)),
            Chunks = chunks,
            Metadata = model.Metadata.Select(MapMetadata).ToArray(),
            Pages = model.Pages.Select(page => MapPage(page, blocks, tables, assets, links, forms)).ToArray(),
            Blocks = blocks.Values.ToArray(),
            Tables = tables.Values.ToArray(),
            Assets = assets.Values.ToArray(),
            Links = links.Values.ToArray(),
            Forms = forms.Values.ToArray(),
            Visuals = visuals,
            Diagnostics = model.Diagnostics.Select(MapDiagnostic).ToArray()
        };
    }

    private static IReadOnlyList<string> BuildReaderCapabilities(IReadOnlyList<string> modelCapabilities) {
        var capabilities = new List<string> {
            "officeimo.reader.visio",
            "officeimo.reader.visio.rich-v5"
        };
        foreach (string capability in modelCapabilities) {
            if (!capabilities.Contains(capability, StringComparer.Ordinal)) capabilities.Add(capability);
        }
        return capabilities;
    }

    private static ReaderPage MapPage(
        CorePage page,
        IReadOnlyDictionary<CoreBlock, ReaderBlock> blocks,
        IReadOnlyDictionary<CoreTable, ReaderTable> tables,
        IReadOnlyDictionary<CoreAsset, ReaderAsset> assets,
        IReadOnlyDictionary<CoreLink, ReaderLink> links,
        IReadOnlyDictionary<CoreForm, ReaderForm> forms) => new ReaderPage {
            Number = page.Number,
            Name = page.Name,
            Width = page.Width,
            Height = page.Height,
            RotationDegrees = page.RotationDegrees,
            Location = MapLocation(page.Location),
            Blocks = page.Blocks.Select(item => blocks.TryGetValue(item, out ReaderBlock? mapped) ? mapped : MapBlock(item)).ToArray(),
            Tables = page.Tables.Select(item => tables.TryGetValue(item, out ReaderTable? mapped) ? mapped : MapTable(item)).ToArray(),
            Assets = page.Assets.Select(item => assets.TryGetValue(item, out ReaderAsset? mapped) ? mapped : MapAsset(item)).ToArray(),
            Links = page.Links.Select(item => links.TryGetValue(item, out ReaderLink? mapped) ? mapped : MapLink(item)).ToArray(),
            Forms = page.Forms.Select(item => forms.TryGetValue(item, out ReaderForm? mapped) ? mapped : MapForm(item)).ToArray()
        };

    private static ReaderBlock MapBlock(CoreBlock block) => new ReaderBlock {
        Id = block.Id,
        Kind = block.Kind,
        Text = block.Text,
        Level = block.Level,
        Marker = block.Marker,
        Location = MapLocation(block.Location),
        Region = MapRegion(block.Region)
    };

    private static ReaderTable MapTable(CoreTable table) => new ReaderTable {
        Title = table.Title,
        Kind = table.Kind,
        Summary = table.Summary,
        PayloadHash = table.PayloadHash,
        Location = MapLocation(table.Location),
        Columns = table.Columns,
        ColumnProfiles = ReaderTableProfiler.CreateProfiles(table.Columns, table.Rows),
        Rows = table.Rows,
        TotalRowCount = table.TotalRowCount,
        Truncated = table.Truncated
    };

    private static ReaderAsset MapAsset(CoreAsset asset) => new ReaderAsset {
        Id = asset.Id,
        Kind = asset.Kind,
        MediaType = asset.MediaType,
        Extension = asset.Extension,
        FileName = asset.FileName,
        AltText = asset.AltText,
        Title = asset.Title,
        Width = asset.Width,
        Height = asset.Height,
        LengthBytes = asset.LengthBytes,
        PayloadHash = asset.PayloadHash,
        PayloadBytes = asset.PayloadBytes,
        SourceObjectId = asset.SourceObjectId,
        Region = MapRegion(asset.Region),
        Location = MapLocation(asset.Location)
    };

    private static ReaderLink MapLink(CoreLink link) => new ReaderLink {
        Id = link.Id,
        Kind = link.Kind,
        Uri = link.Uri,
        DestinationName = link.DestinationName,
        DestinationPageNumber = link.DestinationPageNumber,
        DestinationMode = link.DestinationMode,
        NamedAction = link.NamedAction,
        RemoteFile = link.RemoteFile,
        RemoteDestinationName = link.RemoteDestinationName,
        RemoteDestinationPageNumber = link.RemoteDestinationPageNumber,
        Text = link.Text,
        Location = MapLocation(link.Location),
        Region = MapRegion(link.Region)
    };

    private static ReaderForm MapForm(CoreForm form) => new ReaderForm {
        Id = form.Id,
        Name = form.Name,
        Kind = form.Kind,
        Value = form.Value,
        IsReadOnly = form.IsReadOnly,
        IsRequired = form.IsRequired,
        Location = MapLocation(form.Location),
        Region = MapRegion(form.Region)
    };

    private static ReaderMetadata MapMetadata(CoreMetadata metadata) => new ReaderMetadata {
        Id = metadata.Id,
        Category = metadata.Category,
        Name = metadata.Name,
        Value = metadata.Value,
        ValueType = metadata.ValueType,
        SourceObjectId = metadata.SourceObjectId,
        Location = MapLocation(metadata.Location),
        Attributes = metadata.Attributes
    };

    private static ReaderDiagnostic MapDiagnostic(CoreDiagnostic diagnostic) => new ReaderDiagnostic {
        Severity = diagnostic.Severity switch {
            OfficeIMO.OfficeDocumentModelDiagnosticSeverity.Information => ReaderDiagnosticSeverity.Information,
            OfficeIMO.OfficeDocumentModelDiagnosticSeverity.Error => ReaderDiagnosticSeverity.Error,
            _ => ReaderDiagnosticSeverity.Warning
        },
        Category = diagnostic.Category switch {
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Detection => ReaderDiagnosticCategory.Detection,
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Input => ReaderDiagnosticCategory.Input,
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Parsing => ReaderDiagnosticCategory.Parsing,
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Content => ReaderDiagnosticCategory.Content,
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Security => ReaderDiagnosticCategory.Security,
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Limit => ReaderDiagnosticCategory.Limit,
            OfficeIMO.OfficeDocumentModelDiagnosticCategory.Adapter => ReaderDiagnosticCategory.Adapter,
            _ => ReaderDiagnosticCategory.General
        },
        Code = diagnostic.Code,
        Message = diagnostic.Message,
        Source = diagnostic.Source,
        IsRecoverable = diagnostic.IsRecoverable,
        Location = MapLocation(diagnostic.Location),
        Attributes = diagnostic.Attributes
    };

    private static ReaderLocation MapLocation(CoreLocation? location) => location == null
        ? new ReaderLocation()
        : new ReaderLocation {
            Path = location.Path,
            BlockIndex = location.BlockIndex,
            SourceBlockIndex = location.SourceBlockIndex,
            StartLine = location.StartLine,
            EndLine = location.EndLine,
            NormalizedStartLine = location.NormalizedStartLine,
            NormalizedEndLine = location.NormalizedEndLine,
            HeadingPath = location.HeadingPath,
            HeadingSlug = location.HeadingSlug,
            SourceBlockKind = location.SourceBlockKind,
            BlockAnchor = location.BlockAnchor,
            Sheet = location.Sheet,
            A1Range = location.A1Range,
            Slide = location.Slide,
            Page = location.Page,
            TableIndex = location.TableIndex
        };

    private static ReaderRegion? MapRegion(CoreRegion? region) => region == null
        ? null
        : new ReaderRegion {
            X = region.X,
            Y = region.Y,
            Width = region.Width,
            Height = region.Height
        };
}
