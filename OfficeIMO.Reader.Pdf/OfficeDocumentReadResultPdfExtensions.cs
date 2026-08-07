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
using CoreSource = OfficeIMO.OfficeDocumentModelSource;
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

namespace OfficeIMO.Reader;

/// <summary>Thin compatibility bridge from Reader results to the PDF-owned neutral-model projection.</summary>
public static class OfficeDocumentReadResultPdfExtensions {
    /// <summary>Projects a Reader result into PDF through <see cref="OfficeIMO.Pdf.OfficeDocumentModelPdfExtensions"/>.</summary>
    public static OfficeIMO.Pdf.PdfDocumentConversionResult ToPdfDocumentResult(
        this OfficeDocumentReadResult source,
        OfficeIMO.Pdf.PdfProjectionOptions? options = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(source);
#else
        if (source == null) throw new ArgumentNullException(nameof(source));
#endif
        CoreModel model = ToNeutralModel(source);
        return OfficeIMO.Pdf.OfficeDocumentModelPdfExtensions.ToPdfDocumentResult(model, options);
    }

    private static CoreModel ToNeutralModel(OfficeDocumentReadResult source) {
        var context = new NeutralMappingContext();
        CoreBlock[] blocks = source.Blocks.Select(context.MapBlock).ToArray();
        CoreTable[] tables = source.Tables.Select(context.MapTable).ToArray();
        CoreAsset[] assets = source.Assets.Select(context.MapAsset).ToArray();
        CoreLink[] links = source.Links.Select(context.MapLink).ToArray();
        CoreForm[] forms = source.Forms.Select(context.MapForm).ToArray();

        return new CoreModel {
            Format = MapFormat(source.Kind),
            Source = new CoreSource {
            Path = source.Source.Path,
            SourceId = source.Source.SourceId,
            SourceHash = source.Source.SourceHash,
            LastWriteUtc = source.Source.LastWriteUtc,
            LengthBytes = source.Source.LengthBytes,
            Title = source.Source.Title,
            Author = source.Source.Author,
            Subject = source.Source.Subject,
            Keywords = source.Source.Keywords
            },
            CapabilitiesUsed = source.CapabilitiesUsed,
            Markdown = source.Markdown,
            Html = source.Html,
            Metadata = source.Metadata.Select(MapMetadata).ToArray(),
            Pages = source.Pages.Select(page => MapPage(page, context)).ToArray(),
            Blocks = blocks,
            Tables = tables,
            Assets = assets,
            Links = links,
            Forms = forms,
            Visuals = source.Visuals.Select(MapVisual).ToArray(),
            Diagnostics = source.Diagnostics.Select(MapDiagnostic).ToArray()
        };
    }

    private static CorePage MapPage(ReaderPage page, NeutralMappingContext context) => new CorePage {
        Number = page.Number,
        Name = page.Name,
        Width = page.Width,
        Height = page.Height,
        RotationDegrees = page.RotationDegrees,
        Location = MapLocation(page.Location),
        Blocks = page.Blocks.Select(context.MapBlock).ToArray(),
        Tables = page.Tables.Select(context.MapTable).ToArray(),
        Assets = page.Assets.Select(context.MapAsset).ToArray(),
        Links = page.Links.Select(context.MapLink).ToArray(),
        Forms = page.Forms.Select(context.MapForm).ToArray()
    };

    private static CoreBlock MapBlock(ReaderBlock block) => new CoreBlock {
        Id = block.Id,
        Kind = block.Kind,
        Text = block.Text,
        Level = block.Level,
        Marker = block.Marker,
        Location = MapLocation(block.Location),
        Region = MapRegion(block.Region)
    };

    private static CoreTable MapTable(ReaderTable table) => new CoreTable {
        Title = table.Title,
        Kind = table.Kind,
        Summary = table.Summary,
        PayloadHash = table.PayloadHash,
        Location = table.Location == null ? null : MapLocation(table.Location),
        Columns = table.Columns,
        Rows = table.Rows,
        TotalRowCount = table.TotalRowCount,
        Truncated = table.Truncated
    };

    private static CoreAsset MapAsset(ReaderAsset asset) => new CoreAsset {
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

    private static CoreLink MapLink(ReaderLink link) => new CoreLink {
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

    private static CoreForm MapForm(ReaderForm form) => new CoreForm {
        Id = form.Id,
        Name = form.Name,
        Kind = form.Kind,
        Value = form.Value,
        IsReadOnly = form.IsReadOnly,
        IsRequired = form.IsRequired,
        Location = MapLocation(form.Location),
        Region = MapRegion(form.Region)
    };

    private static OfficeIMO.OfficeDocumentModelVisual MapVisual(ReaderVisual visual) => new OfficeIMO.OfficeDocumentModelVisual {
        Kind = visual.Kind,
        Language = visual.Language,
        Content = visual.Content,
        PayloadHash = visual.PayloadHash,
        SourceName = visual.SourceName,
        MediaType = visual.MimeType,
        Width = visual.Width,
        Height = visual.Height,
        X = visual.X,
        Y = visual.Y,
        PlacedWidth = visual.PlacedWidth,
        PlacedHeight = visual.PlacedHeight,
        PlacementCount = visual.PlacementCount,
        HasGeometry = visual.HasGeometry,
        IsAxisAligned = visual.IsAxisAligned,
        Location = visual.Location == null ? null : MapLocation(visual.Location)
    };

    private static CoreMetadata MapMetadata(ReaderMetadata metadata) => new CoreMetadata {
        Id = metadata.Id,
        Category = metadata.Category,
        Name = metadata.Name,
        Value = metadata.Value,
        ValueType = metadata.ValueType,
        SourceObjectId = metadata.SourceObjectId,
        Location = metadata.Location == null ? null : MapLocation(metadata.Location),
        Attributes = metadata.Attributes
    };

    private static CoreDiagnostic MapDiagnostic(ReaderDiagnostic diagnostic) => new CoreDiagnostic {
        Severity = diagnostic.Severity switch {
            ReaderDiagnosticSeverity.Information => OfficeIMO.OfficeDocumentModelDiagnosticSeverity.Information,
            ReaderDiagnosticSeverity.Error => OfficeIMO.OfficeDocumentModelDiagnosticSeverity.Error,
            _ => OfficeIMO.OfficeDocumentModelDiagnosticSeverity.Warning
        },
        Category = diagnostic.Category switch {
            ReaderDiagnosticCategory.Detection => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Detection,
            ReaderDiagnosticCategory.Input => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Input,
            ReaderDiagnosticCategory.Parsing => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Parsing,
            ReaderDiagnosticCategory.Content => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Content,
            ReaderDiagnosticCategory.Security => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Security,
            ReaderDiagnosticCategory.Limit => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Limit,
            ReaderDiagnosticCategory.Adapter => OfficeIMO.OfficeDocumentModelDiagnosticCategory.Adapter,
            _ => OfficeIMO.OfficeDocumentModelDiagnosticCategory.General
        },
        Code = diagnostic.Code,
        Message = diagnostic.Message,
        Source = diagnostic.Source,
        IsRecoverable = diagnostic.IsRecoverable,
        Location = diagnostic.Location == null ? null : MapLocation(diagnostic.Location),
        Attributes = diagnostic.Attributes
    };

    private static CoreLocation MapLocation(ReaderLocation location) => new CoreLocation {
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

    private static CoreRegion? MapRegion(ReaderRegion? region) => region == null
        ? null
        : new CoreRegion { X = region.X, Y = region.Y, Width = region.Width, Height = region.Height };

    private static OfficeDocumentFormat MapFormat(ReaderInputKind kind) => kind switch {
        ReaderInputKind.Word => OfficeDocumentFormat.Word,
        ReaderInputKind.Excel => OfficeDocumentFormat.Excel,
        ReaderInputKind.PowerPoint => OfficeDocumentFormat.PowerPoint,
        ReaderInputKind.Markdown => OfficeDocumentFormat.Markdown,
        ReaderInputKind.Text => OfficeDocumentFormat.Text,
        ReaderInputKind.Pdf => OfficeDocumentFormat.Pdf,
        ReaderInputKind.Csv => OfficeDocumentFormat.Csv,
        ReaderInputKind.Json => OfficeDocumentFormat.Json,
        ReaderInputKind.Xml => OfficeDocumentFormat.Xml,
        ReaderInputKind.Html => OfficeDocumentFormat.Html,
        ReaderInputKind.Zip => OfficeDocumentFormat.Zip,
        ReaderInputKind.Epub => OfficeDocumentFormat.Epub,
        ReaderInputKind.Visio => OfficeDocumentFormat.Visio,
        ReaderInputKind.Yaml => OfficeDocumentFormat.Yaml,
        ReaderInputKind.Rtf => OfficeDocumentFormat.Rtf,
        ReaderInputKind.OpenDocument => OfficeDocumentFormat.OpenDocument,
        ReaderInputKind.AsciiDoc => OfficeDocumentFormat.AsciiDoc,
        ReaderInputKind.Latex => OfficeDocumentFormat.Latex,
        ReaderInputKind.Email => OfficeDocumentFormat.Email,
        ReaderInputKind.OneNote => OfficeDocumentFormat.OneNote,
        ReaderInputKind.Calendar => OfficeDocumentFormat.Calendar,
        ReaderInputKind.VCard => OfficeDocumentFormat.VCard,
        _ => OfficeDocumentFormat.Unknown
    };

    private sealed class NeutralMappingContext {
        private readonly Dictionary<ReaderBlock, CoreBlock> _blocks = new(ReferenceIdentityComparer<ReaderBlock>.Instance);
        private readonly Dictionary<ReaderTable, CoreTable> _tables = new(ReferenceIdentityComparer<ReaderTable>.Instance);
        private readonly Dictionary<ReaderAsset, CoreAsset> _assets = new(ReferenceIdentityComparer<ReaderAsset>.Instance);
        private readonly Dictionary<ReaderLink, CoreLink> _links = new(ReferenceIdentityComparer<ReaderLink>.Instance);
        private readonly Dictionary<ReaderForm, CoreForm> _forms = new(ReferenceIdentityComparer<ReaderForm>.Instance);

        internal CoreBlock MapBlock(ReaderBlock source) =>
            GetOrAdd(_blocks, source, OfficeDocumentReadResultPdfExtensions.MapBlock);

        internal CoreTable MapTable(ReaderTable source) =>
            GetOrAdd(_tables, source, OfficeDocumentReadResultPdfExtensions.MapTable);

        internal CoreAsset MapAsset(ReaderAsset source) =>
            GetOrAdd(_assets, source, OfficeDocumentReadResultPdfExtensions.MapAsset);

        internal CoreLink MapLink(ReaderLink source) =>
            GetOrAdd(_links, source, OfficeDocumentReadResultPdfExtensions.MapLink);

        internal CoreForm MapForm(ReaderForm source) =>
            GetOrAdd(_forms, source, OfficeDocumentReadResultPdfExtensions.MapForm);

        private static TTarget GetOrAdd<TSource, TTarget>(
            IDictionary<TSource, TTarget> map,
            TSource source,
            Func<TSource, TTarget> factory) where TSource : class {
            if (map.TryGetValue(source, out TTarget? mapped)) return mapped;
            mapped = factory(source);
            map.Add(source, mapped);
            return mapped;
        }
    }

    private sealed class ReferenceIdentityComparer<T> : IEqualityComparer<T> where T : class {
        internal static ReferenceIdentityComparer<T> Instance { get; } = new ReferenceIdentityComparer<T>();

        public bool Equals(T? left, T? right) => ReferenceEquals(left, right);

        public int GetHashCode(T value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
