using System.Globalization;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Builds an owned Keynote 14.4.1 object graph from the creation model. No seed package or opaque archive payload is used.</summary>
internal sealed partial class IWorkKeynoteArchiveBuilder : IWorkNativeCreationGraph {
    private const ulong DocumentId = 1, MetadataId = 2, ShowId = 3, ThemeId = 4;
    private const ulong DefaultCharacterId = 6, DefaultParagraphId = 7, DefaultFrameId = 8, BlankBackgroundId = 9;
    private const ulong TextPresetId = 15, BlankSlideId = 30, BlankNodeId = 31;
    private readonly IWorkKeynoteWriteOptions _limits;
    internal IWorkKeynoteArchiveBuilder(IWorkKeynoteWriteOptions limits, CancellationToken cancellationToken, byte[] modelHash)
        : base(limits.MaximumPackageBytes, cancellationToken, modelHash, IWorkDocumentKind.Keynote, 5, 14) {
        _limits = limits;
    }

    internal string DocumentIdentifier => IWorkCreationIdentity.Create(_modelHash, "document").Text;
    internal (byte[] Document, byte[] Metadata) Build(IWorkCanvasSize size,
        IReadOnlyList<(IWorkKeynoteSlideBuilder Slide, IWorkKeynoteCreationDrawable[] Drawables)> slides) {
        _cancellationToken.ThrowIfCancellationRequested();
        Add(DocumentId, 1, P().Message(3, P().Message(1, P().String(4, "en_US")).String(3, "en")).Reference(2, ShowId));
        AddDefaultStyles();
        if (slides.Any(slide => slide.Drawables.Any(drawable => drawable is IWorkKeynoteImage))) AddImageStyle();
        Add(BlankSlideId, 5, P().Reference(1, BlankBackgroundId).Message(4, Transition())
            .String(10, "Blank").Bool(19, true).Bool(41, false));
        Add(BlankNodeId, 4, Node(BlankSlideId, "blank-node"));

        var tree = P();
        for (int index = 0; index < slides.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            var input = slides[index];
            ulong slideId = Allocate(), nodeId = Allocate(), backgroundId = Allocate();
            AddBackground(backgroundId, input.Slide.BackgroundColor);
            // Field 10 (name) identifies a KNTemplateSlide to Apple. Ordinary show slides must omit it.
            var slide = P().Reference(1, backgroundId).Message(4, Transition()).Bool(19, true)
                .Bool(41, false).Reference(17, BlankSlideId);
            foreach (IWorkKeynoteCreationDrawable drawable in input.Drawables) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (drawable is IWorkKeynoteImage image) {
                    ulong imageId = AddImage(slideId, image);
                    slide.Reference(7, imageId).Reference(42, imageId);
                    continue;
                }
                if (drawable is IWorkKeynoteTable table) {
                    ulong tableId = AddTable(slideId, table.Data);
                    slide.Reference(7, tableId).Reference(42, tableId);
                    continue;
                }
                var box = (IWorkKeynoteTextBox)drawable;
                ulong characterId = Allocate(), paragraphId = Allocate(), frameStyleId = Allocate();
                ulong frameId = Allocate(), storageId = Allocate();
                bool rich = box.Runs.Count > 1 || box.Runs[0].HasDecoration;
                AddTextStyles(characterId, paragraphId, frameStyleId, box.FontName, box.FontSizePoints, box.Color,
                    rich ? box.Runs[0] : null);
                Add(frameId, 2011, Frame(slideId, frameStyleId, storageId, box.Geometry));
                var characters = new List<(int Offset, ulong Style)> { (0, characterId) };
                int offset = box.Runs[0].Text.Length;
                foreach (IWorkTextRunData run in box.Runs.Skip(1)) {
                    ulong runStyle = Allocate();
                    AddCharacterStyle(runStyle, CharacterProperties(run));
                    characters.Add((offset, runStyle));
                    offset = checked(offset + run.Text.Length);
                }
                Add(storageId, 2001, Storage(paragraphId, characters, box.Text));
                slide.Reference(7, frameId).Reference(42, frameId);
            }
            Add(slideId, 5, slide);
            Add(nodeId, 4, Node(slideId, "slide-" + index.ToString(CultureInfo.InvariantCulture)));
            tree.Reference(2, nodeId);
        }
        Add(ShowId, 2, P().Reference(2, ThemeId).Reference(5, StylesheetId).Message(4, Size(size.WidthPoints, size.HeightPoints))
            .Message(3, tree).Bool(6, false).Bool(8, false).UInt(9, 0).Bool(18, false));
        // Table style presets are registered while building their typed drawables.
        AddTheme();

        var stylesheet = P().Bool(4, false);
        foreach (var style in _styles) {
            stylesheet.Reference(1, style.Id).Message(2, P().String(1, style.Identifier).Reference(2, style.Id));
        }
        Add(StylesheetId, 401, stylesheet);
        var component = P().UInt(1, DocumentId).String(2, "Document");
        Version(component, 4, 2, 4, 0);
        Version(component, 5, 11, 2, 4);
        AddImageComponentReferences(component);
        var metadata = P().UInt(1, _nextId - 1).Message(3, component).UInt(9, 2);
        AddImageMetadata(metadata);
        Version(metadata, 5, 2, 4, 0);
        Version(metadata, 6, 11, 2, 4);
        Version(metadata, 7, 14, 4, 1);
        byte[] document = Archive(_records.OrderBy(record => record.Id));
        byte[] metadataBytes = Archive(new[] { (MetadataId, 11006u, metadata) });
        return (IWorkSnappy.EncodeIwa(document, _limits.MaximumPackageBytes, _cancellationToken),
            IWorkSnappy.EncodeIwa(metadataBytes, _limits.MaximumPackageBytes, _cancellationToken));
    }

    private byte[] Archive(IEnumerable<(ulong Id, uint Type, IWorkProtoWriter Message)> records) =>
        IWorkNativeArchiveWriter.Encode(records, _limits.MaximumPackageBytes, _cancellationToken);

    private IWorkProtoWriter Transition() => P().Message(2, P().Message(8,
        P().String(1, "Transition").String(2, "none").Double(3, 1).Double(5, 0).Bool(6, false).UInt(11, 1).Bool(16, false)));

    private IWorkProtoWriter Node(ulong slideId, string role) {
        var templateId = IWorkCreationIdentity.Create(_modelHash, "blank-layout");
        return P().Reference(2, slideId).Bool(4, false).Bool(6, false).Bool(7, false).Bool(8, false)
            .Bool(18, false).Bool(28, false).UInt(21, 1).Bool(14, true).UInt(26, uint.MaxValue).UInt(27, uint.MaxValue)
            .String(11, IWorkCreationIdentity.Create(_modelHash, role).Text)
            .Message(29, P().UInt(1, templateId.Lower).UInt(2, templateId.Upper));
    }

    private IWorkProtoWriter Frame(ulong slideId, ulong styleId, ulong storageId, IWorkGeometry geometry) {
        var drawable = P().Message(1, P().Message(1, Point(geometry.LeftPoints, geometry.TopPoints))
            .Message(2, Size(geometry.WidthPoints, geometry.HeightPoints)).UInt(3, 3).Float(4, 0)).Reference(2, slideId);
        var path = P();
        foreach (var point in new[] { (1u, 0d, 0d), (2u, geometry.WidthPoints, 0d),
                     (2u, geometry.WidthPoints, geometry.HeightPoints), (2u, 0d, geometry.HeightPoints),
                     (5u, 0d, 0d), (1u, 0d, 0d) }) {
            var element = P().UInt(1, point.Item1);
            if (point.Item1 != 5) element.Message(2, Point(point.Item2, point.Item3));
            path.Message(1, element);
        }
        var bezier = P().Message(2, Size(geometry.WidthPoints, geometry.HeightPoints)).Message(3, path);
        var shape = P().Message(1, drawable).Reference(2, styleId).Message(3, P().Message(5, bezier));
        return P().Message(1, shape).Reference(2, storageId).Reference(4, storageId).Bool(6, true);
    }

    private IWorkProtoWriter Storage(ulong paragraphId, IReadOnlyList<(int Offset, ulong Style)> characters, string text) {
        var storage = P().UInt(1, 3).Reference(2, StylesheetId).String(3, text + "\n").Bool(10, true);
        var characterTable = P();
        foreach (var character in characters)
            characterTable.Message(1, P().UInt(1, (ulong)character.Offset).Reference(2, character.Style));
        storage.Message(5, Attribute(paragraphId)).Message(8, characterTable).Message(7, Attribute(ListId));
        // These tables are necessary for native paragraph editing and export, even with no lists or bidi overrides.
        foreach (int field in new[] { 6, 14, 24 }) storage.Message(field, P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0)));
        storage.Message(28, P().Message(1, P().UInt(1, 0)));
        return storage;
    }

    private IWorkProtoWriter Attribute(ulong reference) => P().Message(1, P().UInt(1, 0).Reference(2, reference));
}
