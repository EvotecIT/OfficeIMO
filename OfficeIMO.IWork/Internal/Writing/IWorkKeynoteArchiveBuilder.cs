using System.Globalization;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Builds an owned Keynote 14.4.1 object graph from the creation model. No seed package or opaque archive payload is used.</summary>
internal sealed partial class IWorkKeynoteArchiveBuilder {
    private const ulong DocumentId = 1, MetadataId = 2, ShowId = 3, ThemeId = 4, StylesheetId = 5;
    private const ulong DefaultCharacterId = 6, DefaultParagraphId = 7, DefaultFrameId = 8, BlankBackgroundId = 9;
    private const ulong ListId = 14, TextPresetId = 15, BlankSlideId = 30, BlankNodeId = 31;
    private readonly IWorkKeynoteWriteOptions _limits;
    private readonly CancellationToken _cancellationToken;
    private readonly List<(ulong Id, uint Type, IWorkProtoWriter Message)> _records = new();
    private readonly List<(ulong Id, string Identifier)> _styles = new();
    private readonly byte[] _modelHash;
    private ulong _nextId = 32;

    internal IWorkKeynoteArchiveBuilder(IWorkKeynoteWriteOptions limits, CancellationToken cancellationToken, byte[] modelHash) {
        _limits = limits;
        _cancellationToken = cancellationToken;
        _modelHash = modelHash;
    }

    internal string DocumentIdentifier => IWorkKeynoteIdentity.Create(_modelHash, "document").Text;
    private IWorkProtoWriter P() => new(_limits.MaximumPackageBytes, _cancellationToken);
    private ulong Allocate() => _nextId++;
    private void Add(ulong id, uint type, IWorkProtoWriter message) => _records.Add((id, type, message));

    internal (byte[] Document, byte[] Metadata) Build(IWorkCanvasSize size,
        IReadOnlyList<(IWorkKeynoteSlideBuilder Slide, IWorkKeynoteTextBox[] Boxes)> slides) {
        _cancellationToken.ThrowIfCancellationRequested();
        Add(DocumentId, 1, P().Message(3, P().Message(1, P().String(4, "en_US")).String(3, "en")).Reference(2, ShowId));
        AddDefaultStyles();
        AddTheme();
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
            foreach (IWorkKeynoteTextBox box in input.Boxes) {
                _cancellationToken.ThrowIfCancellationRequested();
                ulong characterId = Allocate(), paragraphId = Allocate(), frameStyleId = Allocate();
                ulong frameId = Allocate(), storageId = Allocate();
                AddTextStyles(characterId, paragraphId, frameStyleId, box.FontName, box.FontSizePoints, box.Color);
                Add(frameId, 2011, Frame(slideId, frameStyleId, storageId, box.Geometry));
                Add(storageId, 2001, Storage(paragraphId, characterId, box.Text));
                slide.Reference(7, frameId).Reference(42, frameId);
            }
            Add(slideId, 5, slide);
            Add(nodeId, 4, Node(slideId, "slide-" + index.ToString(CultureInfo.InvariantCulture)));
            tree.Reference(2, nodeId);
        }
        Add(ShowId, 2, P().Reference(2, ThemeId).Reference(5, StylesheetId).Message(4, Size(size.WidthPoints, size.HeightPoints))
            .Message(3, tree).Bool(6, false).Bool(8, false).UInt(9, 0).Bool(18, false));

        var stylesheet = P().Bool(4, false);
        foreach (var style in _styles) {
            stylesheet.Reference(1, style.Id).Message(2, P().String(1, style.Identifier).Reference(2, style.Id));
        }
        Add(StylesheetId, 401, stylesheet);
        var component = P().UInt(1, DocumentId).String(2, "Document");
        Version(component, 4, 2, 4, 0);
        Version(component, 5, 11, 2, 4);
        var metadata = P().UInt(1, _nextId - 1).Message(3, component).UInt(9, 2);
        Version(metadata, 5, 2, 4, 0);
        Version(metadata, 6, 11, 2, 4);
        Version(metadata, 7, 14, 4, 1);
        byte[] document = Archive(_records.OrderBy(record => record.Id));
        byte[] metadataBytes = Archive(new[] { (MetadataId, 11006u, metadata) });
        return (IWorkSnappy.EncodeIwa(document, _limits.MaximumPackageBytes, _cancellationToken),
            IWorkSnappy.EncodeIwa(metadataBytes, _limits.MaximumPackageBytes, _cancellationToken));
    }

    private byte[] Archive(IEnumerable<(ulong Id, uint Type, IWorkProtoWriter Message)> records) {
        var output = P();
        foreach (var record in records) {
            _cancellationToken.ThrowIfCancellationRequested();
            byte[] payload = record.Message.ToArray();
            var info = P().UInt(1, record.Type).UInt(3, (ulong)payload.Length);
            Version(info, 2, 1, 0, 5);
            foreach (ulong reference in record.Message.References) info.UInt(5, reference);
            byte[] header = P().UInt(1, record.Id).Message(2, info).ToArray();
            output.AppendLength(header.Length);
            output.Append(header);
            output.Append(payload);
        }
        return output.ToArray();
    }

    private static void Version(IWorkProtoWriter message, int field, params uint[] values) {
        foreach (uint value in values) message.UInt(field, value);
    }

    private IWorkProtoWriter Size(double width, double height) => P().Float(1, (float)width).Float(2, (float)height);
    private IWorkProtoWriter Point(double x, double y) => P().Float(1, (float)x).Float(2, (float)y);
    private IWorkProtoWriter Color(IWorkColor color) => P().UInt(1, 1).Float(3, color.Red / 255f)
        .Float(4, color.Green / 255f).Float(5, color.Blue / 255f).Float(6, 1).UInt(12, 1);

    private IWorkProtoWriter Transition() => P().Message(2, P().Message(8,
        P().String(1, "Transition").String(2, "none").Double(3, 1).Double(5, 0).Bool(6, false).UInt(11, 1).Bool(16, false)));

    private IWorkProtoWriter Node(ulong slideId, string role) {
        var templateId = IWorkKeynoteIdentity.Create(_modelHash, "blank-layout");
        return P().Reference(2, slideId).Bool(4, false).Bool(6, false).Bool(7, false).Bool(8, false)
            .Bool(18, false).Bool(28, false).UInt(21, 1).Bool(14, true).UInt(26, uint.MaxValue).UInt(27, uint.MaxValue)
            .String(11, IWorkKeynoteIdentity.Create(_modelHash, role).Text)
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

    private IWorkProtoWriter Storage(ulong paragraphId, ulong characterId, string text) {
        var storage = P().UInt(1, 3).Reference(2, StylesheetId).String(3, text + "\n").Bool(10, true);
        storage.Message(5, Attribute(paragraphId)).Message(8, Attribute(characterId)).Message(7, Attribute(ListId));
        // These tables are necessary for native paragraph editing and export, even with no lists or bidi overrides.
        foreach (int field in new[] { 6, 14, 24 }) storage.Message(field, P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0)));
        storage.Message(28, P().Message(1, P().UInt(1, 0)));
        return storage;
    }

    private IWorkProtoWriter Attribute(ulong reference) => P().Message(1, P().UInt(1, 0).Reference(2, reference));
}
