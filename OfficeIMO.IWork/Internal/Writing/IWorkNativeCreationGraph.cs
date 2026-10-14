using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Owns deterministic object allocation and shared style, image and table records for native creation graphs.</summary>
internal abstract partial class IWorkNativeCreationGraph {
    protected readonly int _maximumBytes;
    protected readonly CancellationToken _cancellationToken;
    protected readonly byte[] _modelHash;
    protected readonly IWorkDocumentKind _kind;
    protected readonly ulong StylesheetId, ListId;
    protected readonly List<(ulong Id, uint Type, IWorkProtoWriter Message)> _records = new();
    protected readonly List<(ulong Id, string Identifier)> _styles = new();
    protected ulong _nextId = 32;

    protected IWorkNativeCreationGraph(int maximumBytes, CancellationToken cancellationToken,
        byte[] modelHash, IWorkDocumentKind kind, ulong stylesheetId, ulong listId) {
        _maximumBytes = maximumBytes; _cancellationToken = cancellationToken; _modelHash = modelHash;
        _kind = kind; StylesheetId = stylesheetId; ListId = listId;
    }

    protected IWorkProtoWriter P() => new(_maximumBytes, _cancellationToken);
    protected ulong Allocate() {
        // Native iWork reserves these metadata identifiers while unarchiving.
        if (_nextId is 71 or 72) _nextId = 73;
        return _nextId++;
    }
    protected void Add(ulong id, uint type, IWorkProtoWriter message) => _records.Add((id, type, message));

    protected void AddBodyAttachmentWrap(IWorkProtoWriter drawable) {
        // Absent exterior wrap makes Pages wrap surrounding body text around attachments.
        if (_kind == IWorkDocumentKind.Pages)
            drawable.Message(3, P().UInt(1, 0).UInt(2, 2).UInt(3, 0).Float(4, 12).Float(5, 0.5f).Bool(6, false));
    }

    protected static void Version(IWorkProtoWriter message, int field, params uint[] values) {
        foreach (uint value in values) message.UInt(field, value);
    }

    protected IWorkProtoWriter Size(double width, double height) => P().Float(1, (float)width).Float(2, (float)height);
    protected IWorkProtoWriter Point(double x, double y) => P().Float(1, (float)x).Float(2, (float)y);
    protected IWorkProtoWriter Color(IWorkColor color) => P().UInt(1, 1).Float(3, color.Red / 255f)
        .Float(4, color.Green / 255f).Float(5, color.Blue / 255f).Float(6, 1).UInt(12, 1);

    protected virtual IWorkProtoWriter Style(ulong id, string name) {
        string identifier = "officeimo-style-" + id.ToString(System.Globalization.CultureInfo.InvariantCulture);
        _styles.Add((id, identifier));
        return P().String(1, name).String(2, identifier).Reference(5, StylesheetId);
    }

    protected void AddTextStyles(ulong characterId, ulong paragraphId, ulong frameId, string fontName,
        float fontSize, IWorkColor color, IWorkTextRunData? firstRun = null) {
        // Modern Keynote renders tsd_fill; font_color remains its legacy color mirror.
        var character = P().Float(3, fontSize).String(5, fontName).Message(7, Color(color))
            .Message(46, P().Message(1, Color(color)));
        AddCharacterStyle(characterId, firstRun == null ? character : CharacterProperties(firstRun));
        var paragraph = P().UInt(1, 0).Reference(40, ListId).Message(13, P()).Message(25, P()).Float(4, 36);
        Add(paragraphId, 2022, P().Message(1, Style(paragraphId, "OfficeIMO Paragraph"))
            .Message(11, character).Message(12, paragraph).UInt(10, 8));
        var shapeProperties = P().Message(1, P()).Float(3, 1).Message(2, P().Message(6, P().UInt(1, 2)));
        var shape = P().Message(1, Style(frameId, "OfficeIMO Text Frame")).Message(11, shapeProperties).UInt(10, 3);
        var textProperties = P().Bool(1, false).UInt(2, 0).Message(4, P().Message(1, P().UInt(1, 1)))
            .Message(6, P().Float(1, 0).Float(2, 0).Float(3, 0).Float(4, 0)).UInt(7, 0).Bool(11, false).Reference(10, paragraphId);
        Add(frameId, 2025, P().Message(1, shape).Message(11, textProperties).UInt(10, 7));
    }

    protected IWorkProtoWriter CharacterProperties(IWorkTextRunData run) =>
        P().Bool(1, run.Bold).Bool(2, run.Italic).Float(3, run.FontSizePoints).String(5, run.FontName)
            .Message(7, Color(run.Color)).Bool(11, run.Underline).Bool(12, run.Strikethrough)
            .Message(46, P().Message(1, Color(run.Color)));

    protected void AddCharacterStyle(ulong id, IWorkProtoWriter character) =>
        Add(id, 2021, P().Message(1, Style(id, "OfficeIMO Text")).Message(11, character).UInt(10, 3));

}
