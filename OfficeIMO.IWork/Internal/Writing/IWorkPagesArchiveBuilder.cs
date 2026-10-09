using System.Globalization;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Builds a body-document graph from owned page and text data, without seed archives.</summary>
internal sealed class IWorkPagesArchiveBuilder {
    private const ulong DocumentId = 1, MetadataId = 2, BodyId = 3, ThemeId = 4, StylesheetId = 5;
    private const ulong SettingsId = 6, SectionId = 7, TemplateId = 8, ColumnId = 9, ListId = 10, PresetId = 11;
    private const ulong FrameStyleId = 12, ImageStyleId = 13, HeaderParagraphId = 14, HeaderCharacterId = 15;
    private readonly int _maximumBytes;
    private readonly CancellationToken _cancellationToken;
    private readonly byte[] _modelHash;
    private readonly List<(ulong Id, uint Type, IWorkProtoWriter Message)> _records = new();
    private readonly List<(ulong Id, string Identifier)> _styles = new();
    private ulong _nextId = 32;

    internal IWorkPagesArchiveBuilder(int maximumBytes, CancellationToken cancellationToken, byte[] modelHash) {
        _maximumBytes = maximumBytes; _cancellationToken = cancellationToken; _modelHash = modelHash;
    }

    internal string DocumentIdentifier => IWorkCreationIdentity.Create(_modelHash, "document", IWorkDocumentKind.Pages).Text;
    private IWorkProtoWriter P() => new(_maximumBytes, _cancellationToken);
    private void Add(ulong id, uint type, IWorkProtoWriter message) => _records.Add((id, type, message));
    private ulong Allocate() {
        if (_nextId is 71 or 72) _nextId = 73;
        return _nextId++;
    }

    internal (byte[] Document, byte[] Metadata) Build(IWorkPagesDocument document, IWorkPagesParagraphData[] paragraphs) {
        Add(DocumentId, 10000, P().Reference(2, StylesheetId).Reference(4, BodyId).Reference(6, ThemeId).Reference(7, SettingsId)
            .Message(15, P().Message(1, P().String(4, "en_US")).String(3, "en"))
            .Float(30, (float)document.PageSize.WidthPoints).Float(31, (float)document.PageSize.HeightPoints)
            .Float(32, document.LeftMarginPoints).Float(33, document.RightMarginPoints)
            .Float(34, document.TopMarginPoints).Float(35, document.BottomMarginPoints)
            .Float(36, 0).Float(37, 0).Float(38, 1).Bool(39, false).Bool(40, false)
            .UInt(42, document.PageSize.WidthPoints > document.PageSize.HeightPoints ? 1u : 0u));
        Add(SettingsId, 10012, P().Bool(1, true).Bool(2, false).Bool(3, false).Bool(8, true).Bool(9, false)
            .Bool(10, true).String(20, ".").String(21, "en").String(23, "en_US"));
        Add(SectionId, 10011, P().Bool(17, true).Bool(18, false).Bool(19, false).UInt(20, 0).UInt(21, 0).UInt(22, 1)
            .Reference(23, TemplateId).Reference(24, TemplateId).Reference(25, TemplateId).String(26, "OfficeIMO").Bool(28, false));
        // Pages writes all three header/footer positions even when document settings hide them.
        // Missing storages can open initially but cause native save to emit unresolved storage identity zero.
        var template = P();
        for (uint index = 0; index < 6; index++) {
            ulong storageId = 16 + index;
            Add(storageId, 2001, P().UInt(1, 1).Reference(2, StylesheetId).String(3, "\n").Bool(10, true)
                .Message(5, Attribute(HeaderParagraphId)).Message(7, Attribute(ListId)).Message(8, Attribute(HeaderCharacterId))
                .Message(12, Attribute(ColumnId)).Message(28, P().Message(1, P().UInt(1, 0)))
                .Message(6, EmptyParagraphData()).Message(14, EmptyParagraphData()).Message(24, EmptyParagraphData()));
            template.Reference(index < 3 ? 1 : 2, storageId);
        }
        Add(TemplateId, 10143, template);
        Add(ColumnId, 2024, P().Message(1, Style(ColumnId, "OfficeIMO Body Layout")).UInt(10, 12)
            .Message(11, P().Bool(1, false).Bool(2, false).Bool(3, false).Float(4, 0).UInt(5, 0)
                .Bool(6, false).Message(7, P().Message(1, P().UInt(1, 1).Float(2, 0)))
                .Bool(8, false).Message(9, ZeroPadding()).Bool(10, false).Message(11, ZeroPadding()).UInt(12, 0)));
        var list = P().Message(1, Style(ListId, "None")).UInt(10, 3);
        for (int index = 0; index < 9; index++) list.UInt(11, 0).Float(12, 0).Float(13, 0);
        Add(ListId, 2023, list);
        var headerCharacter = Character(new IWorkTextRunData(string.Empty, fontSizePoints: 12));
        Add(HeaderCharacterId, 2021, P().Message(1, Style(HeaderCharacterId, "OfficeIMO Header Text"))
            .Message(11, headerCharacter).UInt(10, 8));
        Add(HeaderParagraphId, 2022, P().Message(1, Style(HeaderParagraphId, "OfficeIMO Header Paragraph"))
            .Message(11, headerCharacter).Message(12, P().UInt(1, 0).Float(4, 36).Message(13, P()).Message(25, P())
                .Reference(40, ListId)).UInt(10, 5));

        var paragraphStyles = P(); var characterStyles = P(); var paragraphData = P();
        int offset = 0; ulong firstParagraph = 0, firstCharacter = 0;
        foreach (IWorkPagesParagraphData paragraph in paragraphs) {
            _cancellationToken.ThrowIfCancellationRequested();
            ulong paragraphId = Allocate();
            if (firstParagraph == 0) firstParagraph = paragraphId;
            Add(paragraphId, 2022, P().Message(1, Style(paragraphId, "OfficeIMO Paragraph"))
                .Message(11, Character(paragraph.Runs[0])).Message(12, P().UInt(1, 0).Float(4, 36)
                    .Float(7, 0).Float(11, 0).Message(13, P().UInt(1, 0).Float(2, 1)).Float(19, 0)
                    .Float(20, 0).Float(21, 0).Message(25, P()).Reference(40, ListId)).UInt(10, 10));
            paragraphStyles.Message(1, P().UInt(1, (uint)offset).Reference(2, paragraphId));
            paragraphData.Message(1, P().UInt(1, (uint)offset).UInt(2, 0).UInt(3, 0));
            foreach (IWorkTextRunData run in paragraph.Runs) {
                _cancellationToken.ThrowIfCancellationRequested();
                ulong characterId = Allocate();
                if (firstCharacter == 0) firstCharacter = characterId;
                Add(characterId, 2021, P().Message(1, Style(characterId, "OfficeIMO Text")).Message(11, Character(run)).UInt(10, 8));
                characterStyles.Message(1, P().UInt(1, (uint)offset).Reference(2, characterId));
                offset += run.Text.Length;
            }
            offset++;
        }
        Add(BodyId, 2001, P().UInt(1, 0).Reference(2, StylesheetId)
            .String(3, string.Concat(paragraphs.Select(paragraph => paragraph.Text + "\n"))).Bool(10, true)
            .Message(5, paragraphStyles).Message(6, paragraphData).Message(7, Attribute(ListId))
            .Message(8, characterStyles).Message(12, Attribute(ColumnId)).Message(14, paragraphData)
            .Message(17, Attribute(SectionId)).Message(24, P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0)))
            .Message(28, P().Message(1, P().UInt(1, 0))));
        Add(PresetId, 2050, P().String(1, "officeimo-body").Reference(2, firstParagraph).Reference(3, ListId));
        var shape = P().Message(1, Style(FrameStyleId, "OfficeIMO Text Frame")).UInt(10, 3)
            .Message(11, P().Message(1, P()).Float(3, 1).Message(2, P().Message(6, P().UInt(1, 2))));
        Add(FrameStyleId, 2025, P().Message(1, shape).UInt(10, 7).Message(11, P().Bool(1, false).UInt(2, 0)
            .Message(4, P().Message(1, P().UInt(1, 1))).Message(6, ZeroPadding()).UInt(7, 0)
            .Bool(11, false).Reference(10, firstParagraph)));
        Add(ImageStyleId, 3016, P().Message(1, Style(ImageStyleId, "OfficeIMO Image")).UInt(10, 2)
            .Message(11, P().Float(2, 1).Message(1, P().Message(6, P().UInt(1, 2)))));
        var theme = P().String(3, "OfficeIMO").Reference(4, StylesheetId)
            .Message(100, P().Reference(4, FrameStyleId).Reference(5, FrameStyleId).Reference(6, FrameStyleId)
                .Reference(7, ImageStyleId).Reference(9, FrameStyleId))
            .Message(110, P().Reference(1, ListId).Reference(2, PresetId).Reference(6, firstCharacter).Reference(7, firstParagraph));
        IWorkNativeThemeWriter.AddColorPresets(theme, _maximumBytes, _cancellationToken);
        Add(ThemeId, 10001, P().Message(1, theme));
        var stylesheet = P().Bool(4, false);
        foreach (var style in _styles)
            stylesheet.Reference(1, style.Id).Message(2, P().String(1, style.Identifier).Reference(2, style.Id));
        Add(StylesheetId, 401, stylesheet);
        var component = P().UInt(1, DocumentId).String(2, "Document");
        Version(component, 4, 2, 4, 0); Version(component, 5, 11, 2, 4);
        var metadata = P().UInt(1, _nextId - 1).Message(3, component).UInt(9, 2);
        Version(metadata, 5, 2, 4, 0); Version(metadata, 6, 11, 2, 4); Version(metadata, 7, 14, 4, 1);
        return (IWorkSnappy.EncodeIwa(IWorkNativeArchiveWriter.Encode(_records, _maximumBytes, _cancellationToken), _maximumBytes, _cancellationToken),
            IWorkSnappy.EncodeIwa(IWorkNativeArchiveWriter.Encode(new[] { (MetadataId, 11006u, metadata) }, _maximumBytes, _cancellationToken),
                _maximumBytes, _cancellationToken));
    }

    private IWorkProtoWriter Style(ulong id, string name) {
        string identifier = "officeimo-style-" + id.ToString(CultureInfo.InvariantCulture);
        _styles.Add((id, identifier));
        return P().String(1, name).String(2, identifier).Reference(5, StylesheetId);
    }
    private IWorkProtoWriter Character(IWorkTextRunData run) => P().Bool(1, run.Bold).Bool(2, run.Italic).Float(3, run.FontSizePoints)
        .String(5, run.FontName).Message(7, Color(run.Color)).Bool(11, run.Underline).Bool(12, run.Strikethrough)
        .Message(46, P().Message(1, Color(run.Color)));
    private IWorkProtoWriter Color(IWorkColor color) => P().UInt(1, 1).Float(3, color.Red / 255f)
        .Float(4, color.Green / 255f).Float(5, color.Blue / 255f).Float(6, 1).UInt(12, 1);
    private IWorkProtoWriter ZeroPadding() => P().Float(1, 0).Float(2, 0).Float(3, 0).Float(4, 0);
    private IWorkProtoWriter Attribute(ulong reference) => P().Message(1, P().UInt(1, 0).Reference(2, reference));
    private IWorkProtoWriter EmptyParagraphData() => P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0));
    private static void Version(IWorkProtoWriter message, int field, params uint[] values) {
        foreach (uint value in values) message.UInt(field, value);
    }
}
