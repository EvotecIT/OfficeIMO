using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Builds a body-document graph from owned page and text data, without seed archives.</summary>
internal sealed partial class IWorkPagesArchiveBuilder : IWorkNativeCreationGraph {
    private const ulong DocumentId = 1, MetadataId = 2, BodyId = 3, ThemeId = 4;
    private const ulong SettingsId = 6, SectionId = 7, TemplateId = 8, ColumnId = 9, PresetId = 11;
    private const ulong FrameStyleId = 12, ImageStyleId = 13, HeaderParagraphId = 14, HeaderCharacterId = 15;
    private const ulong DefaultBodyParagraphId = 22;
    private readonly List<ulong> _paragraphVariations = new();
    private bool _customParagraphLayout;

    internal IWorkPagesArchiveBuilder(int maximumBytes, CancellationToken cancellationToken, byte[] modelHash)
        : base(maximumBytes, cancellationToken, modelHash, IWorkDocumentKind.Pages, 5, 10) { }

    internal string DocumentIdentifier => IWorkCreationIdentity.Create(_modelHash, "document", IWorkDocumentKind.Pages).Text;

    internal (byte[] Document, byte[] Metadata) Build(IWorkPagesDocument document, IWorkPagesBodyBlock[] blocks) {
        bool numbered = blocks.OfType<IWorkPagesParagraphData>().Any(paragraph => paragraph.NumberedList != null);
        _customParagraphLayout = numbered || blocks.OfType<IWorkPagesParagraphData>().Any(paragraph => paragraph.Layout.HasCustomProperties);
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

        bool customLayout = _customParagraphLayout;
        if (customLayout) Add(DefaultBodyParagraphId, 2022, P().Message(1, PagesStyle(DefaultBodyParagraphId, "OfficeIMO Body", "paragraph-style-default"))
            .Message(11, Character(new IWorkTextRunData(string.Empty, fontSizePoints: 12)))
            .Message(12, ParagraphProperties(IWorkPagesParagraphLayout.Default)).UInt(10, 19));
        var paragraphStyles = P(); var characterStyles = P(); var paragraphData = P();
        var listStyles = P(); var listStarts = P();
        var attachments = P(); var bodyText = new StringBuilder();
        int offset = 0; ulong firstParagraph = 0, firstCharacter = 0;
        foreach (IWorkPagesBodyBlock block in blocks) {
            _cancellationToken.ThrowIfCancellationRequested();
            IWorkPagesParagraphData paragraph;
            if (block is not IWorkPagesParagraphData) {
                ulong infoId = block is IWorkPagesTable table ? AddTable(BodyId, table.Data)
                    : AddImage(BodyId, ((IWorkPagesImage)block).Data, ((IWorkPagesImage)block).Geometry, ImageStyleId);
                ulong attachmentId = Allocate();
                Add(attachmentId, 2003, P().Reference(1, infoId).UInt(2, 0).Float(3, 0).UInt(4, 0).Float(5, 0));
                attachments.Message(1, P().UInt(1, (uint)offset).Reference(2, attachmentId));
                paragraph = new IWorkPagesParagraphData(new[] { new IWorkTextRunData("\ufffc", fontSizePoints: 12) });
            } else paragraph = (IWorkPagesParagraphData)block;
            ulong selectedList = paragraph.NumberedList is { } numberedList ? AddNumberedListStyle(numberedList) : ListId;
            int listStart = paragraph.NumberedList != null ? EncodeListStart(paragraph) : 0;
            if (numbered) {
                listStyles.Message(1, P().UInt(1, (uint)offset).Reference(2, selectedList));
                listStarts.Message(1, P().UInt(1, (uint)offset).UInt(2, (uint)listStart).UInt(3, 0));
            }
            bodyText.Append(paragraph.Text).Append('\n');
            ulong paragraphId = Allocate();
            if (firstParagraph == 0) firstParagraph = paragraphId;
            // The qualified custom-layout profile uses native paragraph variations and a named default parent.
            IWorkProtoWriter paragraphStyle;
            if (customLayout) {
                _paragraphVariations.Add(paragraphId);
                paragraphStyle = P().Reference(3, DefaultBodyParagraphId).Bool(4, true).Message(5, P().UInt(1, StylesheetId));
            } else paragraphStyle = Style(paragraphId, "OfficeIMO Paragraph");
            // Numbered markers inherit paragraph character defaults. Keep those defaults
            // independent of first-run emphasis; explicit character ranges carry item styling.
            IWorkTextRunData paragraphDefaults = paragraph.NumberedList is { } owner
                ? new IWorkTextRunData(string.Empty, owner.FontName, owner.FontSizePoints)
                : paragraph.Runs[0];
            Add(paragraphId, 2022, P().Message(1, paragraphStyle).Message(11, Character(paragraphDefaults))
                .Message(12, ParagraphProperties(paragraph.Layout, selectedList)).UInt(10, customLayout ? 19u : 10u));
            paragraphStyles.Message(1, P().UInt(1, (uint)offset).Reference(2, paragraphId));
            paragraphData.Message(1, P().UInt(1, (uint)offset).UInt(2, (uint)Math.Max(0, paragraph.ListLevel)).UInt(3, listStart > 0 ? 1u : 0u));
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
        if (customLayout) {
            // Pages imports explicit layout only when its paragraph attribute table has the null end marker.
            // Character attributes retain their existing termination to preserve mixed-run resets.
            paragraphStyles.Message(1, P().UInt(1, (uint)offset));
        }
        if (numbered) {
            // The final LF projects an empty paragraph; it must not acquire a list marker.
            listStyles.Message(1, P().UInt(1, (uint)offset).Reference(2, ListId));
            listStarts.Message(1, P().UInt(1, (uint)offset).UInt(2, 0).UInt(3, 0));
            paragraphData.Message(1, P().UInt(1, (uint)offset).UInt(2, 0).UInt(3, 0));
        }
        var body = P().UInt(1, 0).Reference(2, StylesheetId)
            .String(3, bodyText.ToString()).Bool(10, true)
            .Message(5, paragraphStyles).Message(6, paragraphData).Message(7, numbered ? listStyles : Attribute(ListId))
            .Message(8, characterStyles).Message(12, Attribute(ColumnId)).Message(14, numbered ? listStarts : paragraphData)
            .Message(17, Attribute(SectionId)).Message(24, P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0)))
            .Message(28, P().Message(1, P().UInt(1, 0)));
        if (blocks.Any(block => block is not IWorkPagesParagraphData)) body.Message(9, attachments);
        Add(BodyId, 2001, body);
        ulong themeParagraph = customLayout ? DefaultBodyParagraphId : firstParagraph;
        Add(PresetId, 2050, P().String(1, "officeimo-body").Reference(2, themeParagraph).Reference(3, ListId));
        var shape = P().Message(1, Style(FrameStyleId, "OfficeIMO Text Frame")).UInt(10, 3)
            .Message(11, P().Message(1, P()).Float(3, 1).Message(2, P().Message(6, P().UInt(1, 2))));
        Add(FrameStyleId, 2025, P().Message(1, shape).UInt(10, 7).Message(11, P().Bool(1, false).UInt(2, 0)
            .Message(4, P().Message(1, P().UInt(1, 1))).Message(6, ZeroPadding()).UInt(7, 0)
            .Bool(11, false).Reference(10, themeParagraph)));
        Add(ImageStyleId, 3016, P().Message(1, Style(ImageStyleId, "OfficeIMO Image")).UInt(10, 2)
            .Message(11, P().Float(2, 1).Message(1, P().Message(6, P().UInt(1, 2)))));
        var theme = P().String(3, "OfficeIMO").Reference(4, StylesheetId)
            .Message(100, P().Reference(4, FrameStyleId).Reference(5, FrameStyleId).Reference(6, FrameStyleId)
                .Reference(7, ImageStyleId).Reference(9, FrameStyleId))
            .Message(110, P().Reference(1, ListId).Reference(2, PresetId).Reference(6, firstCharacter).Reference(7, themeParagraph));
        AddTableThemePresets(theme);
        IWorkNativeThemeWriter.AddColorPresets(theme, _maximumBytes, _cancellationToken);
        Add(ThemeId, 10001, P().Message(1, theme));
        var stylesheet = P().Bool(4, false);
        foreach (var style in _styles)
            stylesheet.Reference(1, style.Id).Message(2, P().String(1, style.Identifier).Reference(2, style.Id));
        if (customLayout) {
            var children = P().Reference(1, DefaultBodyParagraphId);
            foreach (ulong id in _paragraphVariations) { stylesheet.Reference(1, id); children.Reference(2, id); }
            stylesheet.Message(5, children);
        }
        Add(StylesheetId, 401, stylesheet);
        var component = P().UInt(1, DocumentId).String(2, "Document");
        Version(component, 4, 2, 4, 0); Version(component, 5, 11, 2, 4);
        AddImageComponentReferences(component);
        var metadata = P().UInt(1, _nextId - 1).Message(3, component).UInt(9, 2);
        AddImageMetadata(metadata);
        Version(metadata, 5, 2, 4, 0); Version(metadata, 6, 11, 2, 4); Version(metadata, 7, 14, 4, 1);
        return (IWorkSnappy.EncodeIwa(IWorkNativeArchiveWriter.Encode(_records, _maximumBytes, _cancellationToken), _maximumBytes, _cancellationToken),
            IWorkSnappy.EncodeIwa(IWorkNativeArchiveWriter.Encode(new[] { (MetadataId, 11006u, metadata) }, _maximumBytes, _cancellationToken),
                _maximumBytes, _cancellationToken));
    }

    private IWorkProtoWriter ParagraphProperties(IWorkPagesParagraphLayout layout, ulong? listId = null) {
        var tabs = P();
        foreach (IWorkTabStop stop in layout.TabStops) {
            _cancellationToken.ThrowIfCancellationRequested();
            tabs.Message(1, P().Float(1, (float)stop.PositionPoints).UInt(2, (uint)stop.Alignment).String(3, ""));
        }
        var properties = P().UInt(1, 0);
        if (_customParagraphLayout) properties.String(3, ".");
        return properties.Float(4, 36).Float(7, 0).Float(11, 0)
            .Message(13, P().UInt(1, 0).Float(2, layout.LineSpacingMultiplier)).Float(19, 0)
            .Float(20, layout.SpaceAfterPoints).Float(21, layout.SpaceBeforePoints)
            .Message(25, tabs).Reference(40, listId ?? ListId);
    }

    protected override IWorkProtoWriter Style(ulong id, string name) => PagesStyle(id, name);

    private IWorkProtoWriter PagesStyle(ulong id, string name, string? identifier = null) {
        identifier ??= "officeimo-style-" + id.ToString(CultureInfo.InvariantCulture);
        _styles.Add((id, identifier));
        var style = P().String(1, name).String(2, identifier);
        // A style's stylesheet back-reference is weak in Apple's object-reference envelope.
        return _customParagraphLayout ? style.Message(5, P().UInt(1, StylesheetId)) : style.Reference(5, StylesheetId);
    }
    private IWorkProtoWriter Character(IWorkTextRunData run) => P().Bool(1, run.Bold).Bool(2, run.Italic).Float(3, run.FontSizePoints)
        .String(5, run.FontName).Message(7, Color(run.Color)).Bool(11, run.Underline).Bool(12, run.Strikethrough)
        .Message(46, P().Message(1, Color(run.Color)));
    private IWorkProtoWriter ZeroPadding() => P().Float(1, 0).Float(2, 0).Float(3, 0).Float(4, 0);
    private IWorkProtoWriter Attribute(ulong reference) => P().Message(1, P().UInt(1, 0).Reference(2, reference));
    private IWorkProtoWriter EmptyParagraphData() => P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0));
}
