using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Resolves retained native thresholds and languages for scene/flow copies without changing public metadata or save requests.</summary>
internal sealed class PublisherQuillKerningProjection {
    private readonly PublisherPublicationState _publication;
    private readonly PublisherImportedResidue? _imported;
    private readonly PublisherParseContext _context;
    private readonly Dictionary<uint, IReadOnlyList<OfficeRichTextParagraph>> _stories = new();
    private PublisherQuillSourceLayout? _layout;
    private bool _layoutRead;

    internal PublisherQuillKerningProjection(PublisherPublicationState publication, PublisherImportedResidue? imported,
        PublisherParseContext context) { _publication = publication; _imported = imported; _context = context; }

    internal IReadOnlyList<OfficeRichTextParagraph> Story(PublisherTextStory story) {
        if (_stories.TryGetValue(story.Id, out var result)) return result;
        result = story.Paragraphs;
        if (_imported?.NativeSource != null && _imported.Baseline != null
            && _imported.Baseline.Text.Stories.TryGetValue(story.Id, out PublisherTextStory? before)
            && story.Paragraphs.Any(paragraph => paragraph.Runs.Count != 0)) {
            if (!TryProject(before, story, out var effective))
                _context.Add("PUB_TEXT_KERNING_SOURCE_UNASSESSED",
                    "Effective native kerning could not be resolved through a qualified retained source definition. Explicit run settings remain active; unspecified runs retain generic renderer defaults.",
                    OfficeConversionLossKind.Unassessed, "Quill/story/" + story.Id);
            else result = effective;
        }
        result = PublisherHyphenationProjection.Project(story, result, _context);
        _stories.Add(story.Id, result);
        return result;
    }

    internal IReadOnlyList<OfficeRichTextParagraph> Cell(uint storyId, int index) {
        PublisherTableStory table = _publication.Text.TableStories[storyId];
        var story = _publication.Text.Stories[storyId];
        IReadOnlyList<OfficeRichTextParagraph> effective = Story(story);
        if (ReferenceEquals(effective, story.Paragraphs)) return table.Segments[index].Paragraphs;
        int start = table.Segments.Take(index).Sum(segment => segment.Paragraphs.Count);
        return effective.Skip(start).Take(table.Segments[index].Paragraphs.Count).ToArray();
    }

    private bool TryProject(PublisherTextStory before, PublisherTextStory after,
        out IReadOnlyList<OfficeRichTextParagraph> paragraphs) {
        paragraphs = after.Paragraphs;
        if (!_layoutRead) {
            _layoutRead = true;
            _layout = PublisherQuillSourceLayout.TryRead(_imported!.NativeSource!.QuillBytes, _context);
        }
        if (_layout == null) return false;
        PublisherQuillStoryStyles? binding = _imported!.Styles.Stories.FirstOrDefault(value => value.StoryId == before.Id);
        if (binding == null || binding.NullCharacters != 0 || binding.NativeLength != before.Text.Length) return false;
        var pc = new List<PublisherQuillFormattingRange>();
        var pp = new List<PublisherQuillFormattingRange>();
        var data = _layout.Data;
        bool unchanged = PublisherTextComparison.StoryEqual(before, after);
        _imported.Baseline!.Text.TableStories.TryGetValue(before.Id, out PublisherTableStory? priorTable);
        _publication.Text.TableStories.TryGetValue(before.Id, out PublisherTableStory? currentTable);
        if ((priorTable == null) != (currentTable == null)) return false;
        if (priorTable != null && (priorTable.Segments.Count != currentTable!.Segments.Count
            || !priorTable.Matches(before) || !currentTable.Matches(after))) return false;
        unchanged &= priorTable == null || priorTable.Segments.Zip(currentTable!.Segments,
            PublisherTextComparison.StoryEqual).All(value => value);
        if (unchanged) {
            foreach (PublisherQuillStyleUse use in binding.Uses) {
                _context.Token.ThrowIfCancellationRequested();
                var range = new PublisherQuillFormattingRange(checked((use.NativeStart + use.NativeLength - binding.NativeStart) * 2),
                    data.Bytes.Skip(use.DirectDefinition.Offset).Take(use.DirectDefinition.Length).ToArray());
                if (use.Tag == "FDPC") pc.Add(range);
                else if (use.Tag == "FDPP") pp.Add(range);
            }
        } else {
            if (priorTable == null) {
                if (!PublisherQuillParagraphWriter.CanRewriteStory(before, after, out _)) return false;
            } else if (!priorTable.Segments.Zip(currentTable!.Segments, (a, b) => PublisherTextComparison.StoryEqual(a, b)
                || PublisherQuillParagraphWriter.CanRewriteStory(a, b, out _)).All(value => value)) return false;
            var resources = new PublisherQuillWriteResources(_layout, _context);
            var typing = new PublisherQuillStyleReader(data, _context, _layout.Chunks, _imported.Contents.Palette, _layout.Text).Read();
            bool cellTopology = priorTable != null && priorTable.Segments.Zip(currentTable!.Segments,
                (a, b) => a.Paragraphs.Count != b.Paragraphs.Count).Any(value => value);
            bool planned = cellTopology
                ? PublisherQuillTableTextWriter.TryPlanParagraphs(data, _context, binding, priorTable!, currentTable!, 0,
                    _context.Token, out pc, out pp, resources, typing)
                : PublisherQuillParagraphWriter.TryPlan(data, _context, binding, before, after, 0, _context.Token,
                    out pc, out pp, resources, typing);
            if (!planned) return false;
        }
        return TryApply(after, pc, pp, out paragraphs);
    }

    private bool TryApply(PublisherTextStory story, IReadOnlyList<PublisherQuillFormattingRange> pc,
        IReadOnlyList<PublisherQuillFormattingRange> pp, out IReadOnlyList<OfficeRichTextParagraph> paragraphs) {
        paragraphs = story.Paragraphs;
        var thresholds = new List<(int End, double? Value, IReadOnlyList<PublisherBlock> Fields)>();
        int last = 0;
        foreach (var range in pc) {
            IReadOnlyList<PublisherBlock> fields = Fields(range.Style);
            if (range.AbsoluteEnd <= last || (range.AbsoluteEnd & 1) != 0
                || !PublisherQuillKerning.TryRead(fields, out double? threshold)) return false;
            thresholds.Add((range.AbsoluteEnd / 2, threshold, fields));
            last = range.AbsoluteEnd;
        }
        if (thresholds.Count == 0 || last != story.Text.Length * 2) return false;
        var selectors = new List<(int End, double Default, uint? Language)>();
        last = 0;
        foreach (var range in pp) {
            if (range.AbsoluteEnd <= last || (range.AbsoluteEnd & 1) != 0
                || !Default(Fields(range.Style), out double value, out uint? language)) return false;
            selectors.Add((range.AbsoluteEnd / 2, value, language));
            last = range.AbsoluteEnd;
        }
        if (selectors.Count == 0 || last != story.Text.Length * 2) return false;
        var result = new List<OfficeRichTextParagraph>();
        int position = 0, character = 0, paragraphIndex = 0, runCount = 0;
        foreach (OfficeRichTextParagraph paragraph in story.Paragraphs) {
            _context.Token.ThrowIfCancellationRequested();
            while (paragraphIndex < selectors.Count && selectors[paragraphIndex].End <= position) paragraphIndex++;
            if (paragraphIndex == selectors.Count) return false;
            double inherited = selectors[paragraphIndex].Default;
            var runs = new List<OfficeRichTextRun>();
            foreach (OfficeRichTextRun run in paragraph.Runs) {
                int offset = 0;
                do {
                    while (character < thresholds.Count && thresholds[character].End <= position) character++;
                    if (character == thresholds.Count) return false;
                    int length = Math.Min(run.Text.Length - offset, thresholds[character].End - position);
                    if (++runCount > 4096) return false;
                    runs.Add(Copy(run, run.Text.Substring(offset, length), run.AutomaticPairKerningThreshold
                        ?? thresholds[character].Value ?? inherited,
                        run.LanguageId ?? Language(thresholds[character].Fields, selectors[paragraphIndex].Language)));
                    offset += length;
                    position += length;
                } while (offset < run.Text.Length);
            }
            OfficeTextParagraphLabel? label = paragraph.Label;
            if (label != null) label = label.WithRun(label.Run.WithAutomaticPairKerningThreshold(
                label.Run.AutomaticPairKerningThreshold ?? inherited));
            if (label != null && ++runCount > 4096 || runCount + result.Count > 4096) return false;
            var copy = label == null ? new OfficeRichTextParagraph(runs, paragraph.Alignment, paragraph.LineHeight,
                paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor)
                : new OfficeRichTextParagraph(runs, label, paragraph.Alignment, paragraph.LineHeight,
                    paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor);
            copy.SuppressAutomaticHyphenation = paragraph.SuppressAutomaticHyphenation;
            copy = copy.WithTabStops(paragraph.TabStops);
            copy.UsesPublisherFontMetrics = paragraph.UsesPublisherFontMetrics;
            copy.UsesPublisherKerningContext = true;
            copy.ContinuesInNextRegion = paragraph.ContinuesInNextRegion;
            result.Add(copy);
            if (position < story.Text.Length) position++; // Native paragraph CR, absent after an unterminated final paragraph.
        }
        if (position != story.Text.Length) return false;
        paragraphs = result;
        return true;
    }

    private IReadOnlyList<PublisherBlock> Fields(byte[] bytes) =>
        new PublisherBlockReader(new PublisherBinaryData(bytes, "Quill/kerning"), _context).Chunk(0, bytes.Length);

    private bool Default(IReadOnlyList<PublisherBlock> fields, out double threshold, out uint? language) {
        threshold = PublisherQuillKerning.DefaultThreshold;
        language = null;
        var selectors = fields.Where(field => field.Id == 0x19).ToArray();
        var catalog = _imported!.Styles.Definitions;
        if (selectors.Length > 1 || selectors.Any(field => field.Type != 0x22) || catalog.Count == 0) return false;
        uint selected = selectors.Length == 0 ? catalog[0].Paragraph.DefaultStyleIndex ?? 0 : selectors[0].Value;
        if (selected >= catalog.Count) return false;
        PublisherQuillSourceSpan span = catalog[(int)selected].CharacterDefinition;
        IReadOnlyList<PublisherBlock> definition = new PublisherBlockReader(_layout!.Data, _context).Chunk(span.Offset, span.Offset + span.Length);
        if (!PublisherQuillKerning.TryRead(definition, out double? value)) return false;
        threshold = value ?? PublisherQuillKerning.DefaultThreshold;
        PublisherQuillSourceSpan global = catalog[0].CharacterDefinition;
        language = Language(definition, Language(new PublisherBlockReader(_layout.Data, _context)
            .Chunk(global.Offset, global.Offset + global.Length), null));
        return true;
    }

    private static uint? Language(IReadOnlyList<PublisherBlock> fields, uint? inherited) {
        PublisherBlock[] languages = fields.Where(field => field.Id == 0x12).ToArray();
        if (languages.Length == 0) return inherited;
        // An opaque direct setting must not acquire a dictionary through a fallback.
        return languages.Length == 1 && languages[0].Type == 0x22 && languages[0].Value <= ushort.MaxValue
            ? languages[0].Value : null;
    }

    private static OfficeRichTextRun Copy(OfficeRichTextRun run, string text, double threshold, uint? language) =>
        new(text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily,
            run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) {
            AutomaticPairKerningThreshold = threshold, LanguageId = language, LinkUri = run.LinkUri, LinkDestinationName = run.LinkDestinationName
        };
}
