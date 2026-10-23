using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footnote")]
    public void Word_Rtf_Mixed_Run_Preserves_Text_Break_Order_In_Stories(string story) {
        using WordDocument source = WordDocument.Create();
        WordParagraph target = CreateMixedRunStory(source, story);
        target._paragraph.RemoveAllChildren<Run>();
        target._paragraph.Append(new Run(new RunProperties(new Bold()),
            new Break(), new Text("FIRST"), new TabChar(), new Text("SECOND"),
            new Break { Type = BreakValues.Page }, new Text("THIRD"),
            new Break { Type = BreakValues.Column }, new Text("FOURTH"), new Break()));
        using WordDocument opened = ReopenMixedRunSource(source);
        Dictionary<string, string> before = SnapshotMixedRunSource(opened);

        RtfDocument converted = opened.ToRtfDocument();
        AssertMixedRunBreakOrder(SelectMixedRunStory(converted, story));
        string artifact = converted.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument readback = RtfDocument.Read(artifact).Document;
        AssertMixedRunBreakOrder(SelectMixedRunStory(readback, story), allowNoteBodyMarker: story == "footnote");
        AssertMixedRunSourceUnchanged(opened, before);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Word_Rtf_Mixed_Run_Preserves_Text_And_Every_Note_Reference(bool endnote) {
        using WordDocument source = WordDocument.Create();
        WordParagraph target = source.AddParagraph();
        if (endnote) {
            target.AddEndNote("NOTE-FIRST");
            target.AddEndNote("NOTE-SECOND");
        } else {
            target.AddFootNote("NOTE-FIRST");
            target.AddFootNote("NOTE-SECOND");
        }
        OpenXmlElement[] references = target._paragraph.Descendants()
            .Where(element => element is FootnoteReference or EndnoteReference)
            .Select(element => element.CloneNode(true)).ToArray();
        Assert.Equal(2, references.Length);
        target._paragraph.RemoveAllChildren<Run>();
        target._paragraph.Append(new Run(new RunProperties(new Bold()),
            new Text("BEFORE"), references[0], references[1], new Text("AFTER")));
        using WordDocument opened = ReopenMixedRunSource(source);
        Dictionary<string, string> before = SnapshotMixedRunSource(opened);

        RtfDocument converted = opened.ToRtfDocument();
        RtfParagraph paragraph = Assert.Single(MixedRunBodyBlocks(converted).OfType<RtfParagraph>());
        RtfNoteKind kind = endnote ? RtfNoteKind.Endnote : RtfNoteKind.Footnote;
        Assert.Collection(paragraph.Inlines,
            inline => {
                RtfRun run = Assert.IsType<RtfRun>(inline);
                Assert.Equal("BEFORE", run.Text);
                Assert.True(run.Bold);
                AssertMixedRunNote(run.Note, kind, "NOTE-FIRST");
            },
            inline => AssertMixedRunNote(Assert.IsType<RtfGeneratedText>(inline).Note, kind, "NOTE-SECOND"),
            inline => {
                RtfRun run = Assert.IsType<RtfRun>(inline);
                Assert.Equal("AFTER", run.Text);
                Assert.True(run.Bold);
                Assert.Null(run.Note);
            });
        string artifact = converted.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        AssertMixedRunArtifactOrder(artifact, "BEFORE", "NOTE-FIRST", "NOTE-SECOND", "AFTER");
        AssertMixedRunSourceUnchanged(opened, before);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Word_Rtf_Mixed_Run_Preserves_Image_Order_With_Text_Or_A_Note(bool noteOnly) {
        byte[] png = CreateOnePixelPng();
        using WordDocument source = WordDocument.Create();
        WordParagraph target = source.AddParagraph();
        using (var image = new MemoryStream(png, writable: false)) {
            target.AddImage(image, "mixed-run.png", 16, 16, WordImageTextWrapping.InLineWithText, "Mixed image");
        }
        WordDrawing drawing = Assert.Single(target._paragraph.Descendants<WordDrawing>());
        drawing.Remove();
        OpenXmlElement? reference = null;
        if (noteOnly) {
            target.AddFootNote("IMAGE-NOTE");
            reference = Assert.Single(target._paragraph.Descendants<FootnoteReference>()).CloneNode(true);
        }
        target._paragraph.RemoveAllChildren<Run>();
        target._paragraph.Append(noteOnly
            ? new Run(drawing, reference!)
            : new Run(new RunProperties(new Bold()), new Text("BEFORE"), drawing, new Text("AFTER")));
        using WordDocument opened = ReopenMixedRunSource(source);
        Dictionary<string, string> before = SnapshotMixedRunSource(opened);

        RtfDocument converted = opened.ToRtfDocument();
        RtfParagraph paragraph = Assert.Single(MixedRunBodyBlocks(converted).OfType<RtfParagraph>());
        if (noteOnly) {
            Assert.Collection(paragraph.Inlines,
                inline => Assert.Equal(png, Assert.IsType<RtfImage>(inline).Data),
                inline => AssertMixedRunNote(Assert.IsType<RtfGeneratedText>(inline).Note, RtfNoteKind.Footnote, "IMAGE-NOTE"));
        } else {
            Assert.Collection(paragraph.Inlines,
                inline => {
                    RtfRun run = Assert.IsType<RtfRun>(inline);
                    Assert.Equal("BEFORE", run.Text);
                    Assert.True(run.Bold);
                },
                inline => Assert.Equal(png, Assert.IsType<RtfImage>(inline).Data),
                inline => {
                    RtfRun run = Assert.IsType<RtfRun>(inline);
                    Assert.Equal("AFTER", run.Text);
                    Assert.True(run.Bold);
                });
        }
        string artifact = converted.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument readback = RtfDocument.Read(artifact).Document;
        IEnumerable<IRtfInline> readbackInlines = MixedRunBodyBlocks(readback).SelectMany(block => block is RtfParagraph content
            ? content.Inlines.AsEnumerable()
            : new[] { Assert.IsType<RtfImage>(block) }.Cast<IRtfInline>());
        AssertMixedRunImageArtifact(readbackInlines, png, noteOnly);
        AssertMixedRunArtifactOrder(artifact, noteOnly ? new[] { @"\pict", "IMAGE-NOTE" }
            : new[] { "BEFORE", @"\pict", "AFTER" });
        AssertMixedRunSourceUnchanged(opened, before);
    }

    private static WordParagraph CreateMixedRunStory(WordDocument source, string story) => story switch {
        "body" => source.AddParagraph(),
        "cell" => source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
        "header" => source.HeaderDefaultOrCreate.AddParagraph(),
        "footnote" => source.AddParagraph("Anchor").AddFootNote("Placeholder").FootNote!.Paragraphs!.Last(),
        _ => throw new ArgumentOutOfRangeException(nameof(story))
    };

    private static RtfParagraph SelectMixedRunStory(RtfDocument document, string story) => story switch {
        "body" => Assert.Single(MixedRunBodyBlocks(document).OfType<RtfParagraph>()),
        "cell" => Assert.Single(Assert.Single(Assert.Single(MixedRunBodyBlocks(document).OfType<RtfTable>()).Rows).Cells).Paragraphs.Single(),
        "header" => document.HeaderFooters.Concat(document.Sections.SelectMany(section => section.HeaderFooters))
            .Distinct().SelectMany(header => header.Paragraphs).Single(paragraph => paragraph.Runs.Any(run => run.Text.Contains("FIRST"))),
        "footnote" => Assert.Single(MixedRunBodyBlocks(document).OfType<RtfParagraph>()).Inlines
            .Select(inline => inline is RtfRun run ? run.Note : (inline as RtfGeneratedText)?.Note)
            .Where(note => note?.Kind == RtfNoteKind.Footnote).Single()!.Paragraphs
            .Single(paragraph => paragraph.Runs.Any(run => run.Text.Contains("FIRST"))),
        _ => throw new ArgumentOutOfRangeException(nameof(story))
    };

    private static IEnumerable<IRtfBlock> MixedRunBodyBlocks(RtfDocument document) => document.Sections.Count > 0
        ? document.Sections.SelectMany(section => section.Blocks) : document.Blocks;

    private static void AssertMixedRunBreakOrder(RtfParagraph paragraph, bool allowNoteBodyMarker = false) {
        IEnumerable<IRtfInline> inlines = paragraph.Inlines;
        if (allowNoteBodyMarker && inlines.FirstOrDefault() is RtfGeneratedText {
                Kind: RtfGeneratedTextKind.NoteReference, Note: null, FallbackText: null }) {
            inlines = inlines.Skip(1);
        }

        var observed = new List<string>();
        var text = new System.Text.StringBuilder();
        foreach (IRtfInline inline in inlines) {
            if (inline is RtfRun run) {
                Assert.Null(run.Note);
                if (run.Text.Length == 0) continue;
                Assert.True(run.Bold);
                text.Append(run.Text);
                continue;
            }

            if (text.Length > 0) {
                observed.Add("text:" + text);
                text.Clear();
            }
            observed.Add("break:" + Assert.IsType<RtfBreak>(inline).Kind);
        }
        if (text.Length > 0) observed.Add("text:" + text);
        Assert.Equal(new[] {
            "break:Line", "text:FIRST\tSECOND", "break:Page", "text:THIRD",
            "break:Column", "text:FOURTH", "break:Line"
        }, observed);
    }

    private static void AssertMixedRunImageArtifact(IEnumerable<IRtfInline> inlines, byte[] png, bool noteOnly) {
        var observed = new List<string>();
        var text = new System.Text.StringBuilder();
        foreach (IRtfInline inline in inlines) {
            if (inline is RtfRun run) {
                Assert.Null(run.Note);
                if (run.Text.Length == 0) continue;
                Assert.True(run.Bold);
                text.Append(run.Text);
                continue;
            }
            if (text.Length > 0) {
                observed.Add("text:" + text);
                text.Clear();
            }
            if (inline is RtfImage image) {
                Assert.Equal(png, image.Data);
                observed.Add("image");
            } else {
                RtfGeneratedText reference = Assert.IsType<RtfGeneratedText>(inline);
                Assert.Equal(RtfGeneratedTextKind.NoteReference, reference.Kind);
                Assert.Null(reference.FallbackText);
                AssertMixedRunNote(reference.Note, RtfNoteKind.Footnote, "IMAGE-NOTE");
                observed.Add("note");
            }
        }
        if (text.Length > 0) observed.Add("text:" + text);
        Assert.Equal(noteOnly ? new[] { "image", "note" }
            : new[] { "text:BEFORE", "image", "text:AFTER" }, observed);
    }

    private static void AssertMixedRunNote(RtfNote? note, RtfNoteKind kind, string text) {
        Assert.NotNull(note);
        Assert.Equal(kind, note!.Kind);
        Assert.Equal(text, note.ToPlainText());
    }

    private static void AssertMixedRunArtifactOrder(string artifact, params string[] markers) {
        int previous = -1;
        foreach (string marker in markers) {
            int current = artifact.IndexOf(marker, previous + 1, StringComparison.Ordinal);
            Assert.True(current > previous, "Missing or reordered marker: " + marker);
            previous = current;
        }
    }

    private static WordDocument ReopenMixedRunSource(WordDocument source) {
        Assert.Empty(new OpenXmlValidator().Validate(source._wordprocessingDocument));
        using var bytes = new MemoryStream(source.ToBytes(), writable: false);
        return WordDocument.Load(bytes);
    }

    private static Dictionary<string, string> SnapshotMixedRunSource(WordDocument source) {
        var parts = new Dictionary<string, string>(StringComparer.Ordinal);
        SnapshotMixedRunParts(source._wordprocessingDocument, parts);
        return parts;
    }

    private static void SnapshotMixedRunParts(OpenXmlPartContainer container, Dictionary<string, string> parts) {
        foreach (IdPartPair pair in container.Parts) {
            OpenXmlPart part = pair.OpenXmlPart;
            string path = part.Uri.ToString();
            if (parts.ContainsKey(path)) continue;
            parts.Add(path, part.RootElement?.OuterXml ?? string.Empty);
            SnapshotMixedRunParts(part, parts);
        }
    }

    private static void AssertMixedRunSourceUnchanged(WordDocument source, Dictionary<string, string> before) {
        Assert.Empty(new OpenXmlValidator().Validate(source._wordprocessingDocument));
        Dictionary<string, string> after = SnapshotMixedRunSource(source);
        Assert.Equal(before.Keys.OrderBy(path => path), after.Keys.OrderBy(path => path));
        foreach (KeyValuePair<string, string> part in before) Assert.Equal(part.Value, after[part.Key]);
    }
}
