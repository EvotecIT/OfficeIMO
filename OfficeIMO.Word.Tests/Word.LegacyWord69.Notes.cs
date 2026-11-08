using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(101)]
    [InlineData(104)]
    public void LegacyWord69_HyphenControlsRetainBodyAndNoteTextAnchorsAndEditableDocx(int version) {
        byte[] input = CreateWord69NotesFixture((ushort)version,
            footnoteText: "\u0002Foot\u001Ekeep\u001Fsoft\r",
            endnoteText: "\u0002End\u001Fsoft\u001Ekeep\r",
            mutate: main => { main[0x400 + 6] = 0x1E; main[0x400 + 16] = 0x1F; });
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        AssertHyphens(document, edited: false);
        WordParagraph run = Assert.Single(document.FootNotes).Paragraphs!.First(paragraph => paragraph.Text.Contains("Foot"));
        run.SetText(run.Text.Replace("Foot", "EditedFoot"));
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        AssertHyphens(reopened, edited: true);
        Assert.Empty(reopened.ValidateDocument());

        static void AssertHyphens(WordDocument doc, bool edited) {
            WordParagraphInfo body = Assert.Single(doc.Inspect().Sections.SelectMany(section => section.Elements).OfType<WordParagraphInfo>());
            Assert.Equal("Before\u2011[F] between\u00AD[E] after", string.Concat(body.Runs.Select(run =>
                run.Footnote != null ? "[F]" : run.Endnote != null ? "[E]" : run.Text)));
            WordFootNote footnote = Assert.Single(doc.FootNotes);
            WordEndNote endnote = Assert.Single(doc.EndNotes);
            Assert.Equal(footnote.ReferenceId, Assert.Single(body.Runs, run => run.Footnote != null).Footnote!.ReferenceId);
            Assert.Equal(endnote.ReferenceId, Assert.Single(body.Runs, run => run.Endnote != null).Endnote!.ReferenceId);
            Assert.Equal(new[] { (edited ? "EditedFoot" : "Foot") + "\u2011keep\u00ADsoft" }, footnote.TextParagraphs);
            Assert.Equal(new[] { "End\u00ADsoft\u2011keep" }, endnote.TextParagraphs);
        }
    }

    [Theory]
    [InlineData(101)]
    [InlineData(104)]
    public void LegacyWord69_AutomaticNotesRetainAnchorsFormattingEncodingAndEditableDocx(int version) {
        using WordDocument document = WordDocument.Load(new MemoryStream(CreateWord69NotesFixture((ushort)version)));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        Assert.Single(document.FootNotes);
        Assert.Single(document.EndNotes);
        AssertWord69NotePackage(document, new[] { "Fřtnote", "", "Second para" }, "Eøndnote");
        document.FootNotes[0].Paragraphs!.Single(paragraph => paragraph.Text == "Second para").SetText("Edited footnote");
        document.EndNotes[0].Paragraphs!.Single(paragraph => paragraph.Text == "Eøndnote").SetText("Edited endnote");
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        AssertWord69NotePackage(reopened, new[] { "Fřtnote", "", "Edited footnote" }, "Edited endnote");
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(reopened.OpenXmlDocument));
    }

    [Fact]
    public void LegacyWord69_AutomaticNotesPreserveEmptyAndUnmarkedTextStories() {
        using WordDocument document = WordDocument.Load(new MemoryStream(CreateWord69NotesFixture(104,
            footnoteText: "\r", endnoteText: "21 literal note text\rSecond paragraph\r\r")));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        var main = reopened.OpenXmlDocument.MainDocumentPart!;
        W.Footnote footnote = Assert.Single(main.FootnotesPart!.Footnotes.Elements<W.Footnote>(), note => note.Id?.Value > 0);
        W.Endnote endnote = Assert.Single(main.EndnotesPart!.Endnotes.Elements<W.Endnote>(), note => note.Id?.Value > 0);
        Assert.Equal("", Assert.Single(footnote.Elements<W.Paragraph>()).InnerText);
        Assert.Equal(new[] { "21 literal note text", "Second paragraph", "" }, endnote.Elements<W.Paragraph>().Select(paragraph => paragraph.InnerText));
        Assert.Single(main.Document.Descendants<W.FootnoteReference>());
        Assert.Single(main.Document.Descendants<W.EndnoteReference>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(reopened.OpenXmlDocument));
    }

    [Theory]
    [InlineData("duplicate-anchor")]
    [InlineData("out-of-body-reference")]
    [InlineData("out-of-story-text")]
    [InlineData("short-endnote-header")]
    [InlineData("truncated-reference-record")]
    [InlineData("out-of-stream-table")]
    public void LegacyWord69_MalformedNoteTablesFailBeforeProjection(string invariant) {
        byte[] source = CreateWord69NotesFixture(104, mutate: main => {
            switch (invariant) {
                case "duplicate-anchor": WriteInt(main, 0xA30, BitConverter.ToInt32(main, 0xA00)); break;
                case "out-of-body-reference": WriteInt(main, 0xA00, BitConverter.ToInt32(main, 0x34)); break;
                case "out-of-story-text": WriteInt(main, 0xA14, BitConverter.ToInt32(main, 0x38) + 1); break;
                case "short-endnote-header": WriteInt(main, 0x18, 0x192); break;
                case "truncated-reference-record": WriteInt(main, 0x6C, 9); break;
                case "out-of-stream-table": WriteInt(main, 0x68, main.Length - 2); break;
            }
        });
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(new MemoryStream(source)));
        Assert.Contains("Word 6/95", error.Message);
    }

    [Theory]
    [InlineData("custom-mark", "doc-word69-footnotes")]
    [InlineData("structural-anchor-control", "doc-word69-footnotes")]
    [InlineData("unqualified-placement", "doc-word69-note-settings")]
    [InlineData("empty-field-cache", "doc-word69-fields-unsupported")]
    [InlineData("unprojected-separator", "doc-word69-running-stories")]
    [InlineData("nested-note-marker", "doc-word69-note-content-unqualified")]
    [InlineData("note-table-controls", "doc-word69-note-content-unqualified")]
    public void LegacyWord69_UnqualifiedNoteFeaturesRemainExplicitStrictLoss(string invariant, string code) {
        string? footnote = invariant switch {
            "empty-field-cache" => "\u0002\u0013 PAGE \u0014\u0015\r",
            "nested-note-marker" => "\u0002Known \u0002 orphan\r",
            "note-table-controls" => "\aCell\r",
            _ => null
        };
        byte[] source = CreateWord69NotesFixture(101, footnoteText: footnote, mutate: main => {
            if (invariant == "custom-mark") {
                WriteWord(main, 0xA08, 0);
                main[0x400 + BitConverter.ToInt32(main, 0xA00)] = (byte)'*';
            }
            if (invariant == "structural-anchor-control") main[0x400 + BitConverter.ToInt32(main, 0xA00)] = (byte)'\t';
            if (invariant == "unqualified-placement") main[0x2C0] = 0;
            if (invariant == "unprojected-separator") {
                main[0x2C1] = 1;
                WriteInt(main, 0xB4, 12);
                WriteInt(main, 0xA68, 1);
            }
        });
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.Contains(document.LegacyDocUnsupportedFeatures, finding => finding.Code == code);
        using var strict = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(strict));
        Assert.Equal(0, strict.Length);
        using var allowed = new MemoryStream();
        document.Save(allowed, new WordSaveOptions { AllowLoss = true });
        allowed.Position = 0;
        using WordDocument reopened = WordDocument.Load(allowed);
        Assert.Single(reopened.EndNotes);
        Assert.Contains(reopened.EndNotes[0].Paragraphs!, paragraph => paragraph.Text == "Eøndnote");
        if (invariant == "custom-mark") {
            Assert.Empty(reopened.FootNotes);
            Assert.Contains("*", reopened.OpenXmlDocument.MainDocumentPart!.Document.Body!.InnerText);
        }
        if (invariant == "note-table-controls") {
            W.Footnote retained = Assert.Single(reopened.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes.Elements<W.Footnote>(),
                note => note.Id?.Value > 0);
            Assert.Equal(new[] { "", "Cell" }, retained.Elements<W.Paragraph>().Select(paragraph => paragraph.InnerText));
        }
        if (invariant == "nested-note-marker") {
            Assert.Contains(reopened.FootNotes[0].Paragraphs!, paragraph => paragraph.Text.Contains("Known") && paragraph.Text.Contains("orphan"));
        }
    }

    [Fact]
    public void LegacyWord69_NoteFieldPlcsRetainInertCacheAndLockInBothStories() {
        byte[] source = CreateWord69NotesFixture(104,
            footnoteText: "\u0002\u0013 FILENAME \u0014foot.doc\u0015\r",
            endnoteText: "\u0002\u0013 FILENAME \u0014end.doc\u0015\r", lockedFields: true);
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        W.SimpleField footnote = Assert.Single(reopened.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes.Descendants<W.SimpleField>());
        W.SimpleField endnote = Assert.Single(reopened.OpenXmlDocument.MainDocumentPart.EndnotesPart!.Endnotes.Descendants<W.SimpleField>());
        Assert.Equal(new[] { "foot.doc", "end.doc" }, new[] { footnote.InnerText, endnote.InnerText });
        Assert.All(new[] { footnote, endnote }, field => {
            Assert.Equal("FILENAME", field.Instruction!.Value!.Trim());
            Assert.True(field.FieldLock?.Value);
            Assert.False(field.Dirty?.Value ?? false);
        });
    }

    [Fact]
    public void LegacyWord69_EndnoteBaseIncludesUnprojectedMacroAndAnnotationStories() {
        using WordDocument document = WordDocument.Load(new MemoryStream(CreateWord69NotesFixture(104, extraStories: true)));
        Assert.Contains(document.LegacyDocUnsupportedFeatures, finding => finding.Code == "doc-word69-wordbasic");
        Assert.Contains(document.LegacyDocUnsupportedFeatures, finding => finding.Code == "doc-word69-comments");
        Assert.DoesNotContain(document.LegacyDocUnsupportedFeatures, finding => finding.Code is "doc-word69-footnotes" or "doc-word69-endnotes");
        AssertWord69NotePackage(document, new[] { "Fřtnote", "", "Second para" }, "Eøndnote");
    }

    [Fact]
    public void LegacyWord69_NotesShareTheAggregateDecodedCharacterBudget() {
        byte[] source = CreateWord69NotesFixture(101);
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        int total = BitConverter.ToInt32(main, 0x34) + BitConverter.ToInt32(main, 0x38)
            + BitConverter.ToInt32(main, 0x3C) + BitConverter.ToInt32(main, 0x48);
        using WordDocument accepted = WordDocument.Load(new MemoryStream(source),
            new WordLoadOptions { MaxDecodedCharacters = total });
        Assert.Single(accepted.FootNotes);
        Assert.Single(accepted.EndNotes);
        Assert.Throws<InvalidDataException>(() => WordDocument.Load(new MemoryStream(source),
            new WordLoadOptions { MaxDecodedCharacters = total - 1 }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LegacyDoc_LoadAsyncPropagatesCancellationAfterSourceCapture(bool word97) {
        byte[] source = word97
            ? LegacyDocTestBuilder.CreateSimpleDocWithFootnoteStory("Body", "Footnote")
            : CreateWord69NotesFixture(104);
        using var cancellation = new CancellationTokenSource();
        using var input = new CancelOnSnapshotRestoreStream(source, cancellation);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => WordDocument.LoadAsync(input, cancellationToken: cancellation.Token));
        Assert.True(input.CancelledOnRestore);
        Assert.True(input.CanRead);
        Assert.Equal(0, input.Position);
    }

    private static void AssertWord69NotePackage(WordDocument document, string[] footnoteParagraphs, string endnoteText) {
        W.Body body = document.OpenXmlDocument.MainDocumentPart!.Document.Body!;
        W.Paragraph paragraph = Assert.Single(body.Elements<W.Paragraph>());
        string anchors = string.Concat(paragraph.Descendants().Select(element => element switch {
            W.Text text => text.Text,
            W.FootnoteReference => "[F]",
            W.EndnoteReference => "[E]",
            _ => ""
        }));
        Assert.Equal("Before [F] between [E] after", anchors);
        W.Footnote footnote = Assert.Single(document.OpenXmlDocument.MainDocumentPart.FootnotesPart!.Footnotes.Elements<W.Footnote>(),
            note => note.Id?.Value > 0);
        W.Endnote endnote = Assert.Single(document.OpenXmlDocument.MainDocumentPart.EndnotesPart!.Endnotes.Elements<W.Endnote>(),
            note => note.Id?.Value > 0);
        Assert.Equal(footnote.Id!.Value, Assert.Single(paragraph.Descendants<W.FootnoteReference>()).Id!.Value);
        Assert.Equal(endnote.Id!.Value, Assert.Single(paragraph.Descendants<W.EndnoteReference>()).Id!.Value);
        Assert.Equal(footnoteParagraphs, footnote.Elements<W.Paragraph>().Select(item => item.InnerText));
        Assert.Equal(endnoteText, Assert.Single(endnote.Elements<W.Paragraph>()).InnerText);
        W.Run firstText = Assert.Single(footnote.Descendants<W.Run>(), run => run.InnerText == "Fřtnote");
        Assert.NotNull(firstText.RunProperties?.Bold);
        Assert.True(firstText.RunProperties!.Bold!.Val?.Value ?? true);
        W.Run endText = Assert.Single(endnote.Descendants<W.Run>(), run => run.InnerText == endnoteText);
        Assert.NotNull(endText.RunProperties?.Italic);
        Assert.True(endText.RunProperties!.Italic!.Val?.Value ?? true);
        W.SectionProperties properties = Assert.Single(body.Elements<W.SectionProperties>());
        Assert.Equal(1, properties.GetFirstChild<W.FootnoteProperties>()!.GetFirstChild<W.NumberingStart>()!.Val!.Value);
        Assert.Equal(W.NumberFormatValues.Decimal, properties.GetFirstChild<W.FootnoteProperties>()!.GetFirstChild<W.NumberingFormat>()!.Val!.Value);
    }

    // Synthetic supplement to the pinned old grammar, not an independent producer oracle.
    private static byte[] CreateWord69NotesFixture(
        ushort version, bool extraStories = false, string? footnoteText = null, string? endnoteText = null,
        bool lockedFields = false, Action<byte[]>? mutate = null) {
        const int textStart = 0x400;
        string body = "Before \u0002 between \u0002 after\r";
        string footnote = footnoteText ?? "\u0002F\u00F8tnote\r\rSecond para\r";
        const string header = "\r";
        string macro = extraStories ? "Macro guard\r" : "";
        string annotation = extraStories ? "Annotation guard\r" : "";
        string endnote = endnoteText ?? "\u0002E\u00F8ndnote\r";
        if (version == 104) {
            // Word95 can cache printable automatic markers and retain an unused PLC tail.
            footnote += "\r";
            endnote += "\r";
        }
        string text = body + footnote + header + macro + annotation + endnote;
        var main = new byte[0x1000];
        WriteWord(main, 0, 0xA5DC);
        WriteWord(main, 2, version);
        WriteWord(main, 6, 0x0409);
        WriteInt(main, 0x18, textStart);
        WriteInt(main, 0x1C, textStart + text.Length);
        WriteInt(main, 0x20, main.Length);
        WriteInt(main, 0x34, body.Length);
        WriteInt(main, 0x38, footnote.Length);
        WriteInt(main, 0x3C, header.Length);
        WriteInt(main, 0x40, macro.Length);
        WriteInt(main, 0x44, annotation.Length);
        WriteInt(main, 0x48, endnote.Length);
        for (int index = 0; index < text.Length; index++) main[textStart + index] = (byte)text[index];
        if (version == 104) {
            main[textStart + body.IndexOf('\u0002')] = (byte)'1';
            main[textStart + body.LastIndexOf('\u0002')] = (byte)'1';
        }
        WriteNoteTables(0x68, 0xA00, body.IndexOf('\u0002'), footnote.Length);
        WriteNoteTables(0x1D2, 0xA30, body.LastIndexOf('\u0002'), endnote.Length);
        WriteInt(main, 0xB0, 0xA60);
        WriteInt(main, 0xB4, 8);
        WriteInt(main, 0xA64, 1);
        WriteInt(main, 0x150, 0x2C0);
        WriteInt(main, 0x154, 56);
        main[0x2C0] = 0x20;
        WriteWord(main, 0x2C2, 4);
        WriteWord(main, 0x2F4, 4);
        WriteWord(main, 0x2F6, 3);
        WriteInt(main, 0xD0, 0x280);
        WriteInt(main, 0xD4, 18);
        WriteWord(main, 0x280, 18);
        for (int font = 0; font < 2; font++) {
            int start = 0x282 + font * 8;
            main[start] = 7;
            main[start + 4] = font == 0 ? (byte)0 : (byte)238;
            main[start + 6] = (byte)('A' + font);
        }
        WriteInt(main, 0xB8, 0x320);
        WriteInt(main, 0xBC, 10);
        WriteInt(main, 0x320, textStart);
        WriteInt(main, 0x324, textStart + text.Length);
        WriteWord(main, 0x328, 4);
        WriteWord(main, 0x18A, 4);
        WriteWord(main, 0x18E, 1);
        WriteInt(main, 0x800, textStart);
        WriteInt(main, 0x804, textStart + body.Length);
        WriteInt(main, 0x808, textStart + body.Length + footnote.Length);
        WriteInt(main, 0x80C, textStart + text.Length);
        WriteCharacterRecord(0, 480, new byte[] { 93, 0, 0 });
        WriteCharacterRecord(1, 488, new byte[] { 85, 1, 93, 1, 0 });
        WriteCharacterRecord(2, 496, new byte[] { 86, 1, 93, 0, 0 });
        main[0x9FF] = 3;
        if (lockedFields) {
            WriteFieldTable(footnote, 0xE8, 0xAC0);
            WriteFieldTable(endnote, 0x1E2, 0xAE0);
        }
        mutate?.Invoke(main);
        return OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });

        void WriteNoteTables(int directory, int offset, int anchor, int length) {
            WriteInt(main, directory, offset);
            WriteInt(main, directory + 4, 10);
            WriteInt(main, directory + 8, offset + 16);
            WriteInt(main, directory + 12, version == 104 ? 12 : 8);
            WriteInt(main, offset, anchor);
            WriteInt(main, offset + 4, body.Length);
            WriteWord(main, offset + 8, 1);
            WriteInt(main, offset + 16, 0);
            WriteInt(main, offset + 20, version == 104 ? length - 1 : length);
            if (version == 104) WriteInt(main, offset + 24, length + 2);
        }

        void WriteCharacterRecord(int index, int offset, byte[] data) {
            main[0x800 + 16 + index] = (byte)(offset / 2);
            main[0x800 + offset] = (byte)data.Length;
            data.CopyTo(main, 0x800 + offset + 1);
        }

        void WriteFieldTable(string story, int directory, int offset) {
            int[] records = Enumerable.Range(0, story.Length).Where(index => story[index] is '\u0013' or '\u0014' or '\u0015').ToArray();
            WriteInt(main, directory, offset);
            WriteInt(main, directory + 4, 4 + 6 * records.Length);
            int properties = offset + (records.Length + 1) * 4;
            for (int index = 0; index < records.Length; index++) {
                WriteInt(main, offset + index * 4, records[index]);
                main[properties + index * 2] = (byte)story[records[index]];
                main[properties + index * 2 + 1] = story[records[index]] == '\u0015' ? (byte)0x90 : (byte)0;
            }
            WriteInt(main, offset + records.Length * 4, int.MaxValue);
        }
    }

    private sealed class CancelOnSnapshotRestoreStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        private int _seeks;
        internal CancelOnSnapshotRestoreStream(byte[] bytes, CancellationTokenSource cancellation) : base(bytes) =>
            _cancellation = cancellation;
        internal bool CancelledOnRestore { get; private set; }
        public override long Seek(long offset, SeekOrigin origin) {
            long position = base.Seek(offset, origin);
            if (++_seeks == 2) {
                CancelledOnRestore = true;
                _cancellation.Cancel();
            }
            return position;
        }
    }
}
