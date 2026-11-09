using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Loads and persists complete automatic footnotes and endnotes in the old single-byte grammar.</summary>
internal sealed partial class Word69NotesWorkload : MigrationWorkload {
    private WordDocument? _document;
    private WordDocumentInfo? _inspection;
    private OfficeConversionReport? _report;

    internal Word69NotesWorkload(int units, ushort fibVersion) : base(CreateInput(units, fibVersion), units) { }

    public override void Execute(string operation) {
        using var input = new MemoryStream(Input, writable: false);
        _document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        _report = report;
        switch (operation) {
            case "LoadInspect":
                _inspection = _document.Inspect(includeImageBytes: false);
                break;
            case "ConvertSave":
                using (var output = new MemoryStream()) {
                    _document.Save(output);
                    Output = output.ToArray();
                }
                break;
            default:
                throw new ArgumentException("Unknown migration operation: " + operation, nameof(operation));
        }
    }

    public override void Validate() {
        if (_document == null || _report == null) throw new InvalidOperationException("The note workload did not load a document.");
        DiagnosticDigest = string.Join("|", _document.LegacyDocUnsupportedFeatures.Select(finding => finding.Code));
        SemanticDigest = string.Empty;
        Require(_document.SourceFormat == WordFileFormat.Doc && DiagnosticDigest.Length == 0
            && _report.Diagnostics.Count == 0 && !_report.HasLoss,
            "source format and loss-free diagnostics: " + DiagnosticDigest);
        Require(_document.FootNotes.Count == Units / 2 && _document.EndNotes.Count == Units / 2,
            "public ordinary note counts");
        ValidateInspection(_inspection ?? _document.Inspect(includeImageBytes: false));
        ValidatePackage(_document.OpenXmlDocument);
        if (Output != null) {
            using (var source = new MemoryStream(Output, writable: false))
            using (var package = WordprocessingDocument.Open(source, isEditable: false)) {
                var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package).Take(5).ToArray();
                Require(errors.Length == 0, "saved schema: " + string.Join("; ", errors.Select(error => error.Description)));
                ValidatePackage(package);
            }
            using var savedSource = new MemoryStream(Output, writable: false);
            using WordDocument reopened = WordDocument.Load(savedSource);
            Require(reopened.SourceFormat == WordFileFormat.Docx && reopened.LegacyDocUnsupportedFeatures.Count == 0,
                "saved document format and diagnostics");
            ValidateInspection(reopened.Inspect(includeImageBytes: false));
        }
        VerifiedUnits = Units;
    }

    public override void ReleaseResults() {
        _inspection = null;
        _report = null;
        _document?.Dispose();
        _document = null;
        base.ReleaseResults();
    }

    private void ValidateInspection(WordDocumentInfo inspection) {
        WordSectionInfo[] sections = inspection.Sections.ToArray();
        WordBlockInfo[] blocks = sections.SelectMany(section => section.Elements).ToArray();
        Require(sections.Length == 1 && sections[0].HeaderCount == 0 && sections[0].FooterCount == 0
            && blocks.Length == Units, "body section and paragraph count");
        var footnoteIds = new HashSet<long>();
        var endnoteIds = new HashSet<long>();
        var ledger = new StringBuilder();
        for (int index = 0; index < Units; index++) {
            Require(blocks[index] is WordParagraphInfo, "body paragraph kind " + index);
            var paragraph = (WordParagraphInfo)blocks[index];
            Require(paragraph.Text == BodyBefore(index) + BodyAfter(index) && !paragraph.IsListItem,
                "body text and reference anchor " + index);
            WordRunInfo[] references = paragraph.Runs.Where(run => run.Footnote != null || run.Endnote != null).ToArray();
            Require(references.Length == 1, "body reference count " + index);
            int referenceIndex = paragraph.Runs.ToList().IndexOf(references[0]);
            Require(string.Concat(paragraph.Runs.Take(referenceIndex).Select(run => run.Text)) == BodyBefore(index)
                && string.Concat(paragraph.Runs.Skip(referenceIndex + 1).Select(run => run.Text)) == BodyAfter(index),
                "body reference position " + index);
            WordRunInfo reference = references[0];
            long? id = IsEndnote(index) ? reference.Endnote?.ReferenceId : reference.Footnote?.ReferenceId;
            Require(id.HasValue && (IsEndnote(index) ? endnoteIds : footnoteIds).Add(id.Value)
                && (IsEndnote(index) ? reference.Footnote == null : reference.Endnote == null),
                "unique reference kind and identity " + index);
            IReadOnlyList<WordParagraphInfo> paragraphs = IsEndnote(index)
                ? reference.Endnote!.Paragraphs : reference.Footnote!.Paragraphs;
            Require(paragraphs.Count == 2 && paragraphs[0].Text == NoteFirst(index)
                && paragraphs[1].Text == NoteSecond(index), "complete ordered note text " + index);
            foreach (WordParagraphInfo noteParagraph in paragraphs) {
                Require(!noteParagraph.IsListItem && noteParagraph.Runs.All(run => run.Footnote == null && run.Endnote == null),
                    "ordinary nonrecursive note paragraphs " + index);
                foreach (WordRunInfo run in noteParagraph.Runs.Where(run => run.Text.Length != 0)) {
                    Require(!run.Bold && !run.Italic && !run.Underline, "plain note formatting " + index);
                }
            }
            ledger.Append(IsEndnote(index) ? 'E' : 'F').Append('|').Append(index).Append('|')
                .Append(paragraph.Text).Append('|').Append(paragraphs[0].Text).Append('|').Append(paragraphs[1].Text).Append('\n');
        }
        Require(footnoteIds.Count == Units / 2 && endnoteIds.Count == Units / 2, "inspected note totals");
        string digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(ledger.ToString())));
        if (SemanticDigest.Length == 0) SemanticDigest = digest;
        else Require(SemanticDigest == digest, "saved semantic ledger digest");
    }

    private void ValidatePackage(WordprocessingDocument package) {
        MainDocumentPart main = package.MainDocumentPart ?? throw new InvalidDataException("Missing main document part.");
        Body bodyElement = main.Document?.Body ?? throw new InvalidDataException("Missing document body.");
        Paragraph[] body = bodyElement.Elements<Paragraph>().ToArray();
        Require(body.Length == Units && bodyElement.ChildElements.Count == Units + 1,
            "saved body paragraphs and section properties");
        SectionProperties section = bodyElement.Elements<SectionProperties>().Single();
        FootnoteProperties? footnoteSettings = section.GetFirstChild<FootnoteProperties>();
        EndnoteProperties? endnoteSettings = section.GetFirstChild<EndnoteProperties>();
        Require(footnoteSettings?.GetFirstChild<FootnotePosition>()?.Val?.Value == FootnotePositionValues.PageBottom
            && main.DocumentSettingsPart?.Settings?.GetFirstChild<EndnoteDocumentWideProperties>()?
                .GetFirstChild<EndnotePosition>()?.Val?.Value == EndnotePositionValues.DocumentEnd,
            "note placement");
        foreach (OpenXmlCompositeElement? settings in new OpenXmlCompositeElement?[] { footnoteSettings, endnoteSettings }) {
            Require(settings != null && settings.GetFirstChild<NumberingFormat>()?.Val?.Value == NumberFormatValues.Decimal
                && settings.GetFirstChild<NumberingStart>()?.Val?.Value == 1
                && settings.GetFirstChild<NumberingRestart>()?.Val?.Value == RestartNumberValues.Continuous,
                "continuous Arabic note numbering from one");
        }
        var footnotes = main.FootnotesPart?.Footnotes?.Elements<Footnote>().Where(IsOrdinary).ToArray()
            ?? throw new InvalidDataException("Missing saved footnotes.");
        var endnotes = main.EndnotesPart?.Endnotes?.Elements<Endnote>().Where(IsOrdinary).ToArray()
            ?? throw new InvalidDataException("Missing saved endnotes.");
        Require(footnotes.Length == Units / 2 && endnotes.Length == Units / 2, "saved ordinary note counts");
        var footnoteMap = footnotes.ToDictionary(note => note.Id!.Value);
        var endnoteMap = endnotes.ToDictionary(note => note.Id!.Value);
        var referencedFootnotes = new HashSet<long>();
        var referencedEndnotes = new HashSet<long>();
        for (int index = 0; index < Units; index++) {
            Require(Text(body[index]) == BodyBefore(index) + BodyAfter(index), "saved body text " + index);
            FootnoteReference[] footnoteReferences = body[index].Descendants<FootnoteReference>().ToArray();
            EndnoteReference[] endnoteReferences = body[index].Descendants<EndnoteReference>().ToArray();
            Require(footnoteReferences.Length == (IsEndnote(index) ? 0 : 1)
                && endnoteReferences.Length == (IsEndnote(index) ? 1 : 0), "saved reference kind " + index);
            OpenXmlCompositeElement note;
            if (IsEndnote(index)) {
                long id = endnoteReferences[0].Id!.Value;
                Require(referencedEndnotes.Add(id) && endnoteMap.ContainsKey(id), "saved endnote reference " + index);
                note = endnoteMap[id];
            } else {
                long id = footnoteReferences[0].Id!.Value;
                Require(referencedFootnotes.Add(id) && footnoteMap.ContainsKey(id), "saved footnote reference " + index);
                note = footnoteMap[id];
            }
            Paragraph[] paragraphs = note.Elements<Paragraph>().ToArray();
            Require(note.ChildElements.Count == 2 && paragraphs.Length == 2
                && Text(paragraphs[0]) == NoteFirst(index) && Text(paragraphs[1]) == NoteSecond(index),
                "saved complete note paragraphs " + index);
            Require(note.Descendants<FootnoteReference>().Count() == 0 && note.Descendants<EndnoteReference>().Count() == 0,
                "saved nonrecursive note content " + index);
        }
        Require(referencedFootnotes.Count == footnotes.Length && referencedEndnotes.Count == endnotes.Length,
            "saved note definitions and references correspond exactly");

        static string Text(Paragraph paragraph) => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
        static bool IsOrdinary(FootnoteEndnoteType note) => note.Type == null || note.Type.Value == FootnoteEndnoteValues.Normal;
    }

    private static void Require(bool condition, string boundary) {
        if (!condition) throw new InvalidDataException("The Word 6/95 note workload changed " + boundary + ".");
    }
}
