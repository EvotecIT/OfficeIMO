using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using static OfficeIMO.Migration.Benchmarks.Word69BenchmarkInput;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Imports ordered Word 6 runs whose formatting pages share two property records.</summary>
internal sealed class Word69Workload : MigrationWorkload {
    private WordDocument? _document;
    private WordDocumentInfo? _inspection;
    private readonly string _expectedText;

    internal Word69Workload(int units) : base(CreateInput(units), units) {
        _expectedText = string.Concat(Enumerable.Repeat("Ab", units));
    }

    public override void Execute(string operation) {
        using var input = new MemoryStream(Input, writable: false);
        _document = WordDocument.Load(input);
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
        if (_document == null) throw new InvalidOperationException("The workload did not load a document.");
        DiagnosticDigest = string.Join("|", _document.LegacyDocUnsupportedFeatures.Select(finding => finding.Code));
        if (_document.SourceFormat != WordFileFormat.Doc || DiagnosticDigest.Length != 0) {
            throw new InvalidDataException("Word 6 input did not retain its expected format and loss-free diagnostics: " + DiagnosticDigest);
        }

        if (Output != null) {
            using (var stream = new MemoryStream(Output, writable: false))
            using (var package = WordprocessingDocument.Open(stream, isEditable: false)) {
                var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package).Take(5).ToArray();
                if (errors.Length != 0) {
                    throw new InvalidDataException("Word 6 DOCX output failed schema validation: " + string.Join("; ", errors.Select(error => error.Description)));
                }
            }
            using var output = new MemoryStream(Output, writable: false);
            using WordDocument reopened = WordDocument.Load(output);
            ValidateInspection(reopened.Inspect(includeImageBytes: false));
        } else {
            ValidateInspection(_inspection ?? throw new InvalidOperationException("The workload did not inspect a document."));
        }
        VerifiedUnits = Units;
    }

    public override void ReleaseResults() {
        _inspection = null;
        _document?.Dispose();
        _document = null;
        base.ReleaseResults();
    }

    private void ValidateInspection(WordDocumentInfo inspection) {
        WordSectionInfo[] sections = inspection.Sections.ToArray();
        WordBlockInfo[] blocks = sections.SelectMany(section => section.Elements).ToArray();
        if (sections.Length != 1 || blocks.Length != 1 || blocks[0] is not WordParagraphInfo paragraph
            || sections[0].HeaderCount != 0 || sections[0].FooterCount != 0) {
            throw new InvalidDataException("Word 6 workload must retain exactly one ordinary body paragraph.");
        }
        if (paragraph.Text != _expectedText || paragraph.Runs.Count != Units || paragraph.IsListItem) {
            throw new InvalidDataException("Word 6 workload changed its text, run count or paragraph kind.");
        }
        for (int index = 0; index < Units; index++) {
            WordRunInfo run = paragraph.Runs[index];
            if (run.Text != "Ab" || run.Bold != (index % 2 == 0) || run.Italic || run.Underline) {
                throw new InvalidDataException("Word 6 workload changed text or formatting at run " + index + ".");
            }
        }
    }

    private static byte[] CreateInput(int units) {
        if (units < 1 || units > 10000) throw new ArgumentOutOfRangeException(nameof(units));
        const int textStart = 0x300;
        const int runsPerPage = 80;
        int textLength = checked(units * 2 + 1);
        int textEnd = textStart + textLength;
        int pageCount = (units + runsPerPage - 1) / runsPerPage;
        int firstPage = (textEnd + 511) & ~511;
        int binStart = firstPage + pageCount * 512;
        int binLength = 4 + pageCount * 6;
        var main = new byte[binStart + binLength];
        WriteWord(main, 0, 0xA5DC);
        WriteWord(main, 2, 101);
        WriteWord(main, 6, 0x0409);
        WriteInt(main, 0x18, textStart);
        WriteInt(main, 0x1C, textEnd);
        WriteInt(main, 0x20, main.Length);
        WriteInt(main, 0x34, textLength);
        WriteInt(main, 0xB8, binStart);
        WriteInt(main, 0xBC, binLength);
        WriteWord(main, 0x18A, checked((ushort)(firstPage / 512)));
        WriteWord(main, 0x18E, checked((ushort)pageCount));
        Encoding.ASCII.GetBytes(string.Concat(Enumerable.Repeat("Ab", units)) + "\r").CopyTo(main, textStart);

        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            int firstRun = pageIndex * runsPerPage;
            int count = Math.Min(runsPerPage, units - firstRun);
            int page = firstPage + pageIndex * 512;
            int positionsEnd = (count + 1) * 4;
            WriteInt(main, binStart + pageIndex * 4, textStart + firstRun * 2);
            WriteWord(main, binStart + (pageCount + 1) * 4 + pageIndex * 2, checked((ushort)(page / 512)));
            for (int index = 0; index <= count; index++) {
                int position = textStart + (firstRun + index) * 2;
                if (firstRun + index == units) position++;
                WriteInt(main, page + index * 4, position);
                if (index < count) main[page + positionsEnd + index] = (byte)((firstRun + index) % 2 == 0 ? 250 : 252);
            }
            main[page + 500] = 2;
            main[page + 501] = 85;
            main[page + 502] = 1;
            main[page + 504] = 2;
            main[page + 505] = 85;
            main[page + 506] = 0;
            main[page + 511] = checked((byte)count);
        }
        WriteInt(main, binStart + pageCount * 4, textEnd);
        return WrapCompound(main);
    }

}
