using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Measures existing WordML formatting projection independently of new list coverage.</summary>
internal sealed class OfficeXmlWorkload : MigrationWorkload {
    private WordDocument? _document;
    private WordDocumentInfo? _inspection;
    private OfficeConversionReport? _report;

    internal OfficeXmlWorkload(int units) : base(CreateInput(units), units) { }

    public override void Execute(string operation) {
        using var source = new MemoryStream(Input, writable: false);
        _document = WordDocument.Load(source, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        _report = report;
        if (operation == "ConvertSave") Output = _document.ToBytes();
        else if (operation == "LoadInspect") _inspection = _document.Inspect(includeImageBytes: false);
        else throw new ArgumentOutOfRangeException(nameof(operation));
    }

    public override void Validate() {
        if (_document == null || _report == null) throw new InvalidOperationException("The WordML operation did not load a document.");
        if (_report.Diagnostics.Count != 0 || _report.HasLoss) throw new InvalidDataException("The repeated formatting workload reported an unexpected diagnostic.");
        DiagnosticDigest = string.Empty;
        if (_inspection != null) {
            WordParagraphInfo[] paragraphs = _inspection.Sections.SelectMany(section => section.Elements).OfType<WordParagraphInfo>().ToArray();
            Require(paragraphs.Length == Units, "public inspection paragraph count");
            for (int index = 0; index < paragraphs.Length; index++) {
                Require(paragraphs[index].Text == RowText(index), "public inspection text " + index);
                Require(paragraphs[index].Runs.Count == 1 && paragraphs[index].Runs[0].Text == RowText(index), "public inspection runs " + index);
            }
        }
        Inspect(_document.OpenXmlDocument);
        if (Output != null) {
            using var bytes = new MemoryStream(Output, writable: false);
            using WordDocument reopened = WordDocument.Load(bytes, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport reopenedReport);
            if (reopenedReport.Diagnostics.Count != 0) throw new InvalidDataException("Saved formatting workload reported an unexpected diagnostic.");
            Inspect(reopened.OpenXmlDocument);
        }
        VerifiedUnits = Units;
    }

    public override void ReleaseResults() {
        _document?.Dispose();
        _document = null;
        _report = null;
        _inspection = null;
        base.ReleaseResults();
    }

    private void Inspect(WordprocessingDocument document) {
        if (new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document).Any())
            throw new InvalidDataException("The repeated formatting workload produced invalid Open XML.");
        Paragraph[] paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToArray();
        Require(paragraphs.Length == Units, "paragraph count");
        for (int index = 0; index < paragraphs.Length; index++) {
            Paragraph paragraph = paragraphs[index];
            Require(paragraph.InnerText == RowText(index), "row text " + index);
            ParagraphProperties? properties = paragraph.ParagraphProperties;
            SpacingBetweenLines? spacing = properties?.SpacingBetweenLines;
            Require(spacing?.Before?.Value == "120" && spacing.After?.Value == "60" && spacing.Line?.Value == "276" &&
                spacing.LineRule?.Value == LineSpacingRuleValues.Auto, "paragraph spacing " + index);
            Indentation? indentation = properties?.Indentation;
            Require((indentation?.Left ?? indentation?.Start)?.Value == "360" && indentation?.Hanging?.Value == "120", "paragraph indentation " + index);
            Require(properties?.Justification?.Val?.Value == JustificationValues.Both, "paragraph alignment " + index);
            Require(properties!.ChildElements.Select(child => child.LocalName).SequenceEqual(new[] { "spacing", "ind", "jc" }), "paragraph property order " + index);
            Run[] runs = paragraph.Elements<Run>().ToArray();
            Require(runs.Length == 1, "run count " + index);
            RunProperties? run = runs[0].RunProperties;
            Require(run?.RunFonts?.Ascii?.Value == "Arial" && run.RunFonts.HighAnsi?.Value == "Arial", "run fonts " + index);
            Require(run?.Bold != null && run.Color?.Val?.Value == "003366" && run.FontSize?.Val?.Value == "22", "run formatting " + index);
            Require(run!.ChildElements.Select(child => child.LocalName).SequenceEqual(new[] { "rFonts", "b", "color", "sz" }), "run property order " + index);
        }
    }

    private static void Require(bool condition, string boundary) {
        if (!condition) throw new InvalidDataException("The repeated WordML formatting workload changed " + boundary + ".");
    }

    private static string RowText(int index) => "Property row " + index.ToString("D5", CultureInfo.InvariantCulture);

    private static byte[] CreateInput(int units) {
        if (units < 1) throw new ArgumentOutOfRangeException(nameof(units));
        var xml = new StringBuilder("<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\"><w:body>");
        for (int index = 0; index < units; index++) {
            xml.Append("<w:p><w:pPr><w:spacing w:before=\"120\" w:after=\"60\" w:line=\"276\" w:line-rule=\"auto\"/>")
                .Append("<w:ind w:left=\"360\" w:hanging=\"120\"/><w:jc w:val=\"both\"/></w:pPr>")
                .Append("<w:r><w:rPr><w:rFonts w:ascii=\"Arial\" w:h-ansi=\"Arial\"/><w:b/><w:color w:val=\"003366\"/><w:sz w:val=\"22\"/></w:rPr><w:t>")
                .Append(RowText(index)).Append("</w:t></w:r></w:p>");
        }
        xml.Append("</w:body></w:wordDocument>");
        return Encoding.UTF8.GetBytes(xml.ToString());
    }
}
