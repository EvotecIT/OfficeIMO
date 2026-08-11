using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xceed.Document.NET;
using NpoiDocument = NPOI.XWPF.UserModel.XWPFDocument;
using NpoiHeaderFooterType = NPOI.WP.UserModel.HeaderFooterType;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using OpenXmlFooter = DocumentFormat.OpenXml.Wordprocessing.Footer;
using OpenXmlHeader = DocumentFormat.OpenXml.Wordprocessing.Header;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using XceedDocX = Xceed.Words.NET.DocX;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Compares equivalent high-level and format-native DOCX creation workloads.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "Create")]
public class WordCreateParagraphComparisonBenchmarks {
    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Validate() {
        WordBenchmarkCorpus.ValidateParagraphDocument(OfficeIMO(), ItemCount);
        WordBenchmarkCorpus.ValidateParagraphDocument(DocX(), ItemCount, requireOpenXmlSdkConformance: false);
        WordBenchmarkCorpus.ValidateParagraphDocument(NPOI(), ItemCount);
        WordBenchmarkCorpus.ValidateParagraphDocument(OpenXmlSdk(), ItemCount);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() {
        using WordDocument document = WordDocument.Create();
        for (int index = 0; index < ItemCount; index++) {
            document.AddParagraph(WordBenchmarkCorpus.ParagraphText(index));
        }
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] DocX() {
        using var stream = new MemoryStream();
        using (XceedDocX document = XceedDocX.Create(stream)) {
            for (int index = 0; index < ItemCount; index++) {
                document.InsertParagraph(WordBenchmarkCorpus.ParagraphText(index));
            }
            document.Save();
        }
        return stream.ToArray();
    }

    [Benchmark]
    public byte[] NPOI() {
        using var document = new NpoiDocument();
        for (int index = 0; index < ItemCount; index++) {
            document.CreateParagraph().CreateRun().SetText(WordBenchmarkCorpus.ParagraphText(index));
        }
        using var stream = new MemoryStream();
        document.Write(stream);
        return stream.ToArray();
    }

    [Benchmark]
    public byte[] OpenXmlSdk() {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
                   stream,
                   WordprocessingDocumentType.Document,
                   autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            var body = new Body();
            for (int index = 0; index < ItemCount; index++) {
                body.Append(new OpenXmlParagraph(new OpenXmlRun(new Text(WordBenchmarkCorpus.ParagraphText(index)))));
            }
            mainPart.Document = new OpenXmlDocument(body);
        }
        return stream.ToArray();
    }
}

/// <summary>Compares creation of a styled report with headers, footers, and a two-column table.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "StructuredReport")]
public class WordCreateReportComparisonBenchmarks {
    [Params(100, 1000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Validate() {
        WordBenchmarkCorpus.ValidateReportDocument(OfficeIMO(), RowCount);
        WordBenchmarkCorpus.ValidateReportDocument(DocX(), RowCount, requireOpenXmlSdkConformance: false);
        WordBenchmarkCorpus.ValidateReportDocument(NPOI(), RowCount);
        WordBenchmarkCorpus.ValidateReportDocument(OpenXmlSdk(), RowCount);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() {
        using WordDocument document = WordDocument.Create();
        document.HeaderDefaultOrCreate.AddParagraph(WordBenchmarkCorpus.ReportHeader);
        document.FooterDefaultOrCreate.AddParagraph(WordBenchmarkCorpus.ReportFooter);
        document.AddParagraph(WordBenchmarkCorpus.ReportTitle).SetBold().SetFontSize(18);
        document.AddParagraph(WordBenchmarkCorpus.ReportSummary);
        WordTable table = document.AddTable(RowCount + 1, 2);
        table.Rows[0].Cells[0].Paragraphs[0].SetText("Record").SetBold();
        table.Rows[0].Cells[1].Paragraphs[0].SetText("Owner").SetBold();
        for (int index = 0; index < RowCount; index++) {
            table.Rows[index + 1].Cells[0].Paragraphs[0].SetText(WordBenchmarkCorpus.RecordId(index));
            table.Rows[index + 1].Cells[1].Paragraphs[0].SetText(WordBenchmarkCorpus.RecordOwner(index));
        }
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] DocX() {
        using var stream = new MemoryStream();
        using (XceedDocX document = XceedDocX.Create(stream)) {
            document.AddHeaders();
            document.Headers.Odd.InsertParagraph(WordBenchmarkCorpus.ReportHeader);
            document.AddFooters();
            document.Footers.Odd.InsertParagraph(WordBenchmarkCorpus.ReportFooter);
            document.InsertParagraph(WordBenchmarkCorpus.ReportTitle).Bold().FontSize(18);
            document.InsertParagraph(WordBenchmarkCorpus.ReportSummary);
            Xceed.Document.NET.Table table = document.AddTable(RowCount + 1, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Append("Record").Bold();
            table.Rows[0].Cells[1].Paragraphs[0].Append("Owner").Bold();
            for (int index = 0; index < RowCount; index++) {
                table.Rows[index + 1].Cells[0].Paragraphs[0].Append(WordBenchmarkCorpus.RecordId(index));
                table.Rows[index + 1].Cells[1].Paragraphs[0].Append(WordBenchmarkCorpus.RecordOwner(index));
            }
            document.InsertTable(table);
            document.Save();
        }
        return stream.ToArray();
    }

    [Benchmark]
    public byte[] NPOI() {
        using var document = new NpoiDocument();
        var header = document.CreateHeader(NpoiHeaderFooterType.DEFAULT);
        header.CreateParagraph().CreateRun().SetText(WordBenchmarkCorpus.ReportHeader);
        var footer = document.CreateFooter(NpoiHeaderFooterType.DEFAULT);
        footer.CreateParagraph().CreateRun().SetText(WordBenchmarkCorpus.ReportFooter);

        var titleRun = document.CreateParagraph().CreateRun();
        titleRun.IsBold = true;
        titleRun.FontSize = 18;
        titleRun.SetText(WordBenchmarkCorpus.ReportTitle);
        document.CreateParagraph().CreateRun().SetText(WordBenchmarkCorpus.ReportSummary);

        var table = document.CreateTable(RowCount + 1, 2);
        SetNpoiCell(table.GetRow(0).GetCell(0), "Record", bold: true);
        SetNpoiCell(table.GetRow(0).GetCell(1), "Owner", bold: true);
        for (int index = 0; index < RowCount; index++) {
            SetNpoiCell(table.GetRow(index + 1).GetCell(0), WordBenchmarkCorpus.RecordId(index), bold: false);
            SetNpoiCell(table.GetRow(index + 1).GetCell(1), WordBenchmarkCorpus.RecordOwner(index), bold: false);
        }

        using var stream = new MemoryStream();
        document.Write(stream);
        return stream.ToArray();
    }

    [Benchmark]
    public byte[] OpenXmlSdk() {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
                   stream,
                   WordprocessingDocumentType.Document,
                   autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new OpenXmlHeader(CreateParagraph(WordBenchmarkCorpus.ReportHeader));
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new OpenXmlFooter(CreateParagraph(WordBenchmarkCorpus.ReportFooter));

            string headerId = mainPart.GetIdOfPart(headerPart);
            string footerId = mainPart.GetIdOfPart(footerPart);
            var body = new Body(
                CreateParagraph(WordBenchmarkCorpus.ReportTitle, bold: true, fontSizeHalfPoints: "36"),
                CreateParagraph(WordBenchmarkCorpus.ReportSummary),
                CreateTable(RowCount),
                new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = footerId }));
            mainPart.Document = new OpenXmlDocument(body);
        }
        return stream.ToArray();
    }

    private static OpenXmlParagraph CreateParagraph(string text, bool bold = false, string? fontSizeHalfPoints = null) {
        var runProperties = new RunProperties();
        if (bold) runProperties.Append(new Bold());
        if (fontSizeHalfPoints is not null) runProperties.Append(new FontSize { Val = fontSizeHalfPoints });
        return new OpenXmlParagraph(new OpenXmlRun(runProperties, new Text(text)));
    }

    private static OpenXmlTable CreateTable(int rowCount) {
        var table = new OpenXmlTable(
            new TableProperties(new TableStyle { Val = "TableGrid" }),
            new TableGrid(new GridColumn(), new GridColumn()));
        table.Append(CreateRow("Record", "Owner", bold: true));
        for (int index = 0; index < rowCount; index++) {
            table.Append(CreateRow(
                WordBenchmarkCorpus.RecordId(index),
                WordBenchmarkCorpus.RecordOwner(index),
                bold: false));
        }
        return table;
    }

    private static TableRow CreateRow(string first, string second, bool bold) =>
        new(
            new TableCell(CreateParagraph(first, bold)),
            new TableCell(CreateParagraph(second, bold)));

    private static void SetNpoiCell(NPOI.XWPF.UserModel.XWPFTableCell cell, string text, bool bold) {
        var run = cell.Paragraphs[0].CreateRun();
        run.IsBold = bold;
        run.SetText(text);
    }
}
