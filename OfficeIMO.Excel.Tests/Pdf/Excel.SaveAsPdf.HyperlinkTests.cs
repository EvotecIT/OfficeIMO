using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_External_Cell_Hyperlinks() {
        const string linkUri = "https://github.com/EvotecIT/OfficeIMO";
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHyperlinks.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Links")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.SetHyperlink(2, 1, linkUri, display: "OfficeIMO");

            ExcelHyperlinkSnapshot hyperlink = Assert.Single(sheet.GetHyperlinks()).Value;
            Assert.True(hyperlink.IsExternal);
            Assert.Equal(linkUri, hyperlink.Target);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("OfficeIMO", text);

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByUri(linkUri));
        Assert.Equal("OfficeIMO", link.Contents);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Link", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/URI (https://github.com/EvotecIT/OfficeIMO)", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Internal_Cell_Hyperlinks_To_Cell_Destinations() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfInternalHyperlinks.xlsx");

        byte[] bytes;
        byte[] summaryOnlyBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.Cell(1, 1, "Name");
            summary.SetInternalLink(2, 1, "Details!B3", display: "Open Details B3");
            ExcelSheet details = document.AddWorkSheet("Details");
            details.Cell(1, 1, "Details Target");
            details.Cell(2, 1, "DestinationValue");
            details.Cell(3, 2, "CellSpecificTarget");

            ExcelHyperlinkSnapshot hyperlink = Assert.Single(summary.GetHyperlinks()).Value;
            Assert.False(hyperlink.IsExternal);
            Assert.Equal("'Details'!B3", hyperlink.Target);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });

            summaryOnlyBytes = document.ToPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "Summary" },
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        PdfCore.PdfNamedDestination destination = Assert.Single(logical.NamedDestinations, item => item.Name.EndsWith("-b3", StringComparison.Ordinal));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName(destination.Name));
        Assert.True(link.IsNamedDestinationLink);
        Assert.Equal("Open Details B3", link.Contents);
        Assert.Equal(destination.Name, link.DestinationName);
        Assert.Contains("details", destination.Name, StringComparison.Ordinal);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Link", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/S /GoTo", rawPdf, StringComparison.Ordinal);

        PdfCore.PdfLogicalDocument summaryOnly = PdfCore.PdfLogicalDocument.Load(summaryOnlyBytes);
        Assert.DoesNotContain(summaryOnly.NamedDestinations, item => item.Name.IndexOf("details", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(summaryOnly.Links, link => link.IsNamedDestinationLink);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_SameSheet_Internal_Cell_Hyperlinks_To_Cell_Destinations() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSameSheetInternalHyperlinks.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Links")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Top Target");
            sheet.SetInternalLink(2, 1, "A1", display: "Back to Top");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        PdfCore.PdfNamedDestination destination = Assert.Single(logical.NamedDestinations, item => item.Name.EndsWith("-a1", StringComparison.Ordinal));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName(destination.Name));
        Assert.Equal("Back to Top", link.Contents);
        Assert.Equal(destination.Name, link.DestinationName);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Drops_Internal_Cell_Hyperlinks_To_Unexported_Target_Cells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfInternalHyperlinkHiddenTarget.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.Cell(1, 1, "Name");
            summary.SetInternalLink(2, 1, "Details!B200", display: "Open Details B200");
            ExcelSheet details = document.AddWorkSheet("Details");
            details.Cell(1, 1, "Details Header");
            details.Cell(2, 1, "Visible Detail");
            details.Cell(200, 2, "Hidden Target");
            document.SetPrintArea(details, "A1:B2");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        Assert.Contains(logical.NamedDestinations, item => item.Name.IndexOf("details", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(logical.NamedDestinations, item => item.Name.EndsWith("-b200", StringComparison.Ordinal));
        Assert.DoesNotContain(logical.Links, link => link.IsNamedDestinationLink && link.Contents == "Open Details B200");
    }

}
