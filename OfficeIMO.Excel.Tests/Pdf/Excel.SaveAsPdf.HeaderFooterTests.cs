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
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_HeaderFooter_Text_Zones() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooter.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Metric");
            sheet.Cell(2, 1, "HeaderFooterBody");
            sheet.SetHeaderFooter(
                headerLeft: "Left Header",
                headerCenter: "Dashboard Header",
                headerRight: "Page &P of &N",
                footerLeft: "Sheet &A",
                footerRight: "Right Footer");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(54)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Left Header", text);
        Assert.Contains("Dashboard Header", text);
        Assert.Contains("Page 1 of 1", text);
        Assert.Contains("Sheet Dashboard", text);
        Assert.Contains("Right Footer", text);
        Assert.Contains("HeaderFooterBody", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_HeaderFooter_DateTime_And_File_Fields() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterFields.xlsx");
        DateTime printedAt = new DateTime(2026, 5, 31, 14, 35, 0);

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            HeaderFooterDateTimeProvider = () => printedAt,
            PageSize = new PdfCore.PageSize(1800, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "HeaderFooterFieldBody");
            sheet.SetHeaderFooter(
                headerLeft: "Printed &D &T Dir &Z File &F",
                footerRight: "Page &P of &N");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Printed " + printedAt.ToString("d", CultureInfo.CurrentCulture), text);
        Assert.Contains(NormalizePdfTextSpaces(printedAt.ToString("t", CultureInfo.CurrentCulture)), NormalizePdfTextSpaces(text));
        Assert.Contains("Dir " + Path.GetDirectoryName(Path.GetFullPath(workbookPath)), text);
        Assert.Contains("File " + Path.GetFileName(workbookPath), text);
        Assert.Contains("Page 1 of 1", text);
        Assert.Contains("HeaderFooterFieldBody", text);
        Assert.DoesNotContain(options.Warnings, warning => warning.Feature == "WorksheetHeaderFooterField");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Simple_Worksheet_HeaderFooter_Formatting() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterFormatting.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Formatting")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "HeaderFooterFormattingBody");
            sheet.SetHeaderFooter(
                headerCenter: "&\"Arial,Bold\"&18&KFF0000Styled Header",
                footerCenter: "&\"Times New Roman,Italic\"&10&K0000FFStyled Footer");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Styled Header", text);
        Assert.Contains("Styled Footer", text);
        Assert.Contains("HeaderFooterFormattingBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 0 1 rg", rawPdf, StringComparison.Ordinal);
        Assert.Matches("Helvetica-Bold|Arial-Bold|Aptos-Bold|Calibri-Bold|LiberationSans-Bold|DejaVuSans-Bold", rawPdf);
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Times-Italic", "TimesNewRoman-Italic", "LiberationSerif-Italic", "DejaVuSerif-Italic");
        Assert.Contains(" 18 Tf", rawPdf, StringComparison.Ordinal);
        Assert.Contains(" 10 Tf", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain(options.Warnings, warning => warning.Feature == "WorksheetHeaderFooterFormatting");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Office_HeaderFooter_Font_Aliases() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterFontAliases.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "FontAliases")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "HeaderFooterFontAliasBody");
            sheet.SetHeaderFooter(
                headerCenter: "&\"Aptos,Bold\"Alias Header",
                footerCenter: "&\"Consolas,Italic\"Alias Footer");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Alias Header", text);
        Assert.Contains("Alias Footer", text);
        Assert.Contains("HeaderFooterFontAliasBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Matches("Helvetica-Bold|Arial-Bold|Aptos-Bold|Calibri-Bold|LiberationSans-Bold|DejaVuSans-Bold", rawPdf);
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Courier-Oblique", "Consolas-Italic", "LiberationMono-Italic", "DejaVuSansMono-Italic");
        Assert.DoesNotContain(options.Warnings, warning => warning.Feature == "WorksheetHeaderFooterFormatting");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Registered_Named_HeaderFooter_Fonts() {
        const string familyName = "Studio Serif";
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterNamedFont.xlsx");
        var options = new ExcelPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                    CompressContentStreams = false
                }
                .RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
                    familyName,
                    OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont())),
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "NamedFont")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "NamedHeaderFooterBody");
            sheet.SetHeaderFooter(
                headerCenter: "&\"Studio Serif,Bold\"Named Header",
                footerCenter: "&\"Studio Serif,Italic\"Named Footer");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Named Header", text);
        Assert.Contains("Named Footer", text);
        Assert.Contains("NamedHeaderFooterBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /StudioSerif-Bold", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /StudioSerif-Italic", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain(options.Warnings, warning => warning.Feature == "WorksheetFontSubstitution");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_DoesNotReserve_Escaped_HeaderFooter_Font_Tokens() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterEscapedFontToken.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54),
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "EscapedFontToken")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "EscapedHeaderFooterBody");
            sheet.SetHeaderFooter(headerCenter: "&&\"Times New Roman\" Literal Header");
            document.Save();

            bytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("\"Times New Roman\" Literal Header", text);
        Assert.Contains("EscapedHeaderFooterBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("TimesNewRoman", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Times-Roman", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_For_Mixed_Worksheet_HeaderFooter_Formatting() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterMixedFormatting.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "MixedFormatting")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "MixedFormattingBody");
            sheet.SetHeaderFooter(
                headerLeft: "&KFF0000Red Left",
                headerCenter: "Plain Center",
                footerCenter: "Plain Footer");
            document.Save();

            result = document.ToPdfDocumentResult(options);
            bytes = result.ToBytes();
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Red Left", text);
        Assert.Contains("Plain Center", text);
        Assert.Contains("MixedFormattingBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning => warning.Source == "MixedFormatting" && warning.Code == "WorksheetHeaderFooterFormatting");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_First_And_Even_HeaderFooter_Text_Zones() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterVariants.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Ledger")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Entry");
            sheet.Cell(1, 2, "Amount");
            for (int row = 2; row <= 90; row++) {
                sheet.Cell(row, 1, "Ledger row " + row.ToString(CultureInfo.InvariantCulture));
                sheet.Cell(row, 2, row * 10);
            }

            sheet.SetHeaderFooter(headerCenter: "Odd Header &A", footerCenter: "Odd Footer &P");
            sheet.SetFirstPageHeaderFooter(headerCenter: "First Header &A", footerCenter: "First Footer &P");
            sheet.SetEvenPageHeaderFooter(headerCenter: "Even Header &A", footerCenter: "Even Footer &P");

            ExcelSheet.HeaderFooterSnapshot snapshot = sheet.GetHeaderFooter();
            Assert.True(snapshot.DifferentFirstPage);
            Assert.True(snapshot.DifferentOddEven);
            Assert.Equal("Odd Header &A", snapshot.HeaderCenter);
            Assert.Equal("First Header &A", snapshot.FirstHeaderCenter);
            Assert.Equal("Even Header &A", snapshot.EvenHeaderCenter);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(48)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages >= 3);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        string thirdPage = pdf.GetPage(3).Text;
        Assert.Contains("First Header Ledger", firstPage);
        Assert.Contains("First Footer 1", firstPage);
        Assert.Contains("Even Header Ledger", secondPage);
        Assert.Contains("Even Footer 2", secondPage);
        Assert.Contains("Odd Header Ledger", thirdPage);
        Assert.Contains("Odd Footer 3", thirdPage);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Blank_First_And_Even_HeaderFooter_Variants() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfBlankHeaderFooterVariants.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Ledger")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Entry");
            for (int row = 2; row <= 90; row++) {
                sheet.Cell(row, 1, "Ledger row " + row.ToString(CultureInfo.InvariantCulture));
            }

            sheet.SetHeaderFooter(headerCenter: "Odd Header &A", footerCenter: "Odd Footer &P");
            sheet.SetFirstPageHeaderFooter();
            sheet.SetEvenPageHeaderFooter();
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(48)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages >= 3);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        string thirdPage = pdf.GetPage(3).Text;
        Assert.DoesNotContain("Odd Header Ledger", firstPage);
        Assert.DoesNotContain("Odd Footer 1", firstPage);
        Assert.DoesNotContain("Odd Header Ledger", secondPage);
        Assert.DoesNotContain("Odd Footer 2", secondPage);
        Assert.Contains("Odd Header Ledger", thirdPage);
        Assert.Contains("Odd Footer 3", thirdPage);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_HeaderFooter_Images() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterImages.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        byte[] disabledBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Metric");
            sheet.Cell(2, 1, "HeaderFooterImageBody");
            sheet.SetHeaderFooter(headerCenter: "Logo Header");
            sheet.SetHeaderImage(HeaderFooterPosition.Center, imageBytes, "image/png", widthPoints: 24, heightPoints: 16);

            ExcelSheet.HeaderFooterSnapshot snapshot = sheet.GetHeaderFooter();
            Assert.True(snapshot.HeaderHasPicturePlaceholder);
            Assert.NotNull(snapshot.HeaderCenterImage);
            Assert.Equal(HeaderFooterPosition.Center, snapshot.HeaderCenterImage!.Position);
            Assert.Equal("image/png", snapshot.HeaderCenterImage.ContentType);
            Assert.Equal(24, snapshot.HeaderCenterImage.WidthPoints);
            Assert.Equal(16, snapshot.HeaderCenterImage.HeightPoints);
            Assert.Equal(imageBytes, snapshot.HeaderCenterImage.Bytes);

            document.Save();

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(54)
            };
            bytes = document.ToPdf(options);

            disabledBytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetHeaderFooterImages = false,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(54)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Logo Header", text);
        Assert.Contains("HeaderFooterImageBody", text);

        var extractedImages = PdfCore.PdfImageExtractor.ExtractImages(bytes);
        var extractedImage = Assert.Single(extractedImages);
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("png", extractedImage.FileExtension);
        Assert.Equal("image/png", extractedImage.MimeType);
        Assert.True(extractedImage.IsImageFile);

        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(disabledBytes));
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Routes_HeaderFooter_Images_To_First_And_Even_Variants() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterVariantImages.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Ledger")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Entry");
            for (int row = 2; row <= 90; row++) {
                sheet.Cell(row, 1, "Ledger row " + row.ToString(CultureInfo.InvariantCulture));
            }

            sheet.SetHeaderFooter(headerCenter: "Odd Header &A", footerCenter: "Odd Footer &P");
            sheet.SetHeaderImage(HeaderFooterPosition.Center, imageBytes, "image/png", widthPoints: 24, heightPoints: 16);

            HeaderFooter headerFooter = sheet.WorksheetPart.Worksheet.GetFirstChild<HeaderFooter>()!;
            headerFooter.DifferentFirst = true;
            headerFooter.DifferentOddEven = true;
            headerFooter.OddHeader = new OddHeader("&COdd Header &A");
            headerFooter.FirstHeader = new FirstHeader("&C&GFirst Header &A");
            headerFooter.EvenHeader = new EvenHeader("&C&GEven Header &A");
            sheet.WorksheetPart.Worksheet.Save();

            ExcelSheet.HeaderFooterSnapshot snapshot = sheet.GetHeaderFooter();
            Assert.Equal("Odd Header &A", snapshot.HeaderCenter);
            Assert.Equal("&GFirst Header &A", snapshot.FirstHeaderCenter);
            Assert.Equal("&GEven Header &A", snapshot.EvenHeaderCenter);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(48)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages >= 3);
        Assert.Contains("First Header Ledger", pdf.GetPage(1).Text);
        Assert.Contains("Even Header Ledger", pdf.GetPage(2).Text);
        Assert.Contains("Odd Header Ledger", pdf.GetPage(3).Text);

        int[] imagePages = PdfCore.PdfImageExtractor
            .ExtractImages(bytes)
            .Select(image => image.PageNumber)
            .Distinct()
            .OrderBy(page => page)
            .ToArray();
        Assert.Contains(1, imagePages);
        Assert.Contains(2, imagePages);
        Assert.DoesNotContain(3, imagePages);
        Assert.All(imagePages, page => Assert.True(page == 1 || page % 2 == 0, "Expected header image only on first and even pages."));
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Can_Disable_Worksheet_HeaderFooter_Text_Zones() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterDisabled.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Metric");
            sheet.Cell(2, 1, "BodyOnly");
            sheet.SetHeaderFooter(headerCenter: "DoNotExportHeader", footerCenter: "DoNotExportFooter");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetHeadersAndFooters = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("BodyOnly", text);
        Assert.DoesNotContain("DoNotExportHeader", text);
        Assert.DoesNotContain("DoNotExportFooter", text);
    }

}
