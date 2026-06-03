using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Records_Warnings_For_Unsupported_HeaderFooter_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterWarnings.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterWarnings.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };
        options.Warnings.Add(new PdfExportWarning("Stale", "test", "This should be cleared before export."));

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("Native warning header text");
            header.AddParagraph().AddTextBox(string.Empty, WrapTextImage.Square);

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 1, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].Paragraphs[0].AddTextBox(string.Empty, WrapTextImage.Square);

            document.AddParagraph("Native warning body text");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "Stale");
        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeHeaderFooterTextBoxUnsupported" &&
            warning.Source == "default header");
        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeHeaderFooterTextBoxUnsupported" &&
            warning.Source == "default footer table");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native warning header text", text);
        Assert.Contains("Native warning body text", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Skips_Unsupported_HeaderFooter_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeUnsupportedHeaderImage.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeUnsupportedHeaderImage.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph().AddImage(imagePath, 32, 32);
            document.AddParagraph("Native unsupported header image body");
            document.Save();
        }

        ReplaceFirstHeaderImagePartWithGif(docPath);

        using (WordDocument document = WordDocument.Load(docPath)) {
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeHeaderFooterImageUnsupported" &&
            warning.Source == "default header image");
        Assert.True(File.Exists(pdfPath));
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Contains("Native unsupported header image body", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_TextBoxes() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterTextBoxes.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterTextBoxes.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordTextBox headerTextBox = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddTextBox("Native header text box");
            headerTextBox.HorizontalAlignment = WordHorizontalAlignmentValues.Center;

            WordParagraph footerParagraph = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph();
            WordTextBox footerTextBox = footerParagraph.AddTextBox("Native footer text box", WrapTextImage.Square);
            footerTextBox.HorizontalAlignment = WordHorizontalAlignmentValues.Right;

            document.AddParagraph("Native text box body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterTextBoxUnsupported");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header text box", text);
        Assert.Contains("Native footer text box", text);
        Assert.Contains("Native text box body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Shapes() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterShapes.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterShapes.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddShape(ShapeType.Rectangle, 36, 16, "#99ccff", "#003366", 1.5);

            WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            WordParagraph footerParagraph = footer.AddParagraph();
            footerParagraph.ParagraphAlignment = JustificationValues.Right;
            footerParagraph.AddShape(ShapeType.Rectangle, 34, 14, "#ffe699", "#663300", 1.25);

            document.AddParagraph("Native header footer shape body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterShapeUnsupported");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header footer shape body", text);

        string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.6 0.8 1 rg", content);
        Assert.Contains("1 0.902 0.6 rg", content);
        Assert.Contains("0 0.2 0.4 RG", content);
        Assert.Contains("0.4 0.2 0 RG", content);
        Assert.Contains(" re B", content);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterImages.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterImages.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerParagraph = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native image header");
            headerParagraph.ParagraphAlignment = JustificationValues.Center;
            headerParagraph.AddImage(imagePath, 32, 32);

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[2].Paragraphs[0].AddImage(imagePath, 32, 32);

            document.AddParagraph("Native header/footer image body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterImageUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        int imageObjectCount = rawPdf.Split(new[] { "/Subtype /Image" }, StringSplitOptions.None).Length - 1;
        Assert.True(imageObjectCount >= 2, "Expected native header and footer images to be emitted as image XObjects.");

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native image header", text);
        Assert.Contains("Native header/footer image body", text);
        Assert.DoesNotContain("Page 1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_PictureControls_To_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPictureControls.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPictureControls.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerParagraph = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native picture-control header");
            headerParagraph.ParagraphAlignment = JustificationValues.Center;
            headerParagraph.AddPictureControl(imagePath, 32, 32, "Header Logo", "HeaderLogo");

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[2].Paragraphs[0].AddPictureControl(imagePath, 32, 32, "Footer Logo", "FooterLogo");

            document.AddParagraph("Native header/footer picture-control body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        int imageObjectCount = rawPdf.Split(new[] { "/Subtype /Image" }, StringSplitOptions.None).Length - 1;
        Assert.True(imageObjectCount >= 2, "Expected native header and footer picture controls to be emitted as image XObjects.");

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native picture-control header", text);
        Assert.Contains("Native header/footer picture-control body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_RepeatingSections_To_Text_Items() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterRepeatingSections.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterRepeatingSections.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordRepeatingSection headerRepeating = RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header repeating section: ")
                .AddRepeatingSection("HeaderTasks", "HeaderTasks", "HeaderTasksTag");
            headerRepeating.SetTextItems(new[] { "Header item one", "Header item two" });

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            WordRepeatingSection footerRepeating = footerTable.Rows[0].Cells[2].Paragraphs[0]
                .AddRepeatingSection("FooterTasks", "FooterTasks", "FooterTasksTag");
            footerRepeating.SetTextItems(new[] { "Footer item one", "Footer item two" });

            document.AddParagraph("Native header/footer repeating-section body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header repeating section:", text);
        Assert.Contains("Header item one", text);
        Assert.Contains("Header item two", text);
        Assert.Contains("Footer item one", text);
        Assert.Contains("Footer item two", text);
        Assert.Contains("Native header/footer repeating-section body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_FormControls_To_Static_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFormControls.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFormControls.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("Native header approval: ").AddCheckBox(true, "Header Approval", "HeaderApproval");
            header.AddParagraph("Native header due: ").AddDatePicker(new DateTime(2026, 5, 31), "Header Due", "HeaderDue");

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "Native footer region: ";
            footerTable.Rows[0].Cells[0].Paragraphs[0].AddDropDownList(new[] { "North", "South" }, "Footer Region", "FooterRegion");
            footerTable.Rows[0].Cells[2].Paragraphs[0].Text = "Native footer status: ";
            footerTable.Rows[0].Cells[2].Paragraphs[0].AddComboBox(new[] { "Red", "Blue" }, "Footer Status", "FooterStatus", defaultValue: "Blue");

            document.AddParagraph("Native header/footer form-control body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        Assert.Empty(PdfCore.PdfInspector.Inspect(bytes).FormFields);

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header approval:", text);
        Assert.Contains("[x]", text);
        Assert.Contains("Native header due:", text);
        Assert.Contains("2026-05-31", text);
        Assert.Contains("Native footer region:", text);
        Assert.Contains("North", text);
        Assert.Contains("Native footer status:", text);
        Assert.Contains("Blue", text);
        Assert.Contains("Native header/footer form-control body", text);
    }
}
