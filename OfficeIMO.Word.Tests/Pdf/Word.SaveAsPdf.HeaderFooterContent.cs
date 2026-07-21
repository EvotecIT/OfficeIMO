using DocumentFormat.OpenXml;
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
        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("Native warning header text");
            header.AddParagraph().AddTextBox(string.Empty, WrapTextImage.Square);

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 1, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].Paragraphs[0].AddTextBox(string.Empty, WrapTextImage.Square);

            document.AddParagraph("Native warning body text");
            document.Save();
            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            result.Save(pdfPath);

            Assert.Contains(result.Warnings, warning =>
                warning.Code == "NativeHeaderFooterTextBoxUnsupported" &&
                warning.Source == "default header");
            Assert.Contains(result.Warnings, warning =>
                warning.Code == "NativeHeaderFooterTextBoxUnsupported" &&
                warning.Source == "default footer table");
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native warning header text", text);
        Assert.Contains("Native warning body text", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Omits_Hidden_HeaderFooter_Text_Runs() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHiddenHeaderFooterText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHiddenHeaderFooterText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerParagraph = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph();
            headerParagraph.AddText("VisibleHeaderStart");
            WordParagraph hiddenHeader = headerParagraph.AddText("HiddenHeaderRun");
            hiddenHeader._run!.RunProperties ??= new RunProperties();
            hiddenHeader._run.RunProperties.Vanish = new Vanish();
            headerParagraph.AddText("VisibleHeaderEnd");

            WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            WordParagraph hiddenFooter = footer.AddParagraph("HiddenFooterOnly");
            hiddenFooter._run!.RunProperties ??= new RunProperties();
            hiddenFooter._run.RunProperties.Vanish = new Vanish();
            footer.AddParagraph("VisibleFooterAfterHidden");

            document.AddParagraph("Hidden header footer body");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("VisibleHeaderStart", text);
        Assert.Contains("VisibleHeaderEnd", text);
        Assert.Contains("VisibleFooterAfterHidden", text);
        Assert.Contains("Hidden header footer body", text);
        Assert.DoesNotContain("HiddenHeaderRun", text);
        Assert.DoesNotContain("HiddenFooterOnly", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Caps_HeaderFooter_Text_Runs() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCapsHeaderFooterText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCapsHeaderFooterText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerParagraph = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph();
            headerParagraph.AddText("headerBeforeCaps ");
            WordParagraph headerCaps = headerParagraph.AddText("capsHeaderRun");
            headerCaps._run!.RunProperties ??= new RunProperties();
            headerCaps._run.RunProperties.Caps = new Caps();

            WordParagraph footerParagraph = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph();
            WordParagraph footerCaps = footerParagraph.AddText("capsFooterRun");
            footerCaps._run!.RunProperties ??= new RunProperties();
            footerCaps._run.RunProperties.Caps = new Caps();

            document.AddParagraph("Caps header footer body");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("headerBeforeCaps", text);
        Assert.Contains("CAPSHEADERRUN", text);
        Assert.Contains("CAPSFOOTERRUN", text);
        Assert.Contains("Caps header footer body", text);
        Assert.DoesNotContain("capsHeaderRun", text);
        Assert.DoesNotContain("capsFooterRun", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Normalizes_Gif_HeaderFooter_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeGifHeaderImage.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeGifHeaderImage.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph().AddImage(imagePath, 32, 32);
            document.AddParagraph("Native GIF header image body");
            document.Save();
        }

        ReplaceFirstHeaderImagePartWithGif(docPath);

        using (WordDocument document = WordDocument.Load(docPath)) {
            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            result.Save(pdfPath);
            Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeHeaderFooterImageUnsupported");
        }
        Assert.True(File.Exists(pdfPath));
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Contains("Native GIF header image body", pdf.GetPage(1).Text);
        Assert.NotEmpty(PdfCore.PdfDocument.Open(pdfPath).Read.ImagePlacements());
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
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Vml_TextPath_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterVmlTextPath.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterVmlTextPath.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header._header.Append(CreateNativeHeaderFooterTextPathParagraph("Native WordArt header"));

            document.AddParagraph("Native textpath body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native WordArt header", text);
        Assert.Contains("Native textpath body", text);
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
        int imageDrawCount = ReadPdfPageContent(bytes).Split(new[] { "/Im" }, StringSplitOptions.None).Length - 1;
        Assert.True(imageObjectCount >= 1, "Expected native header and footer images to be emitted as image XObjects.");
        Assert.True(imageDrawCount >= 2, "Expected native header and footer image placements to draw image XObjects.");

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
        int imageDrawCount = ReadPdfPageContent(bytes).Split(new[] { "/Im" }, StringSplitOptions.None).Length - 1;
        Assert.True(imageObjectCount >= 1, "Expected native header and footer picture controls to be emitted as image XObjects.");
        Assert.True(imageDrawCount >= 2, "Expected native header and footer picture-control placements to draw image XObjects.");

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

    private static Paragraph CreateNativeHeaderFooterTextPathParagraph(string text) {
        var textPath = new DocumentFormat.OpenXml.Vml.TextPath {
            On = true,
            FitShape = true,
            Style = "font-family:\"Calibri\";font-size:1pt",
            String = text
        };

        var shape = new DocumentFormat.OpenXml.Vml.Shape(textPath) {
            Id = "NativeHeaderFooterTextPath",
            Type = "#_x0000_t136",
            Style = "position:absolute;margin-left:0;margin-top:0;width:320pt;height:72pt;rotation:315;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
            FillColor = "silver",
            Stroked = false
        };
        shape.SetAttribute(new OpenXmlAttribute("allowincell", "urn:schemas-microsoft-com:office:office", "false"));

        return new Paragraph(new Run(new Picture(shape)));
    }
}
