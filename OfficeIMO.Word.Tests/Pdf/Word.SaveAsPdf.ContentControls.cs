using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using M = DocumentFormat.OpenXml.Math;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Simple_Equations_To_Static_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleEquations.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleEquations.pdf");
        const string headerOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>h=2</m:t></m:r></m:oMath></m:oMathPara>";
        const string bodyOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>";
        const string tableOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>c=4</m:t></m:r></m:oMath></m:oMathPara>";
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerEquation = RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header equation: ");
            headerEquation.AddEquation(headerOmml);
            headerEquation.AddText(" header-suffix");

            document.AddParagraph("Native body equation:").AddEquation(bodyOmml);

            WordParagraph controlledEquation = document.AddParagraph("Native controlled equation:");
            controlledEquation._paragraph.Append(new SdtRun(
                new SdtProperties(new SdtId { Val = 2076 }),
                new SdtContentRun(
                    new Run(new Text(" control-prefix ")),
                    new M.OfficeMath(new M.Run(new M.Text("controlled=5"))),
                    new Run(new Text(" control-suffix")))));

            WordParagraph interleavedEquation = document.AddParagraph("pdf-prefix ");
            interleavedEquation._paragraph.Append(new M.OfficeMath(new M.Run(new M.Text("pdf-equation"))));
            interleavedEquation.AddText(" pdf-suffix");

            WordParagraph linkedEquation = document.AddParagraph();
            HyperlinkRelationship linkedEquationRelationship = document._wordprocessingDocument.MainDocumentPart!
                .AddHyperlinkRelationship(new Uri("https://officeimo.net/equations"), true);
            linkedEquation._paragraph.Append(new Hyperlink(
                new Run(new RunProperties(new Bold()), new Text("Qlinked-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("linked-equation"))),
                new Run(new RunProperties(new Italic()), new Text(" Xlinked-suffix"))) {
                Id = linkedEquationRelationship.Id
            });

            WordParagraph nestedLinkedEquation = document.AddParagraph();
            HyperlinkRelationship nestedLinkedEquationRelationship = document._wordprocessingDocument.MainDocumentPart!
                .AddHyperlinkRelationship(new Uri("https://officeimo.net/nested-equations"), true);
            nestedLinkedEquation._paragraph.Append(new Hyperlink(
                new SdtRun(
                    new SdtProperties(new SdtId { Val = 2078 }),
                    new SdtContentRun(
                        new Run(new Text("nested-link-prefix ")),
                        new M.OfficeMath(new M.Run(new M.Text("nested-link-equation"))),
                        new Run(new Text(" nested-link-suffix"))))) {
                Id = nestedLinkedEquationRelationship.Id
            });

            WordParagraph equationWithBreak = document.AddParagraph("break-prefix ");
            equationWithBreak._paragraph.Append(
                new M.OfficeMath(new M.Run(new M.Text("break-equation"))),
                new Run(new Break()),
                new Run(new Text("break-suffix")));

            WordParagraph hiddenAdjacentText = document.AddParagraph("visible-prefix ");
            hiddenAdjacentText._paragraph.Append(
                new Run(new RunProperties(new Vanish()), new Text("hidden-equation-adjacent ")),
                new M.OfficeMath(new M.Run(new M.Text("visible-equation"))),
                new Run(new Text(" visible-suffix")));

            WordTable table = document.AddTable(2, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Native table equation:";
            table.Rows[0].Cells[0].Paragraphs[0].AddEquation(tableOmml);
            WordParagraph linkedTableEquation = table.Rows[1].Cells[0].Paragraphs[0];
            HyperlinkRelationship linkedTableEquationRelationship = document._wordprocessingDocument.MainDocumentPart!
                .AddHyperlinkRelationship(new Uri("https://officeimo.net/table-equations"), true);
            linkedTableEquation._paragraph.Append(new Hyperlink(
                new Run(new RunProperties(new Bold()), new Text("Ttable-link-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("table-link-equation"))),
                new Run(new RunProperties(new Italic()), new Text(" Ytable-link-suffix"))) {
                Id = linkedTableEquationRelationship.Id
            });
            HyperlinkRelationship secondaryTableRelationship = document._wordprocessingDocument.MainDocumentPart!
                .AddHyperlinkRelationship(new Uri("https://officeimo.net/table-secondary"), true);
            linkedTableEquation._paragraph.Append(new Hyperlink(new Run(new Text(" secondary-table-link"))) {
                Id = secondaryTableRelationship.Id
            });

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterEquationUnsupported");
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyEquationUnsupported");

        string text = PdfCore.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native header equation: h=2 header-suffix", NormalizePdfText(text));
        Assert.Contains("Native body equation:", text);
        Assert.Contains("(a)/(b)", NormalizePdfText(text));
        string normalizedText = NormalizePdfText(text);
        Assert.Contains("Native controlled equation:", normalizedText);
        Assert.Contains("Native controlled equation: control-prefix controlled=5 control-suffix", normalizedText);
        Assert.Contains("pdf-prefix pdf-equation pdf-suffix", normalizedText);
        Assert.Contains("Qlinked-prefix linked-equation Xlinked-suffix", normalizedText);
        Assert.Contains("nested-link-prefix nested-link-equation nested-link-suffix", normalizedText);
        Assert.Contains("break-prefix break-equation break-suffix", normalizedText);
        Assert.Contains("visible-prefix visible-equation visible-suffix", normalizedText);
        Assert.DoesNotContain("hidden-equation-adjacent", normalizedText, StringComparison.Ordinal);
        Assert.Contains("Native table equation:", normalizedText);
        Assert.Contains("c=4", normalizedText);
        Assert.Contains("Ttable-link-prefix table-link-equation Ytable-link-suffix secondary-table-link", normalizedText);
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            Assert.Contains("Bold", Assert.Single(page.Letters, letter => letter.Value == "Q").FontName, StringComparison.OrdinalIgnoreCase);
            string suffixFont = Assert.Single(page.Letters, letter => letter.Value == "X").FontName;
            Assert.True(
                suffixFont.Contains("Italic", StringComparison.OrdinalIgnoreCase) ||
                suffixFont.Contains("Oblique", StringComparison.OrdinalIgnoreCase),
                suffixFont);
        }
        Assert.Contains("/URI (https://officeimo.net/equations", Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath)), StringComparison.Ordinal);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(File.ReadAllBytes(pdfPath));
        Assert.Equal(3, info.LinkAnnotations.Count(link => link.Uri == "https://officeimo.net/equations"));
        Assert.Equal(3, info.LinkAnnotations.Count(link => link.Uri == "https://officeimo.net/nested-equations"));
        Assert.Equal(3, info.LinkAnnotations.Count(link => link.Uri == "https://officeimo.net/table-equations"));
        Assert.Single(info.LinkAnnotations, link => link.Uri == "https://officeimo.net/table-secondary");
        string normalizedLineBreaks = text.Replace("\r\n", "\n").Replace('\r', '\n');
        Assert.Contains("break-equation\nbreak-suffix", normalizedLineBreaks, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_MapsSimpleAndComplexEqFieldsToStaticText() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeEqFields.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeEqFields.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph simple = document.AddParagraph("Simple field: ");
            AppendSimpleField(simple._paragraph, " EQ \\f(a,b) ", "(a)/(b)");
            simple.AddText(" simple-suffix");

            WordParagraph complex = document.AddParagraph("Complex field: ");
            AppendComplexField(complex._paragraph, " EQ \\r(,x) ", "sqrt(x)");
            complex.AddText(" complex-suffix");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyEquationUnsupported");
        string text = NormalizePdfText(PdfCore.PdfTextExtractor.ExtractAllText(pdfPath));
        Assert.Contains("Simple field: (a)/(b) simple-suffix", text, StringComparison.Ordinal);
        Assert.Contains("Complex field: sqrt(x) complex-suffix", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Records_Warnings_For_Unsupported_Body_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyWarnings.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyWarnings.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
            document.AddParagraph("Native body control text").AddDropDownList(new[] { "One", "Two" }, "BodyControl", "BodyControlTag");

            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "NativeTableControlText";
            table.Rows[0].Cells[0].Paragraphs[0].AddEquation(omml);

            document.AddEmbeddedFragment("<html><body><p>Embedded body fragment</p></body></html>", WordAlternativeFormatImportPartType.Html);
            document.Save();
            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            result.Save(pdfPath);

            Assert.DoesNotContain(result.Warnings, warning =>
                warning.Code == "NativeBodyContentControlUnsupported" &&
                warning.Source == "body paragraph");
            Assert.DoesNotContain(result.Warnings, warning =>
                warning.Code == "NativeBodyEquationUnsupported" &&
                warning.Source == "body table");
            Assert.Contains(result.Warnings, warning =>
                warning.Code == "NativeBodyEmbeddedDocumentUnsupported" &&
                warning.Source == "body");
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native body control text", text);
        Assert.Contains("NativeTableControlText", text);
        Assert.Contains("x=1", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Simple_Text_ContentControls() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleTextContentControls.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleTextContentControls.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header content control: ")
                .AddStructuredDocumentTag("Header control", "HeaderAlias", "HeaderTag");

            document.AddParagraph("Native body content control: ")
                .AddStructuredDocumentTag("Body control", "BodyAlias", "BodyTag");

            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Native cell content control: ";
            table.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag("Cell control", "CellAlias", "CellTag");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header content control:", text);
        Assert.Contains("Header control", text);
        Assert.Contains("Native body content control:", text);
        Assert.Contains("Body control", text);
        Assert.Contains("Native cell content", text);
        Assert.Contains("Cell control", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_RejectsExcessiveStructuredBlockNesting() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDeepStructuredBlocks.docx");
        using WordDocument document = WordDocument.Create(docPath);
        OpenXmlElement nested = new Paragraph(new Run(new Text("deep content")));
        for (int depth = 0; depth < 130; depth++) {
            nested = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Nested " + depth }),
                new SdtContentBlock(nested));
        }

        document._document.Body!.Append(nested);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.ToPdfDocument(new PdfSaveOptions { IncludePageNumbers = false }));
        Assert.Contains("nesting exceeds", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_DropDown_ComboBox_And_DatePicker_To_AcroForm_Fields() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyContentControlFormFields.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyContentControlFormFields.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Native dropdown: ").AddDropDownList(new[] { "Poland", "Germany" }, "Country", "CountryTag");
            document.AddParagraph("Native combo: ").AddComboBox(new[] { "Red", "Blue" }, "Color", "ColorTag", defaultValue: "Blue");
            document.AddParagraph("Native date: ").AddDatePicker(new DateTime(2026, 5, 29), "Due Date", "DueDateTag");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body paragraph");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(3, info.FormFields.Count);

        PdfCore.PdfFormField country = Assert.Single(info.FormFields, field => field.Name == "CountryTag");
        Assert.Equal(PdfCore.PdfFormFieldKind.Choice, country.Kind);
        Assert.True(country.IsCombo);
        Assert.Equal("Poland", country.Value);
        Assert.Equal(new[] { "Poland", "Germany" }, country.Options.Select(option => option.ExportValue).ToArray());

        PdfCore.PdfFormField color = Assert.Single(info.FormFields, field => field.Name == "ColorTag");
        Assert.Equal(PdfCore.PdfFormFieldKind.Choice, color.Kind);
        Assert.True(color.IsCombo);
        Assert.Equal("Blue", color.Value);

        PdfCore.PdfFormField dueDate = Assert.Single(info.FormFields, field => field.Name == "DueDateTag");
        Assert.Equal(PdfCore.PdfFormFieldKind.Text, dueDate.Kind);
        Assert.Equal("2026-05-29", dueDate.Value);

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native dropdown:", text);
        Assert.Contains("Native combo:", text);
        Assert.Contains("Native date:", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_DropDown_ComboBox_And_DatePicker_To_AcroForm_Fields() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellContentControlFormFields.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellContentControlFormFields.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 7200;
            table.ColumnWidth = new[] { 7200 }.ToList();
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Native table controls:";
            paragraph.AddDropDownList(new[] { "Poland", "Germany" }, "Cell Country", "CellCountry");
            paragraph.AddComboBox(new[] { "Red", "Blue" }, "Cell Color", "CellColor", defaultValue: "Blue");
            paragraph.AddDatePicker(new DateTime(2026, 5, 31), "Cell Due Date", "CellDueDate");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(3, info.FormFields.Count);
        Assert.Contains(info.FormFields, field => field.Name == "CellCountry" && field.IsChoiceField && field.Value == "Poland");
        Assert.Contains(info.FormFields, field => field.Name == "CellColor" && field.IsChoiceField && field.Value == "Blue");
        Assert.Contains(info.FormFields, field => field.Name == "CellDueDate" && field.IsTextField && field.Value == "2026-05-31");
        Assert.True(info.Pages[0].HasFormWidgets);

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Native table controls:", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_RepeatingSection_To_Text_Items() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyRepeatingSection.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyRepeatingSection.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Native repeating section:");
            WordRepeatingSection repeatingSection = document.AddParagraph()
                .AddRepeatingSection("Tasks", "Tasks", "TasksTag");
            repeatingSection.SetTextItems(new[] { "Plan roadmap slice", "Validate native PDF output" });
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body paragraph");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native repeating section:", text);
        Assert.Contains("Plan roadmap slice", text);
        Assert.Contains("Validate native PDF output", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_RepeatingSection_To_Text_Items() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRepeatingSection.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRepeatingSection.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 7200;
            table.ColumnWidth = new[] { 7200 }.ToList();
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Table tasks";
            WordRepeatingSection repeatingSection = table.Rows[0].Cells[0].Paragraphs[0]
                .AddRepeatingSection("Tasks", "Tasks", "TasksTag");
            repeatingSection.SetTextItems(new[] { "Render cell item", "Keep table warnings clean" });
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Table tasks", text);
        Assert.Contains("Render cell item", text);
        Assert.Contains("Keep table warnings clean", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_CheckBox_To_AcroForm_Field() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyCheckBox.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyCheckBox.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Accept native checkbox").AddCheckBox(true, "Accept Native", "AcceptNative");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("AcceptNative", field.Name);
        Assert.Equal(PdfCore.PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Accept native checkbox", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_CheckBox_To_AcroForm_Field() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellCheckBox.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellCheckBox.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Table cell approval";
            table.Rows[0].Cells[0].Paragraphs[0].AddCheckBox(true, "Table Cell Approval", "TableCellApproval");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("TableCellApproval", field.Name);
        Assert.Equal(PdfCore.PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);
        Assert.True(info.Pages[0].HasFormWidgets);

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Table cell approval", pdf.GetPage(1).Text);
    }
}
