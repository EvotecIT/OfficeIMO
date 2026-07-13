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
    public void SaveAsPdf_OfficeIMOEngine_Maps_Simple_Equations_To_Static_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleEquations.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleEquations.pdf");
        const string headerOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>h=2</m:t></m:r></m:oMath></m:oMathPara>";
        const string bodyOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>b=3</m:t></m:r></m:oMath></m:oMathPara>";
        const string tableOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>c=4</m:t></m:r></m:oMath></m:oMathPara>";
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header equation:")
                .AddEquation(headerOmml);

            document.AddParagraph("Native body equation:").AddEquation(bodyOmml);

            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Native table equation:";
            table.Rows[0].Cells[0].Paragraphs[0].AddEquation(tableOmml);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterEquationUnsupported");
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyEquationUnsupported");

        string text = PdfCore.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native header equation:", text);
        Assert.Contains("h=2", text);
        Assert.Contains("Native body equation:", text);
        Assert.Contains("b=3", text);
        string normalizedText = NormalizePdfText(text);
        Assert.Contains("Native table equation:", normalizedText);
        Assert.Contains("c=4", normalizedText);
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
