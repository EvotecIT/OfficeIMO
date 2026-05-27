using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLogicalDocumentTests {
    [Fact]
    public void Load_BuildsLogicalPagesWithTextTablesAndImages() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(p => p.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(CreateMinimalRgbPng(), 18, 18)
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal("Logical sample", logical.Metadata.Title);
        PdfLogicalHeading heading = Assert.Single(page.Headings);
        Assert.Equal("Logical Heading", heading.Text);
        Assert.Equal(1, heading.Level);
        Assert.Equal(PdfLogicalElementKind.Heading, heading.Line.Kind);
        Assert.Same(heading, Assert.Single(logical.Headings));
        Assert.Contains(page.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(page.TextBlocks, block =>
            block.Kind == PdfLogicalElementKind.ListItem &&
            Normalize(block.Text).Contains("Detectedlogicalbullet", StringComparison.Ordinal));
        PdfLogicalListItem listItem = Assert.Single(page.ListItems);
        Assert.Equal("Detected logical bullet", listItem.Text);
        Assert.Equal(1, listItem.Level);
        Assert.NotEmpty(listItem.Marker);
        Assert.Equal(PdfLogicalElementKind.ListItem, listItem.Line.Kind);
        Assert.Same(listItem, Assert.Single(logical.ListItems));
        Assert.Contains(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("A-100", StringComparison.Ordinal));

        PdfLogicalTable table = Assert.Single(page.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3);
        Assert.Same(table, Assert.Single(logical.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3));
        Assert.Contains(table.Rows, row => row.Count >= 3 &&
            Normalize(row[0]) == "A-100" &&
            Normalize(row[1]) == "Alpha" &&
            Normalize(row[2]) == "2");
        Assert.Contains(table.Cells, cell =>
            cell.PageNumber == 1 &&
            cell.RowIndex == 1 &&
            cell.ColumnIndex == 0 &&
            Normalize(cell.Text) == "A-100" &&
            cell.Column is not null &&
            cell.Column.From < cell.Column.To);
        Assert.Contains(table.Cells, cell =>
            cell.RowIndex == 2 &&
            cell.ColumnIndex == 2 &&
            Normalize(cell.Text) == "14");

        PdfLogicalImage image = Assert.Single(page.Images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("image/png", image.MimeType);
        Assert.Same(image, Assert.Single(logical.Images));

        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Table);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Image);
    }

    [Fact]
    public void Load_GroupsWrappedTextLinesIntoParagraphs() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("This logical paragraph should wrap across multiple nearby PDF text lines so wrappers can start from paragraph-like objects."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "P-100", "Paragraph table text", "2" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 50, 100, 30 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfLogicalPage page = Assert.Single(PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages);

        PdfLogicalParagraph paragraph = Assert.Single(page.Paragraphs, item => item.Text.Contains("logical paragraph", StringComparison.Ordinal));
        Assert.True(paragraph.Lines.Count > 1);
        Assert.Contains("logical paragraph", paragraph.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("P-100", paragraph.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Load_ExposesSimpleAcroFormFields() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildHierarchicalFormPdf());

        Assert.True(logical.HasFormFields);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms", "Selection.Country" }, logical.FormFields.Select(field => field.Name).ToArray());
        Assert.Equal("OfficeIMO", logical.FormFields[0].Value);
        Assert.Equal("InheritedDraft", logical.FormFields[0].DefaultValue);
        Assert.Equal(64, logical.FormFields[0].MaxLength);
        Assert.True(logical.FormFields[0].IsReadOnly);
        Assert.Equal("Yes", logical.FormFields[1].Value);
        Assert.Equal("DE", logical.FormFields[2].Value);
        Assert.Equal("PL", logical.FormFields[2].DefaultValue);
        Assert.Equal(2, logical.FormFields[2].OptionCount);
        Assert.Equal(new[] { "DE" }, logical.FormFields[2].SelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.Equal(new[] { "PL" }, logical.FormFields[2].DefaultSelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.Equal(3, logical.FormFieldsByName.Count);
        Assert.Contains("Person.Name", logical.FormFieldNames);
        Assert.Contains("AcceptTerms", logical.FormFieldNames);
        Assert.Contains("Selection.Country", logical.FormFieldNames);

        Assert.True(logical.TryGetFormField("Person.Name", out PdfFormField? nameField));
        Assert.Equal("OfficeIMO", nameField!.Value);
        Assert.Equal(new[] { "InheritedDraft" }, nameField.DefaultValues);
        Assert.True(logical.TryGetFormField("AcceptTerms", out PdfFormField? acceptField));
        Assert.Equal("Yes", acceptField!.Value);
        Assert.True(logical.TryGetFormField("Selection.Country", out PdfFormField? countryField));
        Assert.True(countryField!.IsChoiceField);
        Assert.False(logical.TryGetFormField("Missing", out PdfFormField? missingField));
        Assert.Null(missingField);
    }

    [Fact]
    public void Load_ExposesAcroFormFieldKindsAndFlags() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildFieldKindFormPdf());

        PdfFormField text = logical.FormFieldsByName["Notes"];
        Assert.Equal(PdfFormFieldKind.Text, text.Kind);
        Assert.True(text.IsTextField);
        Assert.True(text.IsReadOnly);
        Assert.True(text.IsRequired);
        Assert.True(text.IsNoExport);
        Assert.True(text.IsMultiline);
        Assert.True(text.IsPassword);
        Assert.Equal(42, text.MaxLength);
        Assert.Equal(new[] { "Secret" }, text.Values);
        Assert.Equal("Draft", text.DefaultValue);
        Assert.Equal(new[] { "Draft" }, text.DefaultValues);
        Assert.True(text.HasDefaultValues);
        Assert.False(text.HasOptions);
        Assert.False(text.HasDefaultSelectedOptions);
        Assert.False(text.IsButtonField);
        Assert.False(text.IsChoiceField);

        PdfFormField checkBox = logical.FormFieldsByName["Accept"];
        Assert.Equal(PdfFormFieldKind.Button, checkBox.Kind);
        Assert.True(checkBox.IsButtonField);
        Assert.True(checkBox.IsCheckBox);
        Assert.False(checkBox.IsRadioButton);
        Assert.False(checkBox.IsPushButton);

        PdfFormField radio = logical.FormFieldsByName["Choice"];
        Assert.True(radio.IsRadioButton);
        Assert.True(radio.IsNoToggleToOff);
        Assert.False(radio.IsCheckBox);

        PdfFormField pushButton = logical.FormFieldsByName["Submit"];
        Assert.True(pushButton.IsPushButton);
        Assert.False(pushButton.IsCheckBox);

        PdfFormField choice = logical.FormFieldsByName["Country"];
        Assert.Equal(PdfFormFieldKind.Choice, choice.Kind);
        Assert.Equal("[PL US]", choice.Value);
        Assert.Equal(new[] { "PL", "US" }, choice.Values);
        Assert.Equal("[DE US]", choice.DefaultValue);
        Assert.Equal(new[] { "DE", "US" }, choice.DefaultValues);
        Assert.True(choice.IsChoiceField);
        Assert.True(choice.IsCombo);
        Assert.True(choice.IsEditableChoice);
        Assert.True(choice.IsSortedChoice);
        Assert.True(choice.AllowsMultipleSelection);
        Assert.True(choice.DoesNotSpellCheck);
        Assert.True(choice.CommitsOnSelectionChange);
        Assert.True(choice.HasOptions);
        Assert.Equal(3, choice.OptionCount);
        Assert.Equal("PL", choice.Options[0].ExportValue);
        Assert.Equal("Poland", choice.Options[0].DisplayText);
        Assert.True(choice.Options[0].HasSeparateDisplayText);
        Assert.Equal("DE", choice.Options[1].ExportValue);
        Assert.Equal("DE", choice.Options[1].DisplayText);
        Assert.False(choice.Options[1].HasSeparateDisplayText);
        Assert.Equal("US", choice.Options[2].ExportValue);
        Assert.Equal("United States", choice.Options[2].DisplayText);
        Assert.True(choice.HasSelectedOptions);
        Assert.Equal(2, choice.SelectedOptionCount);
        Assert.Equal(new[] { "PL", "US" }, choice.SelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.True(choice.HasDefaultSelectedOptions);
        Assert.Equal(2, choice.DefaultSelectedOptionCount);
        Assert.Equal(new[] { "DE", "US" }, choice.DefaultSelectedOptions.Select(option => option.ExportValue).ToArray());

        PdfFormField signature = logical.FormFieldsByName["Approval"];
        Assert.Equal(PdfFormFieldKind.Signature, signature.Kind);
        Assert.True(signature.IsSignatureField);

        Assert.Same(text, Assert.Single(logical.GetFormFields(PdfFormFieldKind.Text)));
        Assert.Equal(new[] { "Accept", "Choice", "Submit" }, logical.GetFormFields(PdfFormFieldKind.Button).Select(field => field.Name).ToArray());
        Assert.Same(choice, Assert.Single(logical.FormFieldsByKind[PdfFormFieldKind.Choice]));
        Assert.Same(signature, Assert.Single(logical.GetFormFields(PdfFormFieldKind.Signature)));
        Assert.Empty(logical.GetFormFields(PdfFormFieldKind.Unknown));
    }

    [Fact]
    public void Load_ExposesAcroFormWidgetGeometry() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildWidgetFormPdf());

        PdfFormField field = Assert.Single(logical.FormFields);
        Assert.Equal("AcceptTerms", field.Name);
        Assert.Equal("Btn", field.FieldType);
        Assert.Equal("Yes", field.Value);
        Assert.True(field.HasWidgets);

        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(8, widget.ObjectNumber);
        Assert.Equal(1, widget.PageNumber);
        Assert.Equal(20, widget.X1);
        Assert.Equal(100, widget.Y1);
        Assert.Equal(36, widget.X2);
        Assert.Equal(116, widget.Y2);
        Assert.Equal(16, widget.Width);
        Assert.Equal(16, widget.Height);
        Assert.Equal("Yes", widget.AppearanceState);
        Assert.Equal(4, widget.Flags);

        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfLogicalFormWidget logicalWidget = Assert.Single(page.FormWidgets);
        Assert.Same(field, logicalWidget.Field);
        Assert.Same(widget, logicalWidget.SourceWidget);
        Assert.Equal(PdfLogicalElementKind.FormWidget, logicalWidget.Kind);
        Assert.Equal("AcceptTerms", logicalWidget.FieldName);
        Assert.Equal("Btn", logicalWidget.FieldType);
        Assert.Equal("Yes", logicalWidget.Value);
        Assert.Equal(8, logicalWidget.ObjectNumber);
        Assert.Equal(1, logicalWidget.PageNumber);
        Assert.Equal(20, logicalWidget.X1);
        Assert.Equal(100, logicalWidget.Y1);
        Assert.Equal(36, logicalWidget.X2);
        Assert.Equal(116, logicalWidget.Y2);
        Assert.True(logical.HasFormWidgets);
        Assert.Same(logicalWidget, Assert.Single(logical.FormWidgets));
        Assert.Same(logicalWidget, Assert.Single(logical.FormWidgetsByFieldName["AcceptTerms"]));
        Assert.Same(logicalWidget, Assert.Single(logical.GetFormWidgets("AcceptTerms")));
        Assert.Empty(logical.GetFormWidgets("Missing"));
        Assert.Contains(page.Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);
    }

    [Fact]
    public void Load_ExposesDocumentNavigationObjects() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildNavigationPdf());

        Assert.Equal("FullScreen", logical.CatalogPageMode);
        Assert.Equal("TwoColumnLeft", logical.CatalogPageLayout);
        Assert.Equal("1.7", logical.CatalogVersion);
        Assert.Equal("en-US", logical.CatalogLanguage);

        Assert.True(logical.HasOutlines);
        PdfOutlineItem outline = Assert.Single(logical.Outlines);
        Assert.Equal("Logical outline", outline.Title);
        Assert.Equal(1, outline.PageNumber);

        Assert.True(logical.HasReadablePageLabels);
        PdfPageLabel label = Assert.Single(logical.PageLabels);
        Assert.Equal(0, label.StartPageIndex);
        Assert.Equal("D", label.Style);
        Assert.Equal("A-", label.Prefix);
        Assert.Equal(3, label.StartNumber);

        Assert.True(logical.HasNamedDestinations);
        PdfNamedDestination destination = Assert.Single(logical.NamedDestinations);
        Assert.Equal("Chapter1", destination.Name);
        Assert.Equal(1, destination.PageNumber);

        Assert.True(logical.HasReadableOpenAction);
        Assert.NotNull(logical.OpenAction);
        Assert.Equal("Destination", logical.OpenAction!.ActionType);
        Assert.Equal(1, logical.OpenAction.PageNumber);

        Assert.True(logical.HasReadableViewerPreferences);
        Assert.NotNull(logical.ViewerPreferences);
        Assert.True(logical.ViewerPreferences!.GetBoolean("HideToolbar"));
        Assert.True(logical.ViewerPreferences.GetBoolean("DisplayDocTitle"));
    }

    [Fact]
    public void Load_ReadsStreamFromCurrentPosition() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.Text("Logical stream marker."))
            .ToBytes();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        PdfLogicalDocument logical = PdfLogicalDocument.Load(stream);

        Assert.Single(logical.Pages);
        Assert.Contains(logical.Pages[0].TextBlocks, block => block.Text.Contains("Logical stream marker", StringComparison.Ordinal));
    }

    [Fact]
    public void Load_ExposesLinkAnnotationsAsLogicalElements() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Linked heading", linkUri: "https://evotec.xyz/logical-link", linkContents: "Logical link metadata")
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        PdfLogicalPage page = Assert.Single(logical.Pages);

        PdfLogicalLinkAnnotation link = Assert.Single(page.Links);
        Assert.Equal(1, link.PageNumber);
        Assert.True(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.Equal("https://evotec.xyz/logical-link", link.Uri);
        Assert.Equal("Logical link metadata", link.Contents);
        Assert.True(link.Width > 0);
        Assert.True(link.Height > 0);
        Assert.Equal(1, link.SourceLink.PageNumber);
        Assert.True(logical.HasLinks);
        Assert.Same(link, Assert.Single(logical.Links));
        Assert.Same(link, Assert.Single(logical.LinksByUri["https://evotec.xyz/logical-link"]));
        Assert.Same(link, Assert.Single(logical.GetLinksByUri("https://evotec.xyz/logical-link")));
        Assert.Empty(logical.GetLinksByUri("https://evotec.xyz/missing"));
        Assert.Empty(logical.GetLinksByDestinationName("Missing"));
        Assert.Contains(page.Elements, element => element.Kind == PdfLogicalElementKind.LinkAnnotation);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.LinkAnnotation);
    }

    private static string Normalize(string text) {
        return new string(text.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
    }

    private static byte[] BuildHierarchicalFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R 8 0 R 9 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Person) /Ff 1 /MaxLen 64 /DV (InheritedDraft) /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /T (Name) /TU (Display name) /TM (ExportName) /V (OfficeIMO) >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes >>",
            "endobj",
            "9 0 obj",
            "<< /FT /Ch /T (Selection) /V /DE /DV (PL) /Opt [[(PL) (Poland)] (DE)] /Kids [10 0 R] >>",
            "endobj",
            "10 0 obj",
            "<< /T (Country) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFieldKindFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R 7 0 R 8 0 R 9 0 R 10 0 R 11 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Notes) /V (Secret) /DV (Draft) /Ff 12295 /MaxLen 42 >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Btn /T (Accept) /V /Yes >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (Choice) /V /A /Ff 49152 >>",
            "endobj",
            "9 0 obj",
            "<< /FT /Btn /T (Submit) /Ff 65536 >>",
            "endobj",
            "10 0 obj",
            "<< /FT /Ch /T (Country) /V [(PL) /US] /DV [(DE) /US] /Ff 74317826 /Opt [[(PL) (Poland)] (DE) [/US (United States)]] >>",
            "endobj",
            "11 0 obj",
            "<< /FT /Sig /T (Approval) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 12 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildWidgetFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 36 116] /F 4 /AS /Yes >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNavigationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageMode /FullScreen /PageLayout /TwoColumnLeft /Version /1.7 /Lang (en-US) /PageLabels 5 0 R /Dests 6 0 R /OpenAction [3 0 R /XYZ 0 200 0] /ViewerPreferences 7 0 R /Outlines 8 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Nums [0 << /S /D /P (A-) /St 3 >>] >>",
            "endobj",
            "6 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "7 0 obj",
            "<< /HideToolbar true /DisplayDocTitle true >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Outlines /First 9 0 R /Last 9 0 R /Count 1 >>",
            "endobj",
            "9 0 obj",
            "<< /Title (Logical outline) /Parent 8 0 R /Dest [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }
}
