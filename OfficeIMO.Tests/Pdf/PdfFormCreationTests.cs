using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormCreationTests {
    [Fact]
    public void TextField_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDocument.Create()
            .Meta(title: "Generated form")
            .Paragraph(p => p.Text("Generated field:"))
            .TextField("Person.Name", width: 180, height: 24, value: "Ada Lovelace", spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Contains("/FT /Tx", raw);
        Assert.Contains("/AP << /N", raw);
        Assert.True(info.HasReadableFormFields);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Person.Name", field.Name);
        Assert.Equal(PdfFormFieldKind.Text, field.Kind);
        Assert.Equal("Ada Lovelace", field.Value);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(1, widget.PageNumber);
        Assert.True(widget.Width > 170);
        Assert.True(widget.Height > 20);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void TextField_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDocument.Create()
            .TextField("Person.Name", width: 180, height: 24, value: "Original")
            .ToBytes();

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["Person.Name"] = "Created fill"
        });

        string raw = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.Contains("<437265617465642066696C6C> Tj", raw);
    }

    [Fact]
    public void CheckBox_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Generated checkbox:"))
            .CheckBox("AcceptTerms", isChecked: true, size: 16, spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/NeedAppearances false", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Contains("/FT /Btn", raw);
        Assert.Contains("/AS /Yes", raw);
        Assert.Contains("/AP << /N << /Off", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal(false, info.AcroFormNeedAppearances);
        Assert.Equal("AcceptTerms", field.Name);
        Assert.Equal(PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal("Yes", widget.AppearanceState);
        Assert.True(widget.HasNormalAppearanceState("Off"));
        Assert.True(widget.HasNormalAppearanceState("Yes"));
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void TableCellCheckBox_CreatesInspectableAcroFormFieldInsideCell() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] {
                    PdfTableCell.WithCheckBoxes(
                        "Table approval",
                        new[] { new PdfTableCellCheckBox("Table.Approved", isChecked: true, size: 12) })
                }
            }, style: TableStyles.Light())
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/NeedAppearances false", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Table.Approved", field.Name);
        Assert.Equal(PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(1, widget.PageNumber);
        Assert.True(widget.Width >= 11);
        Assert.True(widget.Height >= 11);
        Assert.True(info.Pages[0].HasFormWidgets);
    }

    [Fact]
    public void TableCellCheckBox_RendersInlineWithSingleLineText() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageSize = new PageSize(300, 180),
                Margins = PageMargins.Uniform(24)
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.WithCheckBoxes(
                        "Table approval",
                        new[] { new PdfTableCellCheckBox("Table.InlineApproved", isChecked: true, size: 12) })
                }
            }, style: TableStyles.Light())
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfFormWidget widget = Assert.Single(Assert.Single(info.FormFields).Widgets);
        using var pdfDocument = PdfPigDocument.Open(new MemoryStream(pdf));
        var line = FindLine(pdfDocument.GetPage(1), "Table approval");
        double lineEndX = line.Max(letter => letter.EndBaseLine.X);
        double baselineY = line[0].StartBaseLine.Y;
        double widgetCenterY = (widget.Y1 + widget.Y2) / 2D;

        Assert.True(widget.X1 > lineEndX);
        Assert.InRange(widget.X1 - lineEndX, 0D, 12D);
        Assert.InRange(widgetCenterY - baselineY, 0D, 6D);
    }

    [Fact]
    public void TableCellFormFields_CreateInspectableTextAndChoiceFieldsInsideCell() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] {
                    PdfTableCell.WithFormFields(
                        "Table form fields",
                        new[] {
                            PdfTableCellFormField.TextField("Table.DueDate", "2026-05-31", width: 140, height: 18),
                            PdfTableCellFormField.ChoiceField("Table.Country", new[] { "Poland", "Germany" }, value: "Germany", width: 140, height: 18)
                        })
                }
            }, style: TableStyles.Light())
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Equal(2, info.FormFields.Count);
        PdfFormField dueDate = Assert.Single(info.FormFields, field => field.Name == "Table.DueDate");
        Assert.Equal(PdfFormFieldKind.Text, dueDate.Kind);
        Assert.Equal("2026-05-31", dueDate.Value);
        Assert.Equal(1, Assert.Single(dueDate.Widgets).PageNumber);

        PdfFormField country = Assert.Single(info.FormFields, field => field.Name == "Table.Country");
        Assert.Equal(PdfFormFieldKind.Choice, country.Kind);
        Assert.True(country.IsChoiceField);
        Assert.True(country.IsCombo);
        Assert.Equal("Germany", country.Value);
        Assert.Equal(new[] { "Poland", "Germany" }, country.Options.Select(option => option.ExportValue).ToArray());
        Assert.True(info.Pages[0].HasFormWidgets);
    }

    [Fact]
    public void CheckBox_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDocument.Create()
            .CheckBox("AcceptTerms")
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });
        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.Equal("Yes", Assert.Single(filledInfo.FormFields).Value);
        Assert.Equal("Yes", Assert.Single(filledInfo.FormWidgets).AppearanceState);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.DoesNotContain("/Subtype /Widget", raw);
        Assert.Contains("/OfficeIMOForm1 Do", raw);
    }

    [Fact]
    public void RadioButtonGroup_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Generated radio buttons:"))
            .RadioButtonGroup("Payment.Method", new[] { "Card", "Cash", "Wire" }, value: "Cash", size: 16, gap: 5, spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/FT /Btn", raw);
        Assert.Contains("/Ff 49152", raw);
        Assert.Contains("/Kids [", raw);
        Assert.Contains("/V /Cash", raw);
        using var pdfDocument = PdfPigDocument.Open(new MemoryStream(pdf));
        string pageText = pdfDocument.GetPage(1).Text;
        Assert.Contains("Card", pageText);
        Assert.Contains("Cash", pageText);
        Assert.Contains("Wire", pageText);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Payment.Method", field.Name);
        Assert.Equal(PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsRadioButton);
        Assert.True(field.IsNoToggleToOff);
        Assert.Equal("Cash", field.Value);
        Assert.Equal(3, field.WidgetCount);
        Assert.Contains(field.Widgets, widget => widget.AppearanceState == "Cash" && widget.HasNormalAppearanceState("Cash"));
        Assert.Equal(2, field.Widgets.Count(widget => widget.AppearanceState == "Off"));
        Assert.All(field.Widgets, widget => {
            Assert.Equal(1, widget.PageNumber);
            Assert.True(widget.HasNormalAppearanceState("Off"));
        });
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void RadioButtonGroup_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDocument.Create()
            .RadioButtonGroup("Payment.Method", new[] { "Card", "Cash", "Wire" }, value: "Card")
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Payment.Method"] = "Wire"
        });
        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Wire", filledField.Value);
        Assert.Contains(filledField.Widgets, widget => widget.AppearanceState == "Wire");
        Assert.Equal(2, filledField.Widgets.Count(widget => widget.AppearanceState == "Off"));

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["Payment.Method"] = "Wire"
        });
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.DoesNotContain("/Subtype /Widget", raw);
        Assert.Contains("/OfficeIMOForm", raw);
    }

    [Fact]
    public void RadioButtonGroup_RejectsUnknownFillValue() {
        byte[] pdf = PdfDocument.Create()
            .RadioButtonGroup("Payment.Method", new[] { "Card", "Cash", "Wire" }, value: "Card")
            .ToBytes();

        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Payment.Method"] = "WireTransfer"
        }));

        Assert.Contains("not one of the available appearance states", exception.Message);
    }

    [Fact]
    public void GeneratedFields_CanStyleAppearances() {
        var style = new PdfFormFieldStyle {
            BackgroundColor = PdfColor.FromRgb(238, 242, 255),
            BorderColor = PdfColor.FromRgb(30, 64, 175),
            BorderWidth = 2,
            TextColor = PdfColor.FromRgb(127, 29, 29),
            MarkColor = PdfColor.FromRgb(22, 101, 52)
        };

        byte[] pdf = PdfDocument.Create()
            .TextField("Styled.Name", value: "Ada", style: style)
            .CheckBox("Styled.Accept", isChecked: true, style: style)
            .ChoiceField("Styled.Country", new[] { "Poland", "Germany" }, value: "Poland", style: style)
            .RadioButtonGroup("Styled.Contact", new[] { "Email", "Phone" }, value: "Phone", style: style)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/BC [0.118 0.251 0.686] /BG [0.933 0.949 1]", raw, StringComparison.Ordinal);
        Assert.Contains("/Helv 10 Tf 0.498 0.114 0.114 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.118 0.251 0.686 RG 2 w", raw, StringComparison.Ordinal);
        Assert.Contains("0.086 0.396 0.204 RG 1.25 w", raw, StringComparison.Ordinal);
        Assert.Contains("0.086 0.396 0.204 rg", raw, StringComparison.Ordinal);
        Assert.Equal(4, PdfInspector.Inspect(pdf).FormFields.Count);

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Styled.Name"] = "Filled"
        });
        string filledRaw = Encoding.ASCII.GetString(filled);

        Assert.Contains("<46696C6C6564> Tj", filledRaw, StringComparison.Ordinal);
        Assert.Contains("0.118 0.251 0.686 RG 1 w", filledRaw, StringComparison.Ordinal);
        Assert.Contains("0.498 0.114 0.114 rg", filledRaw, StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedFields_UseEmbeddedHelveticaAppearanceResourceWhenDocumentFontIsEmbedded() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Form Body Font", fontPath)
            .Paragraph(paragraph => paragraph.Text("Embedded body font"))
            .TextField("Styled.Name", value: "Ada")
            .ChoiceField("Styled.Country", new[] { "Poland", "Germany" }, value: "Poland")
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/AcroForm", raw, StringComparison.Ordinal);
        Assert.Contains("/DA (/Helv 10 Tf", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("<416461> Tj", raw, StringComparison.Ordinal);
        Assert.Matches(@"/DR << /Font << /Helv \d+ 0 R >> >>", raw);
    }

    [Fact]
    public void GeneratedFields_EmitAccessibleMetadata() {
        var style = new PdfFormFieldStyle {
            AlternateName = "Accessible field",
            MappingName = "accessible.field"
        };

        byte[] pdf = PdfDocument.Create()
            .TextField("Accessible.Name", value: "Ada", style: style)
            .CheckBox("Accessible.Accept", isChecked: true, style: style)
            .ChoiceField("Accessible.Country", new[] { "Poland", "Germany" }, value: "Poland", style: style)
            .RadioButtonGroup("Accessible.Contact", new[] { "Email", "Phone" }, value: "Phone", style: style)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.Equal(4, CountOccurrences(raw, "/TU <41636365737369626C65206669656C64>"));
        Assert.Equal(4, CountOccurrences(raw, "/TM <61636365737369626C652E6669656C64>"));
        Assert.All(info.FormFields, field => Assert.Equal("Accessible field", field.AlternateName));
        Assert.All(info.FormFields, field => Assert.Equal("accessible.field", field.MappingName));

        PdfFormFieldStyle clone = style.Clone();
        style.AlternateName = "Changed";
        style.MappingName = "changed";
        Assert.Equal("Accessible field", clone.AlternateName);
        Assert.Equal("accessible.field", clone.MappingName);
    }

    [Fact]
    public void ChoiceField_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Generated choice:"))
            .ChoiceField("Country", new[] { "Poland", "Germany", "United States" }, value: "Germany", width: 180, height: 24, spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Contains("/FT /Ch", raw);
        Assert.Contains("/Opt [ <506F6C616E64> <4765726D616E79> <556E6974656420537461746573> ]", raw);
        Assert.Contains("/Ff 131072", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Country", field.Name);
        Assert.Equal(PdfFormFieldKind.Choice, field.Kind);
        Assert.True(field.IsCombo);
        Assert.Equal("Germany", field.Value);
        Assert.Equal(new[] { "Poland", "Germany", "United States" }, field.Options.Select(option => option.DisplayText).ToArray());
        Assert.Equal("Germany", Assert.Single(field.SelectedOptions).DisplayText);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(1, widget.PageNumber);
        Assert.True(widget.Width > 170);
        Assert.True(widget.Height > 20);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void ChoiceField_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDocument.Create()
            .ChoiceField("Country", new[] { "Poland", "Germany", "United States" }, value: "Poland")
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Country"] = "United States"
        });
        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);

        Assert.Equal("United States", filledField.Value);
        Assert.Equal("United States", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("<556E6974656420537461746573> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["Country"] = "United States"
        });
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.DoesNotContain("/Subtype /Widget", raw);
        Assert.Contains("/OfficeIMOForm1 Do", raw);
    }

    [Fact]
    public void MultiSelectChoiceField_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Generated multi-select choice:"))
            .MultiSelectChoiceField("Countries", new[] { "Poland", "Germany", "United States" }, values: new[] { "Poland", "United States" }, width: 190, height: 72)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/FT /Ch", raw);
        Assert.Contains("/V [<506F6C616E64> <556E6974656420537461746573>]", raw);
        Assert.Contains("/DV [<506F6C616E64> <556E6974656420537461746573>]", raw);
        Assert.Contains("/Ff 2097152", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Countries", field.Name);
        Assert.Equal(PdfFormFieldKind.Choice, field.Kind);
        Assert.False(field.IsCombo);
        Assert.True(field.AllowsMultipleSelection);
        Assert.Equal(new[] { "Poland", "United States" }, field.Values);
        Assert.Equal(new[] { "Poland", "United States" }, field.SelectedOptions.Select(option => option.DisplayText).ToArray());
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.True(widget.Width > 180);
        Assert.True(widget.Height > 70);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void MultiSelectChoiceField_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDocument.Create()
            .MultiSelectChoiceField("Countries", new[] { "Poland", "Germany", "United States" }, values: new[] { "Poland" })
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, PdfFormFieldValue> {
            ["Countries"] = PdfFormFieldValue.FromValues("Germany", "United States")
        });
        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal(new[] { "Germany", "United States" }, filledField.Values);
        Assert.Equal(new[] { "Germany", "United States" }, filledField.SelectedOptions.Select(option => option.DisplayText).ToArray());
        Assert.Contains("<4765726D616E792C20556E6974656420537461746573> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);

        byte[] scalarFilled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Countries"] = "Germany"
        });
        string scalarRaw = Encoding.ASCII.GetString(scalarFilled);
        PdfFormField scalarField = Assert.Single(PdfInspector.Inspect(scalarFilled).FormFields);

        Assert.Equal(new[] { "Germany" }, scalarField.Values);
        Assert.Contains("/V [", scalarRaw, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79>", scalarRaw, StringComparison.Ordinal);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, PdfFormFieldValue> {
            ["Countries"] = PdfFormFieldValue.FromValues("Germany", "United States")
        });
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.DoesNotContain("/Subtype /Widget", raw);
        Assert.Contains("/OfficeIMOForm1 Do", raw);
    }

    [Fact]
    public void MultiSelectChoiceField_WithSingleSelectedValueStoresArrayValue() {
        byte[] pdf = PdfDocument.Create()
            .MultiSelectChoiceField("Countries", new[] { "Poland", "Germany" }, values: new[] { "Poland" })
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(pdf).FormFields);

        Assert.Equal(new[] { "Poland" }, field.Values);
        Assert.Contains("/V [<506F6C616E64>]", raw, StringComparison.Ordinal);
        Assert.Contains("/DV [<506F6C616E64>]", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposeRowsAndItems_CanPlaceGeneratedFormFields() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Compose(document => document.Page(page => page.Content(content => {
                content.Item(item => item
                    .TextField("Item.Name", width: 120, height: 20, value: "Ada", spacingAfter: 8)
                    .Element(element => element.CheckBox("Element.Accept", isChecked: true, size: 14, spacingAfter: 10)));

                content.Row(row => row
                    .Gap(24)
                    .Column(50, column => column
                        .Paragraph(p => p.Text("Left column"))
                        .TextField("Left.Email", width: 120, height: 20, value: "left@example.com", spacingAfter: 8)
                        .ChoiceField("Left.Country", new[] { "Poland", "Germany" }, value: "Poland", width: 120, height: 20))
                    .Column(50, column => column
                        .Paragraph(p => p.Text("Right column"))
                        .CheckBox("Right.Enabled", isChecked: true, size: 14, align: PdfAlign.Center, spacingAfter: 8)
                        .MultiSelectChoiceField("Right.Countries", new[] { "Poland", "Germany", "United States" }, values: new[] { "Germany" }, width: 120, height: 44)
                        .RadioButtonGroup("Right.Contact", new[] { "Email", "Phone" }, value: "Phone", size: 12, gap: 4)));
            })))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.Equal(7, info.FormFields.Count);
        Assert.Contains(info.FormFields, field => field.Name == "Item.Name" && field.IsTextField && field.Value == "Ada");
        Assert.Contains(info.FormFields, field => field.Name == "Element.Accept" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, field => field.Name == "Left.Email" && field.IsTextField && field.Value == "left@example.com");
        Assert.Contains(info.FormFields, field => field.Name == "Left.Country" && field.IsChoiceField && field.Value == "Poland");
        Assert.Contains(info.FormFields, field => field.Name == "Right.Enabled" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, field => field.Name == "Right.Countries" && field.IsChoiceField && field.AllowsMultipleSelection && field.Values.SequenceEqual(new[] { "Germany" }));
        Assert.Contains(info.FormFields, field => field.Name == "Right.Contact" && field.IsRadioButton && field.Value == "Phone");

        PdfFormWidget leftEmail = Assert.Single(info.GetFormWidgets("Left.Email"));
        PdfFormWidget rightEnabled = Assert.Single(info.GetFormWidgets("Right.Enabled"));
        PdfFormWidget rightCountries = Assert.Single(info.GetFormWidgets("Right.Countries"));

        Assert.Equal(1, leftEmail.PageNumber);
        Assert.Equal(1, rightEnabled.PageNumber);
        Assert.True(leftEmail.X1 < rightEnabled.X1);
        Assert.True(rightCountries.X1 > leftEmail.X1);
        Assert.True(leftEmail.Width > 110);
        Assert.True(rightCountries.Height > 40);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, PdfFormFieldValue> {
            ["Left.Email"] = PdfFormFieldValue.From("filled@example.com"),
            ["Right.Countries"] = PdfFormFieldValue.FromValues("Poland", "United States")
        });

        PdfDocumentInfo flattenedInfo = PdfInspector.Inspect(flattened);
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(flattenedInfo.HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.Contains("<66696C6C6564406578616D706C652E636F6D> Tj", raw);
        Assert.Contains("/OfficeIMOForm", raw);
    }

    [Fact]
    public void GeneratedFields_ValidateFlowGeometry() {
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().TextField(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().TextField("Name", width: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().TextField("Name", height: -1));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().TextField("Name", align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().CheckBox(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().CheckBox("AcceptTerms", size: 0));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().CheckBox("AcceptTerms", align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().CheckBox("AcceptTerms", checkedValueName: " "));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().CheckBox("AcceptTerms", checkedValueName: "Off"));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().ChoiceField(" ", new[] { "One" }));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().ChoiceField("Country", null!));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().ChoiceField("Country", Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().ChoiceField("Country", new[] { "One", "One" }));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().ChoiceField("Country", new[] { "One", " " }));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().ChoiceField("Country", new[] { "One" }, value: "Two"));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().ChoiceField("Country", new[] { "One" }, width: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().ChoiceField("Country", new[] { "One" }, height: -1));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().ChoiceField("Country", new[] { "One" }, align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().MultiSelectChoiceField("Countries", Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().MultiSelectChoiceField("Countries", new[] { "One" }, values: Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().MultiSelectChoiceField("Countries", new[] { "One" }, values: new[] { "Two" }));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().MultiSelectChoiceField("Countries", new[] { "One" }, values: new[] { "One", "One" }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().MultiSelectChoiceField("Countries", new[] { "One" }, height: 0));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup(" ", new[] { "One" }));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().RadioButtonGroup("Group", null!));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One", "One" }));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One", " " }));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One", "Off" }));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One" }, value: "Two"));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "Y\u2713" }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One" }, size: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One" }, gap: -1));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().RadioButtonGroup("Group", new[] { "One" }, align: PdfAlign.Justify));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfFormFieldStyle { BorderWidth = -1 });
        Assert.Throws<ArgumentException>(() => new PdfFormFieldStyle { AlternateName = " " });
        Assert.Throws<ArgumentException>(() => new PdfFormFieldStyle { MappingName = " " });

        Assert.Throws<ArgumentException>(() => PdfDocument.Create()
            .TextField("Email")
            .CheckBox("Email")
            .ToBytes());
    }

    private static List<UglyToad.PdfPig.Content.Letter> FindLine(UglyToad.PdfPig.Content.Page page, string expectedText) {
        foreach (var group in page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))) {
            var ordered = group.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            string normalizedText = text.Replace(" ", string.Empty);
            string normalizedExpected = expectedText.Replace(" ", string.Empty);
            if (normalizedText.Contains(normalizedExpected, StringComparison.Ordinal)) {
                return ordered;
            }
        }

        throw new InvalidOperationException("Could not find text line '" + expectedText + "' in rendered PDF.");
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }
}
