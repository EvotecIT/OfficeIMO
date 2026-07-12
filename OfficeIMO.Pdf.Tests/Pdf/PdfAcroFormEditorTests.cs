using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAcroFormEditorTests {
    [Fact]
    public void Edit_AppliesFieldTreeWidgetOrderAndSelectiveFlattenTransaction() {
        byte[] source = PdfDocument.Create()
            .TextField("Person.Name", value: "Ada")
            .CheckBox("AcceptTerms", isChecked: true)
            .ChoiceField("Country", new[] { "Poland", "Germany" }, value: "Poland")
            .RadioButtonGroup("Payment.Method", new[] { "Card", "Cash" }, value: "Cash")
            .PageBreak()
            .Paragraph(p => p.Text("Second page"))
            .ToBytes();

        PdfAcroFormEditResult result = PdfAcroFormEditor.Edit(source, edit => edit
            .Rename("Person.Name", "Person.FullName")
            .SetDefaultValue("Person.FullName", "Unknown")
            .SetFlags("Person.FullName", 3)
            .Remove("AcceptTerms")
            .Move("Country", 2, 72, 540, 160, 24)
            .Create(new PdfFormFieldCreateOptions {
                Name = "Added.Note",
                Kind = PdfFormFieldCreationKind.Text,
                PageNumber = 2,
                X = 72,
                Y = 500,
                Width = 180,
                Height = 24,
                Value = "Created"
            })
            .PlaceSignatureField("Approval.Signature", 2, 72, 450, 180, 40)
            .SetCalculationOrder("Person.FullName", "Country", "Added.Note")
            .SetTabOrder(2, PdfPageTabOrder.Row)
            .Flatten("Payment.Method"));

        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());
        Assert.Equal(new[] { "Added.Note", "Approval.Signature", "Country", "Person.FullName" }, info.FormFieldNames.OrderBy(static name => name, StringComparer.Ordinal).ToArray());
        Assert.DoesNotContain("AcceptTerms", info.FormFieldNames);
        Assert.DoesNotContain("Payment.Method", info.FormFieldNames);
        PdfFormField renamed = info.FormFieldsByName["Person.FullName"];
        Assert.Equal("Ada", renamed.Value);
        Assert.Equal("Unknown", renamed.DefaultValue);
        Assert.Equal(3, renamed.Flags);
        PdfFormWidget moved = Assert.Single(info.FormFieldsByName["Country"].Widgets);
        Assert.Equal(2, moved.PageNumber);
        Assert.Equal(72D, moved.X1, 3);
        Assert.Equal(540D, moved.Y1, 3);
        Assert.Equal(232D, moved.X2, 3);
        Assert.Equal(564D, moved.Y2, 3);
        Assert.Equal("Created", info.FormFieldsByName["Added.Note"].Value);
        Assert.Equal(PdfFormFieldKind.Signature, info.FormFieldsByName["Approval.Signature"].Kind);
        Assert.Null(info.FormFieldsByName["Approval.Signature"].Value);
        Assert.Equal("R", info.Pages[1].TabOrder);
        Assert.Equal(new[] { "Person.FullName", "Country", "Added.Note" }, result.CalculationOrder);
        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Equal(10, result.Operations.Count);

        string raw = Encoding.ASCII.GetString(result.ToBytes());
        Assert.Contains("/FT /Sig", raw, StringComparison.Ordinal);
        Assert.Contains("/CO [", raw, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOForm", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void FlattenFields_ExactNamesPreserveUnselectedFields() {
        byte[] source = PdfDocument.Create()
            .TextField("Keep", value: "Visible")
            .CheckBox("Flatten", isChecked: true)
            .ToBytes();

        byte[] output = PdfFormFiller.FlattenFields(source, new[] { "Flatten" });
        PdfDocumentInfo info = PdfInspector.Inspect(output);

        PdfFormField keep = Assert.Single(info.FormFields);
        Assert.Equal("Keep", keep.Name);
        Assert.Equal("Visible", keep.Value);
        Assert.DoesNotContain("Flatten", info.FormFieldNames);
        Assert.Contains("/OfficeIMOForm", Encoding.ASCII.GetString(output), StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_CreatesAcroFormInDocumentWithoutFields() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("No form yet")).ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "Created.Name",
            Value = "Ada",
            X = 72,
            Y = 500,
            Width = 180,
            Height = 24
        }));

        PdfFormField field = Assert.Single(result.Fields);
        Assert.Equal("Created.Name", field.Name);
        Assert.Equal("Ada", field.Value);
        Assert.Equal(1, Assert.Single(field.Widgets).PageNumber);
    }

    [Fact]
    public void Edit_AllowsFurtherLayoutChangesToUnsignedSignatureField() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Signature page")).ToBytes();
        byte[] placed = PdfAcroFormEditor.Edit(source, edit => edit.PlaceSignatureField("Approval", 1, 72, 500, 180, 40)).ToBytes();

        PdfAcroFormEditResult moved = PdfAcroFormEditor.Edit(placed, edit => edit.Move("Approval", 1, 90, 440, 200, 50));

        PdfFormWidget widget = Assert.Single(Assert.Single(moved.Fields).Widgets);
        Assert.Equal(90D, widget.X1, 3);
        Assert.Equal(440D, widget.Y1, 3);
        Assert.Equal(290D, widget.X2, 3);
        Assert.Equal(490D, widget.Y2, 3);
    }

    [Fact]
    public void Edit_RejectsXfaWithoutChangingItsPackets() {
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
            "5 0 obj", "<< /Fields [] /XFA (unsupported-packet) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));

        var exception = Assert.Throws<NotSupportedException>(() => PdfAcroFormEditor.Edit(source, edit => edit.Create(new PdfFormFieldCreateOptions { Name = "Name" })));

        Assert.Contains("does not modify XFA packets", exception.Message, StringComparison.Ordinal);
        Assert.Contains("unsupported-packet", Encoding.ASCII.GetString(source), StringComparison.Ordinal);
    }
}
