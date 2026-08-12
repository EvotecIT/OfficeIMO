using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAcroFormEditorTests {
    [Fact]
    public void Edit_ValidatesOnlyFinalOrderingCommandsInOneTransaction() {
        byte[] source = PdfDocument.Create()
            .TextField("First", value: "1")
            .TextField("Second", value: "2")
            .ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .SetTabOrder(1, PdfPageTabOrder.Column)
            .SetCalculationOrder("First", "Second")
            .SetTabOrder(1, PdfPageTabOrder.Row)
            .SetCalculationOrder("Second", "First"));

        Assert.Equal("R", PdfInspector.Inspect(result.ToBytes()).Pages[0].TabOrder);
        Assert.Equal(new[] { "Second", "First" }, result.CalculationOrder);
    }

    [Fact]
    public void Edit_RetargetsDeferredFlatteningAndCalculationOrderThroughLaterTreeEdits() {
        byte[] source = PdfDocument.Create()
            .TextField("FlattenMe", value: "one")
            .TextField("RenameMe", value: "two")
            .TextField("RemoveMe", value: "three")
            .ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .SetCalculationOrder("RenameMe", "RemoveMe")
            .Flatten("FlattenMe")
            .Rename("FlattenMe", "FlattenedName")
            .Rename("RenameMe", "Renamed")
            .Remove("RemoveMe"));

        Assert.Equal(new[] { "Renamed" }, result.CalculationOrder);
        Assert.Equal(new[] { "Renamed" }, result.Fields.Select(static field => field.Name).ToArray());
    }

    [Fact]
    public void Edit_AllowsRecreatingOriginalNameAfterFlattenedFieldIsRenamed() {
        byte[] source = PdfDocument.Create().TextField("f", value: "old").ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Flatten("f")
            .Rename("f", "g")
            .Create(new PdfFormFieldCreateOptions { Name = "f", Value = "new" }));

        PdfFormField field = Assert.Single(result.Fields);
        Assert.Equal("f", field.Name);
        Assert.Equal("new", field.Value);
    }

    [Fact]
    public void Edit_AllowsReplacingAFieldAfterEarlierFlagsAndFlattenCommands() {
        byte[] source = PdfDocument.Create().TextField("f", value: "old").ToBytes();

        PdfAcroFormEditResult afterFlags = PdfDocument.Open(source).Forms.Edit(edit => edit
            .SetFlags("f", 1)
            .Remove("f")
            .Create(new PdfFormFieldCreateOptions { Name = "f", Value = "flags replacement" }));
        PdfAcroFormEditResult afterFlatten = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Flatten("f")
            .Remove("f")
            .Create(new PdfFormFieldCreateOptions { Name = "f", Value = "flatten replacement" }));

        Assert.Equal("flags replacement", Assert.Single(afterFlags.Fields).Value);
        Assert.Equal("flatten replacement", Assert.Single(afterFlatten.Fields).Value);
    }

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

        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result.ToBytes(), null).Map;
        PdfIndirectObject parentObject = Assert.Single(objects.Values, static item =>
            item.Value is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("T", out PdfObject? name) &&
            name is PdfStringObj text && text.Value == "Created");
        PdfDictionary parent = Assert.IsType<PdfDictionary>(parentObject.Value);
        PdfReference childReference = Assert.IsType<PdfReference>(Assert.Single(Assert.IsType<PdfArray>(parent.Items["Kids"]).Items));
        PdfDictionary child = Assert.IsType<PdfDictionary>(objects[childReference.ObjectNumber].Value);
        Assert.Equal("Name", Assert.IsType<PdfStringObj>(child.Items["T"]).Value);
        Assert.Equal(parentObject.ObjectNumber, Assert.IsType<PdfReference>(child.Items["Parent"]).ObjectNumber);
    }

    [Fact]
    public void Edit_RejectsTerminalNamesThatCollideWithNonterminalParents() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Field hierarchy collisions")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions { Name = "Customer.Notes", Value = "kept" })
            .Create(new PdfFormFieldCreateOptions { Name = "Other", Value = "rename source" }))
            .ToBytes();

        Assert.Throws<ArgumentException>(() => PdfDocument.Open(authored).Forms.Edit(edit => edit.Create(
            new PdfFormFieldCreateOptions { Name = "Customer", Value = "invalid" })));
        Assert.Throws<ArgumentException>(() => PdfDocument.Open(authored).Forms.Edit(edit => edit.Rename("Other", "Customer")));
    }

    [Fact]
    public void Edit_MoveRegeneratesAnEmptyChoiceAppearanceForTheNewBounds() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Empty choice move")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "Country",
            Kind = PdfFormFieldCreationKind.Choice,
            ChoiceOptions = new[] { "Poland", "Germany" },
            Width = 120D,
            Height = 24D
        })).ToBytes();

        byte[] moved = PdfDocument.Open(authored).Forms.Edit(edit => edit.Move("Country", 1, 72D, 440D, 200D, 40D)).ToBytes();
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(moved, null).Map;
        PdfDictionary widget = Assert.IsType<PdfDictionary>(Assert.Single(objects.Values, static item =>
            item.Value is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("T", out PdfObject? name) &&
            name is PdfStringObj text && text.Value == "Country").Value);
        PdfDictionary appearances = Assert.IsType<PdfDictionary>(widget.Items["AP"]);
        PdfReference normalReference = Assert.IsType<PdfReference>(appearances.Items["N"]);
        PdfStream normalAppearance = Assert.IsType<PdfStream>(objects[normalReference.ObjectNumber].Value);
        PdfArray boundingBox = Assert.IsType<PdfArray>(normalAppearance.Dictionary.Items["BBox"]);

        Assert.Equal(new[] { 0D, 0D, 200D, 40D }, boundingBox.Items.Cast<PdfNumber>().Select(static number => number.Value));
    }

    [Fact]
    public void Edit_MoveRegeneratesPushButtonCaptionAppearanceForTheNewBounds() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Push button move")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "Submit",
            Kind = PdfFormFieldCreationKind.PushButton,
            Caption = "Submit",
            Width = 120D,
            Height = 24D
        })).ToBytes();

        byte[] moved = PdfDocument.Open(authored).Forms.Edit(edit => edit.Move("Submit", 1, 72D, 440D, 200D, 40D)).ToBytes();
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(moved, null).Map;
        PdfDictionary widget = Assert.IsType<PdfDictionary>(Assert.Single(objects.Values, static item =>
            item.Value is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("T", out PdfObject? name) &&
            name is PdfStringObj text && text.Value == "Submit").Value);
        PdfDictionary appearances = Assert.IsType<PdfDictionary>(widget.Items["AP"]);
        PdfReference normalReference = Assert.IsType<PdfReference>(appearances.Items["N"]);
        PdfStream normalAppearance = Assert.IsType<PdfStream>(objects[normalReference.ObjectNumber].Value);
        PdfArray boundingBox = Assert.IsType<PdfArray>(normalAppearance.Dictionary.Items["BBox"]);

        Assert.Equal(new[] { 0D, 0D, 200D, 40D }, boundingBox.Items.Cast<PdfNumber>().Select(static number => number.Value));
    }

    [Fact]
    public void Edit_MoveRejectsResizingAnIconBearingPushButton() {
        byte[] source = BuildIconPushButtonPdf();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Open(source).Forms.Edit(edit => edit.Move("Icon", 1, 20D, 80D, 120D, 40D)));

        Assert.Contains("icon", exception.Message, StringComparison.OrdinalIgnoreCase);
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

    [Fact]
    public void Edit_SetFlagsRejectsConvertingAButtonFieldToPushButton() {
        byte[] source = PdfDocument.Create().CheckBox("Action", isChecked: true).ToBytes();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Forms.Edit(edit =>
            edit.SetFlags("Action", 1 << 16)));

        Assert.Contains("Converting", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Edit_UsesTheLastFlagsAssignmentInOneTransaction() {
        byte[] source = PdfDocument.Create().TextField("Name", value: "Ada").ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .SetFlags("Name", 1)
            .SetFlags("Name", 2));

        Assert.Equal(2, Assert.Single(result.Fields).Flags);
    }

    [Fact]
    public void Edit_SetFlagsRegeneratesAppearanceFromInheritedValue() {
        byte[] source = BuildInheritedValueFieldPdf();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit.SetFlags("Group.Name", 4096));
        Assert.Equal("Inherited value", Assert.Single(result.Fields).Value);
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result.ToBytes(), null).Map;
        PdfDictionary widget = Assert.IsType<PdfDictionary>(Assert.Single(objects.Values, static item =>
            item.Value is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("T", out PdfObject? name) &&
            name is PdfStringObj text && text.Value == "Name").Value);
        PdfDictionary appearances = Assert.IsType<PdfDictionary>(widget.Items["AP"]);
        PdfReference normalReference = Assert.IsType<PdfReference>(appearances.Items["N"]);
        PdfStream normalAppearance = Assert.IsType<PdfStream>(objects[normalReference.ObjectNumber].Value);

        Assert.Contains("496E686572697465642076616C7565", Encoding.ASCII.GetString(normalAppearance.Data), StringComparison.Ordinal);
    }

    private static byte[] BuildInheritedValueFieldPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.4",
        "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 200] /Resources << /Font << /Helv 8 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>", "endobj",
        "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
        "5 0 obj", "<< /Fields [6 0 R] /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 8 0 R >> >> >>", "endobj",
        "6 0 obj", "<< /FT /Tx /T (Group) /V (Inherited value) /Kids [7 0 R] >>", "endobj",
        "7 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 6 0 R /T (Name) /Rect [20 100 220 125] /DA (/Helv 10 Tf 0 g) /F 4 >>", "endobj",
        "8 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
        "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
    }));

    private static byte[] BuildIconPushButtonPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.4",
        "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 200] /Contents 4 0 R /Annots [7 0 R] >>", "endobj",
        "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
        "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
        "6 0 obj", "<< /FT /Btn /Ff 65536 /T (Icon) /Kids [7 0 R] >>", "endobj",
        "7 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 100 100 125] /MK << /I 8 0 R >> /AP << /N 8 0 R >> /F 4 >>", "endobj",
        "8 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 80 25] /Length 0 >>", "stream", "", "endstream", "endobj",
        "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
    }));

    [Fact]
    public void Edit_RejectsUndefinedTabOrderValuesAtThePublicBoundary() {
        var session = new PdfAcroFormEditSession();

        Assert.Throws<ArgumentOutOfRangeException>(() => session.SetTabOrder(1, (PdfPageTabOrder)99));
    }
}
