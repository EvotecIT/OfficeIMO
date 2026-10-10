using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormNumericConstraintTests {
    [Theory]
    [InlineData("123.45", false)]
    [InlineData("1,000.00", false)]
    [InlineData("1,00.00", true)]
    [InlineData("123.456", true)]
    [InlineData("1000.01", true)]
    [InlineData("-1", true)]
    [InlineData("1e2", true)]
    public void IndependentParentConstraintsEnforceFormattingPrecisionAndRange(string value, bool errors) {
        var field = PdfFormOcrReviewTests.Fixture("reportlab-scanned-form.pdf").Inspect().FormFieldsByName["Amount"];
        Assert.Equal(errors, PdfFormFieldValueAssessment.Assess(field, value).HasErrors);
    }

    [Theory]
    [InlineData("AFNumber_Keystroke(2, 0, 0, 0, '', true); app.alert('bad');")]
    [InlineData("// AFNumber_Keystroke(2, 0, 0, 0, '', true);")]
    [InlineData("AFNumber_Keystroke(2, 0, 0, 0, '$', true);")]
    [InlineData("AFNumber_Keystroke(2, 0, 0, 0, '', true); eval('1');")]
    [InlineData("AFNumber_Keystroke(99, 0, 0, 0, '', true);")]
    public void AdditionalOrUnsupportedScriptCannotMasqueradeAsQualifiedNumericMetadata(string source) {
        var field = new PdfFormField(1, "Amount", "Amount", "Tx", "", null, null, null,
            actions: [new("K", "JavaScript", source, null)]);
        var issue = Assert.Single(PdfFormFieldValueAssessment.Assess(field, "123.45").Issues);
        Assert.Equal(PdfFormFieldValueIssueCode.UnsupportedScriptConstraint, issue.Code);
        Assert.True(issue.IsError);
    }

    [Fact]
    public void DecimalCommaAndNoGroupingAreAssessedWithoutExecutingScripts() {
        var field = new PdfFormField(1, "Amount", "Amount", "Tx", "", null, null, null,
            actions: [new("K", "JavaScript", "AFNumber_Keystroke(2, 3, 0, 0, '', false)", null)]);
        Assert.False(PdfFormFieldValueAssessment.Assess(field, "123,45").HasErrors);
        Assert.True(PdfFormFieldValueAssessment.Assess(field, "123.45").HasErrors);
        Assert.True(PdfFormFieldValueAssessment.Assess(field, "1.123,45").HasErrors);
    }

    [Fact]
    public void ManualFillPreservesOpaqueScriptsWhileInteractiveReviewRejectsThem() {
        var source = PdfFormOcrReviewTests.Fixture("reportlab-scanned-form.pdf");
        var field = source.Inspect().FormFieldsByName["CustomRule"];
        Assert.True(PdfFormFieldValueAssessment.Assess(field, "78").HasErrors);
        var result = source.Forms.Fill(new Dictionary<string, PdfFormFieldValue> { ["CustomRule"] = "78" });
        Assert.Equal("78", result.Inspect().FormFieldsByName["CustomRule"].Value);
        Assert.Equal(field.JavaScript, result.Inspect().FormFieldsByName["CustomRule"].JavaScript);
    }

    [Fact]
    public void RewritePreservationDetectsChangedTerminalFieldActionPayload() {
        byte[] original = PdfFormOcrReviewTests.Fixture("reportlab-scanned-form.pdf").ToBytes();
        byte[] changed = PdfEncoding.Latin1GetBytes(PdfEncoding.Latin1GetString(original)
            .Replace("AFNumber", "BFNumber"));
        Assert.NotEqual(PdfDocument.Load(original).Inspect().FormFieldsByName["Amount"].JavaScript,
            PdfDocument.Load(changed).Inspect().FormFieldsByName["Amount"].JavaScript);
        var report = PdfRewritePreservation.Assess(original, changed,
            new PdfRewritePreservationOptions { PreserveFormWidgetActions = true });
        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue => issue.Feature == "FormWidgetActions");
    }

    [Fact]
    public void ParentActionsRemainOwnedDuringEditingAndAreRemovedWithFlattenedFields() {
        var source = PdfFormOcrReviewTests.Fixture("reportlab-scanned-form.pdf");
        string? script = source.Inspect().FormFieldsByName["Amount"].JavaScript;
        var edited = source.Forms.Edit(edit => edit.SetDefaultValue("Amount", "200.50"));
        var changed = PdfDocument.Load(edited.ToBytes());
        Assert.Equal("200.50", changed.Inspect().FormFieldsByName["Amount"].DefaultValue);
        Assert.Equal(script, changed.Inspect().FormFieldsByName["Amount"].JavaScript);
        Assert.True(edited.PreservationReport.IsPreserved);

        byte[] selected = PdfFormFiller.FlattenFields(changed.ToBytes(), ["Amount"]);
        var selectedInfo = PdfInspector.Inspect(selected);
        Assert.DoesNotContain("Amount", selectedInfo.FormFieldNames);
        Assert.Equal(source.Inspect().FormFieldsByName["CustomRule"].JavaScript,
            selectedInfo.FormFieldsByName["CustomRule"].JavaScript);
        Assert.True(PdfReadDocument.Open(selected).HasOnlyFormOwnedActiveContent());
        Assert.DoesNotContain("AFNumber", PdfEncoding.Latin1GetString(selected));

        byte[] flattened = PdfFormFiller.FlattenFields(changed.ToBytes());
        Assert.Empty(PdfInspector.Inspect(flattened).FormFields);
        Assert.False(PdfInspector.Inspect(flattened).HasActiveContent);
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("catalog", false)]
    [InlineData("page", false)]
    [InlineData("field-data", false)]
    [InlineData("unowned", false)]
    public void OnlyParsedTerminalFieldActionEntriesArePreservable(string extraLocation, bool allowed) {
        string action = "<< /S /JavaScript /JS (event.value=77;) >>";
        string catalogExtra = extraLocation == "catalog" ? " /OpenAction " + action : extraLocation == "unowned" ? " /OfficeIMO 8 0 R" : "";
        string pageExtra = extraLocation == "page" ? " /AA << /O " + action + " >>" : "";
        string fieldExtra = extraLocation == "field-data" ? " /OfficeIMO " + action : "";
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes("%PDF-1.7\n" +
            "1 0 obj << /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R" + catalogExtra + " >> endobj\n" +
            "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n" +
            "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [7 0 R]" + pageExtra + " >> endobj\n" +
            "5 0 obj << /Fields [6 0 R] /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 9 0 R >> >> >> endobj\n" +
            "6 0 obj << /FT /Tx /T (Amount) /V (before) /Kids [7 0 R] /AA << /V " + action + " >>" + fieldExtra + " >> endobj\n" +
            "7 0 obj << /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 160 48] /P 3 0 R >> endobj\n" +
            "8 0 obj " + action + " endobj\n9 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n" +
            "trailer << /Root 1 0 R /Size 10 >>\n%%EOF\n");
        var document = PdfDocument.Load(bytes);
        Assert.Equal(allowed, PdfReadDocument.Open(bytes).HasOnlyFormOwnedActiveContent());
        if (allowed) {
            var filled = document.Forms.Fill(new Dictionary<string, PdfFormFieldValue> { ["Amount"] = "78" });
            Assert.Equal("78", filled.Inspect().FormFieldsByName["Amount"].Value);
            Assert.Equal("event.value=77;", filled.Inspect().FormFieldsByName["Amount"].JavaScript);
        } else {
            Assert.Throws<PdfMutationBlockedException>(() => document.Forms.Fill(new Dictionary<string, PdfFormFieldValue> { ["Amount"] = "78" }));
            Assert.Throws<PdfMutationBlockedException>(() => document.Forms.Edit(edit => edit.SetDefaultValue("Amount", "78")));
            Assert.Throws<PdfMutationBlockedException>(() => PdfFormFiller.FlattenFields(bytes));
        }
    }
}
