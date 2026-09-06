using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormFieldValueAssessmentTests {
    [Fact]
    public void EditingDefaultValuesUsesTheSameUnicodeLengthAsCreatingFields() {
        var document = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)))
            .Forms.Edit(edit => edit.Create(new() { Name = "Code", Value = "AB", Style = new() { MaxLength = 2 } })).ToDocument();
        var updated = document.Forms.Edit(edit => edit.SetDefaultValue("Code", "A🚀"));
        Assert.Equal("A🚀", Assert.Single(updated.ToDocument().Inspect().FormFields).DefaultValue);
        Assert.Throws<ArgumentException>(() => document.Forms.Edit(edit => edit.SetDefaultValue("Code", "AB🚀")));
    }

    [Fact]
    public void RequiredValuesRemainSaveableAndLimitsCountUnicodeScalars() {
        var field = Field(new() { IsRequired = true, MaxLength = 2 }, "");
        var empty = PdfFormFieldValueAssessment.Assess(field, "");
        Assert.False(empty.HasErrors);
        Assert.Equal(PdfFormFieldValueIssueCode.RequiredValue, Assert.Single(empty.Issues).Code);
        Assert.Empty(PdfFormFieldValueAssessment.Assess(field, "A🚀").Issues);
        var longValue = PdfFormFieldValueAssessment.Assess(field, "AB🚀");
        Assert.True(longValue.HasErrors);
        Assert.Equal(PdfFormFieldValueIssueCode.MaximumLength, Assert.Single(longValue.Issues).Code);
        Assert.Equal("A🚀", Field(new() { MaxLength = 2, IsPassword = true }, "A🚀").Value);
    }

    [Fact]
    public void ReadOnlyAndScalarConstraintsAreReportedWithoutMutatingValues() {
        var field = Field(new() { IsReadOnly = true });
        var result = PdfFormFieldValueAssessment.Assess(field, PdfFormFieldValue.FromValues("one", "two"));
        Assert.True(result.HasErrors);
        Assert.Equal(new[] { PdfFormFieldValueIssueCode.ReadOnly, PdfFormFieldValueIssueCode.MultipleValues }, result.Issues.Select(issue => issue.Code));
        Assert.Equal("Original", field.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CustomChoiceDependsOnDeclaredEditableCapability(bool editable) {
        var document = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)))
            .Forms.Edit(edit => edit.Create(new() {
                Name = "Value", Kind = PdfFormFieldCreationKind.Choice, IsComboBox = true,
                ChoiceOptions = ["One", "Two"], Style = new() { IsEditableChoice = editable }
            })).ToDocument();
        var field = Assert.Single(document.Inspect().FormFields);
        Assert.False(PdfFormFieldValueAssessment.Assess(field, "One").HasErrors);
        Assert.Equal(!editable, PdfFormFieldValueAssessment.Assess(field, "Custom").HasErrors);
    }

    private static PdfFormField Field(PdfFormFieldStyle style, string value = "Original") => Assert.Single(
        PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)))
            .Forms.Edit(edit => edit.Create(new() { Name = "Value", Value = value, Style = style }))
            .ToDocument().Inspect().FormFields);
}
