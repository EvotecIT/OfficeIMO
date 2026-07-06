using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormAppearanceProofTests {
    [Fact]
    public void FormAppearanceProof_FillsChoiceDisplayTextAndFlattensVisibleAppearances() {
        PdfFormAppearanceProofResult proof = PdfFormAppearanceProofTestSupport.BuildFormAppearanceProof();

        Assert.True(proof.FilledInfo.HasReadableFormFields);
        Assert.Equal(false, proof.FilledInfo.AcroFormNeedAppearances);
        Assert.Equal(new[] { "Name", "Country", "AcceptTerms", "Payment.Method", "Notes", "Code", "Countries" }, proof.FilledInfo.FormFieldNames);
        Assert.Equal("Visible Value", proof.FilledInfo.FormFieldsByName["Name"].Value);

        PdfFormField country = proof.FilledInfo.FormFieldsByName["Country"];
        Assert.Equal("PL", country.Value);
        Assert.Equal("Poland", Assert.Single(country.SelectedOptions).DisplayText);

        PdfFormField acceptTerms = proof.FilledInfo.FormFieldsByName["AcceptTerms"];
        Assert.True(acceptTerms.IsCheckBox);
        Assert.Equal("Yes", acceptTerms.Value);
        PdfFormWidget acceptWidget = Assert.Single(acceptTerms.Widgets);
        Assert.Equal("Yes", acceptWidget.AppearanceState);
        Assert.Contains("Yes", acceptWidget.NormalAppearanceStates);
        Assert.Contains("Off", acceptWidget.NormalAppearanceStates);

        PdfFormField paymentMethod = proof.FilledInfo.FormFieldsByName["Payment.Method"];
        Assert.True(paymentMethod.IsRadioButton);
        Assert.Equal("Wire", paymentMethod.Value);
        Assert.Contains(paymentMethod.Widgets, widget => widget.AppearanceState == "Wire");
        Assert.Contains(paymentMethod.Widgets, widget => widget.AppearanceState == "Off");
        Assert.All(paymentMethod.Widgets, widget => Assert.Contains("Off", widget.NormalAppearanceStates));

        PdfFormField notes = proof.FilledInfo.FormFieldsByName["Notes"];
        Assert.True(notes.IsMultiline);
        Assert.Equal(PdfFormFieldTextAlignment.Center, notes.TextAlignment);
        Assert.Equal("Line one\nLine two", notes.Value);

        PdfFormField code = proof.FilledInfo.FormFieldsByName["Code"];
        Assert.True(code.IsComb);
        Assert.Equal(4, code.MaxLength);
        Assert.Equal("ZX91", code.Value);

        PdfFormField countries = proof.FilledInfo.FormFieldsByName["Countries"];
        Assert.True(countries.AllowsMultipleSelection);
        Assert.Equal(new[] { "DE", "US" }, countries.Values);
        Assert.Equal(new[] { "Germany", "United States" }, countries.SelectedOptions.Select(option => option.DisplayText).ToArray());

        Assert.Contains("/AP << /N", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<56697369626C652056616C7565> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<506F6C616E64> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<4C696E65206F6E65> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<4C696E652074776F> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<5A> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<58> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<39> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<31> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.DoesNotContain("<5A583931> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.DoesNotContain("<4765726D616E792C20556E6974656420537461746573> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("/AS /Yes", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("/AS /Wire", proof.FilledRaw, StringComparison.Ordinal);

        Assert.False(proof.FlattenedInfo.HasForms);
        Assert.DoesNotContain("/AcroForm", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<56697369626C652056616C7565> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<506F6C616E64> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<4C696E65206F6E65> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<4C696E652074776F> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<5A> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<58> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<39> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<31> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.DoesNotContain("<5A583931> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.DoesNotContain("<4765726D616E792C20556E6974656420537461746573> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("1.25 w", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("Wire selected", proof.FlattenedAppearanceText, StringComparison.Ordinal);
    }
}
