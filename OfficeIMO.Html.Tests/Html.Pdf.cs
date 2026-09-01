using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Filters;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfTests {
    [Fact]
    public void Pdf_SaveAsHtmlAsync_LinksTheMethodTokenIntoRenderOptions() {
        using var optionsCancellation = new System.Threading.CancellationTokenSource();
        using var methodCancellation = new System.Threading.CancellationTokenSource();
        var options = PdfHtmlSaveOptions.CreateSemanticProfile();
        options.CancellationToken = optionsCancellation.Token;

        PdfHtmlSaveOptions renderOptions = PdfHtmlConverterExtensions.CreateAsyncRenderOptions(
            options,
            methodCancellation.Token,
            out System.Threading.CancellationTokenSource? linkedCancellation);

        using (linkedCancellation) {
            Assert.NotSame(options, renderOptions);
            Assert.False(renderOptions.CancellationToken.IsCancellationRequested);
            methodCancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => renderOptions.CancellationToken.ThrowIfCancellationRequested());
        }
    }

    [Fact]
    public void Pdf_ToHtmlResult_StopsAtTheConfiguredOutputCharacterLimit() {
        PdfHtmlSaveOptions options = PdfHtmlSaveOptions.CreatePositionedReviewProfile();
        options.MaximumOutputCharacters = 128;

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfCore.PdfDocumentReadResult.Load(CreateLogicalSamplePdf()).ToHtmlResult(options));

        Assert.Contains("128-character output limit", exception.Message, StringComparison.Ordinal);
        Assert.Contains("being rendered", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_DoesNotTranslateInvalidProfileAsAnOutputLimitFailure() {
        PdfHtmlSaveOptions options = PdfHtmlSaveOptions.CreateSemanticProfile();
        options.Profile = (PdfHtmlProfile)int.MaxValue;
        options.MaximumOutputCharacters = 128;

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfCore.PdfDocumentReadResult.Load(CreateLogicalSamplePdf()).ToHtmlResult(options));

        Assert.Equal(nameof(PdfHtmlSaveOptions.Profile), exception.ParamName);
        Assert.DoesNotContain("output limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pdf_ToHtmlResult_EnforcesTheLimitAfterRequestedNewlineNormalization() {
        PdfCore.PdfDocumentReadResult document = PdfCore.PdfDocumentReadResult.Load(CreateLogicalSamplePdf());
        PdfHtmlSaveOptions unbounded = PdfHtmlSaveOptions.CreateSemanticProfile();
        unbounded.NewLine = "\r\n";
        string expected = document.ToHtmlResult(unbounded).Value;
        PdfHtmlSaveOptions bounded = PdfHtmlSaveOptions.CreateSemanticProfile();
        bounded.NewLine = "\r\n";
        bounded.MaximumOutputCharacters = expected.Length - 1;

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            document.ToHtmlResult(bounded));

        Assert.Contains("output limit", exception.Message, StringComparison.Ordinal);
        Assert.Contains("being rendered", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfToHtml_ResultAndBodyClassAreImmutableComposedContracts() {
        PdfHtmlSaveOptions options = PdfHtmlSaveOptions.CreateSemanticProfile();
        options.DocumentOutput.BodyClass = "customer-shell officeimo-html customer-shell";

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(CreateLogicalSamplePdf()).ToHtmlResult(options);

        Assert.Contains(
            "<body class=\"officeimo-html officeimo-pdf-html officeimo-pdf-semantic customer-shell\"",
            result.Value,
            StringComparison.Ordinal);
        Assert.True(result.Report.IsReadOnly);
        Assert.Throws<NotSupportedException>(() =>
            ((IList<PdfCore.PdfConversionWarning>)result.Report.Warnings).Clear());
        Assert.Throws<InvalidOperationException>(() => result.Report.Add(
            new PdfCore.PdfConversionWarning(
                "OfficeIMO.Tests",
                "Late",
                "PDF to HTML",
                "late",
                PdfCore.PdfConversionWarningSeverity.Warning)));
    }

    [Fact]
    public void HtmlToPdf_StandardControlsBecomeAccessibleInteractiveFormFields() {
        const string html = """
            <form>
              <label for="contact">Contact name</label>
              <input id="contact" name="contact" value="Ada" required maxlength="32">
              <label><input type="checkbox" name="accept" value="Accepted" checked disabled> Accept terms</label>
              <label for="country">Country</label>
              <select id="country" name="country"><option>Poland</option><option selected>Germany</option></select>
              <label for="notes">Review notes</label>
              <textarea id="notes" name="notes" readonly>Line one&#10;Line two</textarea>
              <label><input type="radio" name="method" value="Email"> Email</label>
              <label><input type="radio" name="method" value="Phone" checked> Phone</label>
            </form>
            """;

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(5, info.FormFields.Count);
        PdfCore.PdfFormField contact = Assert.Single(info.FormFields, field => field.Name == "contact");
        Assert.Equal("Ada", contact.Value);
        Assert.True(contact.IsRequired);
        Assert.Equal(32, contact.MaxLength);
        Assert.Equal("Contact name", contact.AlternateName);

        PdfCore.PdfFormField accept = Assert.Single(info.FormFields, field => field.Name == "accept");
        Assert.True(accept.IsCheckBox);
        Assert.True(accept.IsReadOnly);
        Assert.True(accept.IsNoExport);
        Assert.Equal("Accepted", accept.Value);

        PdfCore.PdfFormField country = Assert.Single(info.FormFields, field => field.Name == "country");
        Assert.True(country.IsCombo);
        Assert.Equal("Germany", country.Value);
        Assert.Equal(new[] { "Poland", "Germany" }, country.Options.Select(option => option.DisplayText).ToArray());

        PdfCore.PdfFormField notes = Assert.Single(info.FormFields, field => field.Name == "notes");
        Assert.True(notes.IsMultiline);
        Assert.True(notes.IsReadOnly);
        Assert.False(notes.IsNoExport);
        Assert.Equal("Line one\nLine two", notes.Value);

        PdfCore.PdfFormField method = Assert.Single(info.FormFields, field => field.Name == "method");
        Assert.True(method.IsRadioButton);
        Assert.Equal("Phone", method.Value);
        Assert.Equal(2, method.Widgets.Count);
    }

    [Fact]
    public void HtmlToPdf_TextAreaPreservesAuthoredEdgeWhitespaceAcrossFieldAndStaticAppearance() {
        const string content = "  Alpha  \n";
        const string html = "<textarea id='edges' name='edges' style='width:160px;height:40px;font:12px Arial'>  Alpha  &#10;</textarea>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderVisual[] scene = rendered.Pages.SelectMany(page => EnumeratePdfSceneVisuals(page.Scene)).ToArray();
        HtmlRenderFormField renderedField = Assert.Single(scene.OfType<HtmlRenderFormField>());
        string appearance = string.Concat(scene
            .OfType<HtmlRenderText>()
            .Where(text => text.Source == "textarea#edges")
            .Select(text => text.Text));
        PdfCore.PdfFormField pdfField = Assert.Single(PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.Equal(content, renderedField.Value);
        Assert.Equal("  Alpha  ", appearance);
        Assert.Equal(content, pdfField.Value);
    }

    [Fact]
    public void HtmlToPdf_TextInputPreservesAuthoredWhitespaceAcrossFieldAndStaticAppearance() {
        const string content = "  A  B  ";
        const string html = "<input id='spaces' name='spaces' value='  A  B  ' style='width:180px;font:12px Arial'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderVisual[] scene = rendered.Pages.SelectMany(page => EnumeratePdfSceneVisuals(page.Scene)).ToArray();
        HtmlRenderFormField renderedField = Assert.Single(scene.OfType<HtmlRenderFormField>());
        HtmlRenderText appearance = Assert.Single(scene.OfType<HtmlRenderText>(), text => text.Source == "input#spaces");
        PdfCore.PdfFormField pdfField = Assert.Single(PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.Equal(content, renderedField.Value);
        Assert.Equal(content, appearance.Text);
        Assert.Equal(content, pdfField.Value);
    }

    [Fact]
    public void HtmlToPdf_PasswordSerializesOnlyTheMaskedAppearance() {
        const string password = "OfficeIMO-secret-2026";
        string html = "<input id='secret' name='secret' type='password' value='" + password + "' style='width:400px'>";
        var options = new HtmlPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions { CompressContentStreams = false }
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string raw = Encoding.ASCII.GetString(pdf);

        Assert.True(field.IsPassword);
        Assert.Equal(string.Empty, field.Value);
        Assert.DoesNotContain(password, raw, StringComparison.Ordinal);
        string passwordHex = BitConverter.ToString(Encoding.UTF8.GetBytes(password)).Replace("-", string.Empty);
        Assert.DoesNotContain(passwordHex, raw, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<" + string.Concat(Enumerable.Repeat("2A", password.Length)) + "> Tj", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ArabicPresentationFormsPreserveLogicalExtraction() {
        if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfCore.PdfEmbeddedFontFamily? family) ||
            family == null) {
            return;
        }

        const string arabic = "العربية";
        var options = new HtmlPdfSaveOptions { FontFamily = family };
        byte[] pdf = HtmlConversionDocument.Parse(
            "<p style='font-family:Arial'>Evidence · " + arabic + " · terminal</p>").ToPdf(options);
        using PdfPigDocument document = PdfPigDocument.Open(new MemoryStream(pdf));
        string extracted = string.Concat(document.GetPages().Select(page => page.Text));
        string logical = OfficeArabicTextShaper.ToLogicalText(extracted);
        string canonical = string.Concat(logical.Where(character =>
            !char.IsWhiteSpace(character) &&
            CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.Format));

        Assert.Contains(arabic, canonical, StringComparison.Ordinal);
        Assert.DoesNotContain(canonical, character => character >= '\uFE70' && character <= '\uFEFF');
    }

    [Fact]
    public void HtmlToPdf_ZeroSelectSizeUsesAComboBoxAndTheFirstEnabledOption() {
        const string html = "<select name='choice' size='0'><option value='first'>First</option><option value='second'>Second</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderFormField renderedField = Assert.Single(
            EnumeratePdfSceneVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderFormField>());
        PdfCore.PdfFormField pdfField = Assert.Single(
            PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.True(renderedField.IsComboBox);
        Assert.Equal("first", renderedField.Value);
        Assert.Contains("First", rendered.Text, StringComparison.Ordinal);
        Assert.True(pdfField.IsCombo);
        Assert.Equal("first", pdfField.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" multiple size='3'")]
    public void HtmlToPdf_EmptySelectsUseTruthfulStaticFallback(string attributes) {
        string html = "<select name='choice'" + attributes + "></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ChoiceEmptyOptionsStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_DisabledAndReadOnlyControlsAreExcludedFromRequiredConstraintValidation() {
        const string html = "<input name='enabled' required><input name='disabled' required disabled><textarea name='readonly' required readonly></textarea>";

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf());

        Assert.True(Assert.Single(info.FormFields, field => field.Name == "enabled").IsRequired);
        Assert.False(Assert.Single(info.FormFields, field => field.Name == "disabled").IsRequired);
        Assert.False(Assert.Single(info.FormFields, field => field.Name == "readonly").IsRequired);
    }

    [Fact]
    public void HtmlToPdf_UnnamedControlsRemainInteractiveButAreExcludedFromFormData() {
        const string html = "<input value='secret'><input id='by-id' value='identifier'><input name='' value='empty-name'><input name='named' value='included'>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfFormDataSet exported = PdfCore.PdfDocument.Load(pdf).Forms.ExportData();

        Assert.Equal(4, info.FormFields.Count);
        Assert.Equal(3, info.FormFields.Count(field => field.IsNoExport));
        Assert.All(info.FormFields.Where(field => field.IsNoExport), field => Assert.Null(field.MappingName));
        PdfCore.PdfFormDataField field = Assert.Single(exported.Fields);
        Assert.Equal("named", field.Name);
        Assert.Equal(new[] { "included" }, field.Values);
    }

    [Theory]
    [InlineData("linear-gradient(90deg,transparent,blue)")]
    [InlineData("radial-gradient(circle,transparent,blue)")]
    public void HtmlToPdf_TranslucentGradientsUseFaithfulManagedRasterFallback(string gradient) {
        string html = "<div style='width:40px;height:20px;background:" + gradient + "'></div>";

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult();
        byte[] pdf = result.ToBytes();

        Assert.Contains(result.Report.Warnings, warning => warning.Code == "HtmlPdfTranslucentGradientRasterized");
        Assert.Contains("/SMask", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_PageRelayoutPreservesInteractiveFieldNameIdentity() {
        const string html = """
            <style>
              @page { size:300px 180px; margin:50px; }
              @page report { size:400px 180px; margin-left:100px; margin-right:100px; margin-top:50px; margin-bottom:50px; }
            </style>
            <p style="margin:0">Opening page</p>
            <input name="contact" value="Ada" style="page:report;break-before:page;width:50vw">
            """;

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);

        Assert.Equal("contact", field.Name);
        Assert.Equal("contact", field.MappingName);
        Assert.Equal("Ada", field.Value);
    }

    [Fact]
    public void HtmlToPdf_EmptyInteractiveTextFieldPreservesPlaceholderAsItsInitialAppearance() {
        const string html = "<input name='email' placeholder='Email address'>";

        HtmlRenderFormField renderedField = Assert.Single(
            EnumeratePdfSceneVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene).OfType<HtmlRenderFormField>());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);

        Assert.Equal("Email address", renderedField.Placeholder);
        Assert.Equal("email", field.Name);
        Assert.Equal(string.Empty, field.Value);
        var (objects, _) = PdfCore.PdfSyntax.ParseObjects(pdf);
        PdfCore.PdfDictionary fieldObject = Assert.IsType<PdfCore.PdfDictionary>(objects.Values
            .Select(item => item.Value)
            .Single(item => item is PdfCore.PdfDictionary dictionary && dictionary.Get<PdfCore.PdfStringObj>("T")?.Value == "email"));
        PdfCore.PdfDictionary appearance = Assert.IsType<PdfCore.PdfDictionary>(fieldObject.Items["AP"]);
        PdfCore.PdfReference normalAppearance = Assert.IsType<PdfCore.PdfReference>(appearance.Items["N"]);
        PdfCore.PdfStream stream = Assert.IsType<PdfCore.PdfStream>(objects[normalAppearance.ObjectNumber].Value);
        string appearanceContent = Encoding.ASCII.GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
        Assert.Contains("BT", appearanceContent, StringComparison.Ordinal);
        Assert.Contains("Tj", appearanceContent, StringComparison.Ordinal);

        byte[] flattened = PdfCore.PdfFormFiller.FlattenFields(pdf);
        Assert.False(PdfCore.PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.Contains("Email address", PdfCore.PdfReadDocument.Open(flattened).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ChoiceFieldsPreserveExportValuesAndDuplicateDisplayLabels() {
        const string html = """
            <select name="country">
              <option value="US">Same label</option>
              <option value="CA" selected>Same label</option>
            </select>
            """;

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);

        Assert.Equal("CA", field.Value);
        Assert.Equal(new[] { "US", "CA" }, field.Options.Select(option => option.ExportValue).ToArray());
        Assert.Equal(new[] { "Same label", "Same label" }, field.Options.Select(option => option.DisplayText).ToArray());
        Assert.Contains("Same label", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ListBoxAppearanceShowsAuthoredRowsAndSelection() {
        const string html = "<select name='letters' multiple size='3'><option value='a'>Alpha</option><option value='b' selected>Beta</option><option value='g'>Gamma</option></select>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        string appearance = GetFieldAppearanceContent(pdf, "letters");

        Assert.Contains("<416C706861> Tj", appearance, StringComparison.Ordinal);
        Assert.Contains("<42657461> Tj", appearance, StringComparison.Ordinal);
        Assert.Contains("<47616D6D61> Tj", appearance, StringComparison.Ordinal);
        Assert.Contains("0.153 0.392 0.8 rg", appearance, StringComparison.Ordinal);
        Assert.Contains("1 1 1 rg", appearance, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ListBoxSelectedIndexDisambiguatesDuplicateExportValues() {
        const string html = "<select name='choice' size='2'><option value='x'>First</option><option value='x' selected>Second</option></select>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        string syntax = Encoding.ASCII.GetString(pdf);
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);

        Assert.Contains("/I [1]", syntax, StringComparison.Ordinal);
        Assert.Equal(new[] { 1 }, field.SelectedIndices);
        Assert.Equal("Second", Assert.Single(field.SelectedOptions).DisplayText);
        Assert.Contains("<4669727374> Tj", GetFieldAppearanceContent(pdf, "choice"), StringComparison.Ordinal);
        Assert.Contains("<5365636F6E64> Tj", GetFieldAppearanceContent(pdf, "choice"), StringComparison.Ordinal);
        string searchableText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Second", searchableText, StringComparison.Ordinal);
        Assert.DoesNotContain("First", searchableText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ComboWithAmbiguousDuplicateExportUsesTruthfulStaticFallback() {
        const string html = "<select name='choice'><option value='x'>First</option><option value='x' selected>Second</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ChoiceDuplicateSelectedValueStaticFallback);
        Assert.Contains("Second", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_UniformRoundedControlUsesRoundedWidgetAppearance() {
        const string html = "<input name='rounded' value='Rounded' style='width:120px;height:28px;border:2px solid #123456;border-radius:10px;background:#ffffff'>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        string appearance = GetFieldAppearanceContent(pdf, "rounded");

        Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains(" c h f", appearance, StringComparison.Ordinal);
        Assert.Contains(" c h S", appearance, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_NonUniformRoundedControlUsesTruthfulStaticFallback() {
        const string html = "<input name='rounded' value='Static rounded' style='width:120px;height:28px;border-radius:10px 2px'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldNonUniformRadiusStaticFallback);
        Assert.Contains("Static rounded", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_DashedControlBorderRemainsDashedInWidgetAndAppearance() {
        const string html = "<input name='dashed' value='Dashed' style='width:120px;height:28px;border:2px dashed #ff0000'>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        string syntax = Encoding.ASCII.GetString(pdf);
        string appearance = GetFieldAppearanceContent(pdf, "dashed");

        Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("/BS << /S /D /W 1.5", syntax, StringComparison.Ordinal);
        Assert.Contains("[3] 0 d", appearance, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("dotted")]
    [InlineData("double")]
    public void HtmlToPdf_UnrepresentableControlBorderUsesTruthfulStaticFallback(string borderStyle) {
        string html = "<input name='styled' value='Static border' style='border:2px " + borderStyle + " red'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldBorderStyleStaticFallback);
        Assert.Contains("Static border", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_NoWrapTextareaUsesTruthfulStaticFallback() {
        const string html = "<textarea name='notes' wrap='off' style='width:70px;height:60px;font:12px Arial'>Alpha beta gamma delta</textarea>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldNoWrapStaticFallback);
        Assert.Contains(HtmlRenderDiagnosticCodes.FormFieldNoWrapStaticFallback, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.FormFieldNoWrapStaticFallback, out _));
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string text = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("Alphabetagammadelta", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_UnrepresentableControlTypographyUsesTruthfulStaticFallback() {
        const string html = "<input name='bold' value='Bold value' style='font-weight:bold'>"
            + "<textarea name='italic' style='font-style:italic'>Italic value</textarea>"
            + "<select name='family' style='font-family:Courier New'><option selected>Courier value</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        string text = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Equal(3, rendered.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldTypographyStaticFallback));
        Assert.Contains(HtmlRenderDiagnosticCodes.FormFieldTypographyStaticFallback, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.FormFieldTypographyStaticFallback, out _));
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Boldvalue", text, StringComparison.Ordinal);
        Assert.Contains("Italicvalue", text, StringComparison.Ordinal);
        Assert.Contains("Couriervalue", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("background:rgba(0,0,255,.2)")]
    [InlineData("border:2px solid rgba(255,0,0,.4)")]
    [InlineData("color:rgba(0,128,0,.6)")]
    public void HtmlToPdf_TranslucentControlPaintUsesTruthfulStaticFallback(string style) {
        string html = "<input name='styled' value='Static alpha' style='" + style + "'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldColorTransparencyStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Static alpha", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_TranslucentRadioPaintMakesTheEntireGroupStatic() {
        const string html = "<input type='radio' name='choice' value='one'><input type='radio' name='choice' value='two' checked style='background:rgba(0,0,255,.2)'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldColorTransparencyStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_ControlBackgroundImageUsesTruthfulStaticFallback() {
        const string html = "<input name='gradient' value='Static gradient' style='background:linear-gradient(red,blue)'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldBackgroundImageStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Static gradient", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_RadioBackgroundImageMakesTheEntireGroupStatic() {
        const string html = "<input type='radio' name='choice' value='one'>"
            + "<input type='radio' name='choice' value='two' checked style='background:linear-gradient(red,blue)'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldBackgroundImageStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_BlankChoiceLabelsUseTruthfulStaticFallback() {
        const string html = "<select name='choice'><option value='' selected></option><option value='one'>One</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ChoiceBlankLabelStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.DoesNotContain("Option 1", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_CanKeepStaticFormControlPaint() {
        const string html = "<label>Name <input name='name' value='Static Ada'></label>";
        var options = new HtmlPdfSaveOptions { InteractiveFormControls = false };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Static Ada", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.False(options.ClonePdf().InteractiveFormControls);
    }

    [Fact]
    public void HtmlToPdf_RadioGroupPreservesDifferentlySizedWidgets() {
        const string html = """
            <input type="radio" name="size" value="Small" style="width:12px;height:12px">
            <input type="radio" name="size" value="Large" checked style="width:24px;height:18px">
            """;

        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.Equal("Large", field.Value);
        Assert.Equal(10.5D, field.Widgets[0].Width, 3);
        Assert.Equal(19.5D, field.Widgets[1].Width, 3);
        Assert.Equal(15D, field.Widgets[1].Height, 3);
    }

    [Fact]
    public void HtmlToPdf_RadioWidgetsRetainEachOptionsAccessibleName() {
        const string html = "<input type='radio' name='contact' value='email' aria-label='Email option'><input type='radio' name='contact' value='phone' aria-label='Phone option'>";

        string raw = Encoding.ASCII.GetString(HtmlConversionDocument.Parse(html).ToPdf());

        Assert.Contains("/TU <456D61696C206F7074696F6E>", raw, StringComparison.Ordinal);
        Assert.Contains("/TU <50686F6E65206F7074696F6E>", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_RadioGroupMergesRequiredStateAcrossWidgetsAndPages() {
        const string html = """
            <style>@page{size:3in 2in;margin:10px}</style>
            <form>
              <input type="radio" name="delivery" value="Email">
              <input type="radio" name="delivery" value="Phone" required>
              <div style="break-before:page"></div>
              <input type="radio" name="delivery" value="Post">
            </form>
            """;

        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.True(field.IsRequired);
        Assert.Equal(3, field.Widgets.Count);
    }

    [Fact]
    public void HtmlToPdf_RadioGroupsUseExactAuthoredNameWhitespace() {
        const string html = "<form><input type='radio' name='a b' value='one' checked><input type='radio' name='a b' value='two'><input type='radio' name='a  b' value='three'><input type='radio' name='a  b' value='four' checked></form>";

        PdfCore.PdfFormField[] fields = PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields
            .OrderBy(field => field.Name, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(new[] { "a  b", "a b" }, fields.Select(field => field.Name).ToArray());
        Assert.Equal(new[] { "four", "one" }, fields.Select(field => field.Value).ToArray());
        Assert.All(fields, field => Assert.Equal(2, field.Widgets.Count));
    }

    [Fact]
    public void HtmlToPdf_DuplicateRadioValuesUseTruthfulStaticFallback() {
        const string html = "<input type='radio' name='answer' value='same'><input type='radio' name='answer' value='same' checked>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.RadioDuplicateValueStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_DuplicateSelectedChoiceValuesUseTruthfulStaticFallback() {
        const string html = "<select name='choice' multiple size='3'><option value='same' selected>First</option><option value='same' selected>Second</option><option value='other'>Other</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ChoiceDuplicateSelectedValueStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string searchableText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("First", searchableText, StringComparison.Ordinal);
        Assert.Contains("Second", searchableText, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<label><input type='radio' name='delivery' value='email'>Email</label><label><input type='radio' name='delivery' value='post' disabled checked>Post</label>")]
    [InlineData("<label><input type='radio' name='delivery' value='post' disabled checked>Post</label><label><input type='radio' name='delivery' value='email'>Email</label>")]
    public void HtmlToPdf_MixedDisabledRadioGroupUsesTruthfulStaticFallback(string html) {

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.RadioMixedDisabledStateStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_SelectWithDisabledOptionsUsesTruthfulStaticFallback() {
        const string html = "<select name='country'><option value='PL' selected>Poland</option><optgroup label='Unavailable' disabled><option value='DE'>Germany</option></optgroup></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ChoiceDisabledOptionStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Poland", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_CheckBoxPreservesUnicodeExportValueSeparatelyFromAppearanceState() {
        const string html = "<input type='checkbox' name='choice' value='caf\u00E9' checked>";

        HtmlRenderFormField renderedField = Assert.Single(EnumeratePdfSceneVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene).OfType<HtmlRenderFormField>());
        PdfCore.PdfFormField pdfField = Assert.Single(PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.Equal("caf\u00E9", renderedField.Value);
        Assert.NotEqual(renderedField.Value, renderedField.RadioOption);
        Assert.Equal("caf\u00E9", Assert.Single(pdfField.Options).ExportValue);
        Assert.Equal("caf\u00E9", Assert.Single(PdfCore.PdfDocument.Load(HtmlConversionDocument.Parse(html).ToPdf()).Forms.ExportData().Fields).Values[0]);
    }

    [Fact]
    public void HtmlToPdf_BlankCheckboxValuesUseTruthfulStaticFallback() {
        const string html = "<input type='checkbox' name='empty' value='' checked><input type='checkbox' name='spaces' value='   ' checked>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Equal(2, rendered.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldBlankButtonValueStaticFallback));
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_BlankRadioValueMakesWholeGroupStatic() {
        const string html = "<input type='radio' name='answer' value='' checked><input type='radio' name='answer' value='yes'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldBlankButtonValueStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_MissingButtonValuesUseHtmlOnDefault() {
        const string html = "<input type='checkbox' name='check' checked><input type='radio' name='radio' checked>";

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf());

        Assert.Equal("on", Assert.Single(info.FormFields, field => field.Name == "check").Value);
        Assert.Equal("on", Assert.Single(info.FormFields, field => field.Name == "radio").Value);
    }

    [Fact]
    public void HtmlToPdf_RadioButtonPreservesUnicodeExportValueSeparatelyFromAppearanceState() {
        const string html = "<input type='radio' name='choice' value='caf\u00E9' checked>";

        HtmlRenderFormField renderedField = Assert.Single(EnumeratePdfSceneVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene).OfType<HtmlRenderFormField>());
        PdfCore.PdfFormField pdfField = Assert.Single(PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf()).FormFields);

        Assert.Equal("caf\u00E9", renderedField.Value);
        Assert.NotEqual(renderedField.Value, renderedField.RadioOption);
        Assert.Equal("caf\u00E9", Assert.Single(pdfField.Options).ExportValue);
        Assert.Equal("caf\u00E9", Assert.Single(PdfCore.PdfDocument.Load(HtmlConversionDocument.Parse(html).ToPdf()).Forms.ExportData().Fields).Values[0]);
    }

    [Fact]
    public void HtmlToPdf_UnselectedScalarChoiceOmitsEmptyValueArrays() {
        const string html = "<select name='choice' size='2'><option value='one'>One</option><option value='two'>Two</option></select>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();
        string syntax = Encoding.ASCII.GetString(pdf);

        Assert.DoesNotContain("/V []", syntax, StringComparison.Ordinal);
        Assert.DoesNotContain("/DV []", syntax, StringComparison.Ordinal);
        Assert.False(Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields).HasValues);
    }

    [Fact]
    public void HtmlToPdf_PartiallyClippedControlUsesStaticAppearance() {
        const string html = "<div style='width:80px;height:24px;overflow:hidden'><input name='clipped' value='Clipped value' style='width:120px;height:20px'></div>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string searchableText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("Clippedvalue", searchableText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ControlInsideRoundedClipUsesStaticAppearance() {
        const string html = "<div style='position:relative;width:100px;height:40px;overflow:hidden;border-radius:20px'><input name='rounded' value='Rounded value' style='position:absolute;left:0;top:0;width:40px;height:18px'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Contains(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderPathClipGroup);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string searchableText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("Roundedvalue", searchableText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ControlOutsideClipDoesNotCreateAnInteractiveWidget() {
        const string html = "<div style='position:relative;width:80px;height:24px;overflow:hidden'><input name='outside' value='Outside value' style='position:absolute;left:100px;width:80px;height:20px'></div>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_TransformedControlUsesSearchableStaticAppearance() {
        const string html = "<input name='tilted' value='Tilted value' style='width:140px;transform:rotate(12deg)'>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string searchableText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("Tiltedvalue", searchableText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_AncestorTransformedControlUsesSearchableStaticAppearance() {
        const string html = "<div style='transform:rotate(12deg)'><input name='child' value='Descendant value'></div>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        string searchableText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("Descendantvalue", searchableText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_ZeroMaximumLengthUsesTruthfulStaticAppearance() {
        const string html = "<input name='empty' value='Authored value' maxlength='0'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldZeroMaximumLengthStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Authored value", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_InitialValueExceedingMaximumLengthUsesTruthfulStaticAppearance() {
        const string html = "<input name='code' value='abcd' maxlength='2'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldInitialValueExceedsMaximumLengthStaticFallback);
        Assert.Contains(HtmlRenderDiagnosticCodes.FormFieldInitialValueExceedsMaximumLengthStaticFallback, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.FormFieldInitialValueExceedsMaximumLengthStaticFallback, out _));
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("abcd", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_MultipleFileInputUsesTruthfulStaticAppearance() {
        const string html = "<input type='file' name='attachment'><input type='file' name='attachments' multiple>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        HtmlRenderFormField renderedField = Assert.Single(
            rendered.Pages.SelectMany(page => EnumeratePdfSceneVisuals(page.Scene)).OfType<HtmlRenderFormField>());
        Assert.Equal("attachment", renderedField.Name);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FileMultipleSelectionStaticFallback);
        PdfCore.PdfFormField pdfField = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Equal("attachment", pdfField.Name);
        Assert.True(pdfField.IsFileSelect);
        Assert.Contains("Choose file", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_EmptyFileInputRetainsItsVisiblePromptInTheInteractiveAppearance() {
        const string html = "<input type='file' name='attachment'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderFormField renderedField = Assert.Single(
            rendered.Pages.SelectMany(page => EnumeratePdfSceneVisuals(page.Scene)).OfType<HtmlRenderFormField>());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.True(renderedField.IsFileSelect);
        Assert.Empty(renderedField.Value);
        Assert.Equal("Choose file", renderedField.Placeholder);
        PdfCore.PdfFormField pdfField = Assert.Single(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.True(pdfField.IsFileSelect);
        byte[] flattened = PdfCore.PdfFormFiller.FlattenFields(pdf);
        Assert.False(PdfCore.PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.Contains("Choose file", PdfCore.PdfReadDocument.Open(flattened).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_RepeatedNonRadioNamesUseTruthfulStaticFallback() {
        const string html = """
            <form><input type="checkbox" name="tag" value="One"><input type="checkbox" name="tag" value="Two"></form>
            <form><input type="radio" name="status" value="Internal" checked><input type="radio" name="status" value="External"></form>
            <form><input type="radio" name="status" value="Public" checked><input type="radio" name="status" value="Private"></form>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf());

        Assert.Empty(info.FormFields.Where(field => field.MappingName == "tag"));
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback);
        PdfCore.PdfFormField[] statuses = info.FormFields.Where(field => field.MappingName == "status").ToArray();
        Assert.Equal(2, statuses.Length);
        Assert.All(statuses, field => Assert.Equal(2, field.Widgets.Count));
        Assert.Equal(2, statuses.Select(field => field.Name).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void HtmlToPdf_RepeatedTableHeaderAndFooterControlsUseTruthfulStaticFallback() {
        string rows = string.Concat(Enumerable.Range(1, 18).Select(index => "<tr><td style='height:28px'>Row " + index + "</td></tr>"));
        string html = "<table><thead><tr><th><input name='filter' value='All rows'></th></tr></thead>"
            + "<tbody>" + rows + "</tbody>"
            + "<tfoot><tr><td><input name='summary' value='Totals'></td></tr></tfoot></table>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(320D / HtmlRenderOptions.CssPixelsPerInch, 180D / HtmlRenderOptions.CssPixelsPerInch),
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.True(rendered.Pages.Count > 1);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields.Where(field => field.MappingName is "filter" or "summary"));
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("All rows", text, StringComparison.Ordinal);
        Assert.Contains("Totals", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_FixedControlsRepeatedAcrossPagesUseTruthfulStaticFallback() {
        string body = string.Concat(Enumerable.Range(1, 24).Select(index => "<p>Paragraph " + index + "</p>"));
        string html = "<input name='fixed-search' value='Search' style='position:fixed;top:4px;left:4px'>" + body;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(320D / HtmlRenderOptions.CssPixelsPerInch, 180D / HtmlRenderOptions.CssPixelsPerInch),
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.True(rendered.Pages.Count > 1);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields.Where(field => field.MappingName == "fixed-search"));
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback);
        Assert.Contains("Paragraph", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_AuthoredFieldNameWhitespaceIsPreservedExactly() {
        const string html = "<form><input name='a b' value='one'><input name='a  b' value='two'><input name=' edge ' value='three'></form>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf());

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback);
        Assert.Equal(new[] { " edge ", "a  b", "a b" }, info.FormFields.Select(field => field.Name).OrderBy(name => name, StringComparer.Ordinal).ToArray());
        Assert.Equal(new[] { " edge ", "a  b", "a b" }, info.FormFields.Select(field => field.MappingName).OrderBy(name => name, StringComparer.Ordinal).ToArray());
    }

    [Fact]
    public void HtmlToPdf_DottedAuthoredNamesUsePdfSafePartialNamesAndRetainMappingNames() {
        const string html = "<form><input name='user.email' value='one'><input name='user-email' value='two'>"
            + "<input type='radio' name='contact.kind' value='mail' checked><input type='radio' name='contact.kind' value='phone'></form>";

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdf());
        PdfCore.PdfFormField[] fields = info.FormFields.ToArray();

        Assert.Equal(3, fields.Length);
        Assert.All(fields, field => Assert.DoesNotContain('.', field.PartialName));
        Assert.Equal(3, fields.Select(field => field.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(fields, field => field.MappingName == "user.email");
        Assert.Contains(fields, field => field.MappingName == "user-email");
        Assert.Contains(fields, field => field.MappingName == "contact.kind" && field.IsRadioButton && field.Widgets.Count == 2);
    }

    [Fact]
    public void HtmlToPdf_WhitespaceOnlyFieldNameUsesTruthfulStaticFallback() {
        const string html = "<input name='   ' value='Authored value'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldBlankNameStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
        Assert.Contains("Authored value", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToPdf_MixedControlTypesSharingOneNameUseTruthfulStaticFallback() {
        const string html = "<form><input type='radio' name='answer' value='yes' checked><input type='checkbox' name='answer' value='details' checked></form>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.DoesNotContain(EnumeratePdfSceneVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback);
        Assert.Empty(PdfCore.PdfInspector.Inspect(pdf).FormFields);
    }

    [Fact]
    public void HtmlToPdf_SkipsEmptySvgWithAlternativeText() {
        string svg = Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'></svg>"));
        string html = "<img src='data:image/svg+xml;base64," + svg +
            "' alt='Empty vector'><p>After empty vector</p>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Contains(
            "After empty vector",
            PdfCore.PdfReadDocument.Open(pdf).ExtractText(),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Html_DirectOutputs_UseOneSharedOptionsShape() {
        const string html = "<main><h1>Quarterly report</h1><p>Direct HTML rendering.</p></main>";
        var options = new HtmlPdfSaveOptions {
            ViewportWidth = 640D,
            Margins = HtmlRenderMargins.All(24D),
            Scale = 1D
        };

        byte[] png = HtmlConversionDocument.Parse(html).ToPng(options);
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.True(png.Length > 8);
        Assert.True(pdf.Length > 8);
        Assert.StartsWith("<svg", svg, StringComparison.Ordinal);
        Assert.Equal(HtmlRenderMode.Paged, options.Mode);
    }

    [Fact]
    public void Html_ToPdfResult_ReturnsDiagnosticsWithoutMutatingReusableOptions() {
        const string html = "<p><img src='https://example.invalid/missing.png'>Report</p>";
        var options = new HtmlPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult first = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        PdfCore.PdfDocumentConversionResult second = OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Clean</p>").ToPdfDocumentResult(options);

        Assert.Contains(first.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.DoesNotContain(second.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.Equal(HtmlRenderMode.Paged, options.Mode);
    }

    [Fact]
    public async Task Html_Pdf_BytesDocumentFileAndStream_AreConsistent() {
        const string html = "<article><h1>API contract</h1><p>One direct renderer.</p></article>";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            byte[] bytes = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf();
            PdfCore.PdfDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocument();
            using var stream = new MemoryStream();
            await OfficeIMO.Html.HtmlConversionDocument.Parse(html).SaveAsPdfAsync(stream);
            OfficeIMO.Html.HtmlConversionDocument.Parse(html).SaveAsPdf(path);

            Assert.Equal((byte)'%', bytes[0]);
            Assert.True(document.ToBytes().Length > 8);
            Assert.Equal((byte)'%', stream.ToArray()[0]);
            Assert.True(new FileInfo(path).Length > 8L);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Html_OfficeProjections_AreExplicitTargets() {
        const string html = "<article><h1>Projection</h1><p>Explicit conversion.</p></article>";

        using OfficeIMO.Word.WordDocument word = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
        OfficeIMO.Markdown.MarkdownDoc markdown = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToMarkdownDocument();

        Assert.NotNull(word);
        Assert.Contains("Projection", markdown.ToMarkdown(), StringComparison.Ordinal);
    }
    [Fact]
    public void PdfHtml_ProfileContracts_CoverSupportedProfiles() {
        PdfHtmlProfileContract semantic = PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic);
        PdfHtmlProfileContract positioned = PdfHtmlProfileContracts.Get(PdfHtmlProfile.PositionedReview);

        Assert.Equal(2, PdfHtmlProfileContracts.All.Count);
        Assert.Equal(HtmlConversionProfile.Semantic, semantic.SharedProfile);
        Assert.Equal("pdf-html-semantic", semantic.Id);
        Assert.Contains("logical model", semantic.Pipeline, StringComparison.Ordinal);
        Assert.Contains("Search", semantic.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("OCR", semantic.UnsupportedScope, StringComparison.Ordinal);
        Assert.Contains("tables", semantic.PreservedSignals);
        Assert.Contains("export-summary", semantic.OutputArtifacts);
        Assert.Contains("no-editable-office-reconstruction", semantic.RendererBoundaries);
        Assert.Equal(HtmlConversionProfile.PositionedReview, positioned.SharedProfile);
        Assert.Equal("pdf-html-positioned-review", positioned.Id);
        Assert.Contains("positioned review hints", positioned.Pipeline, StringComparison.Ordinal);
        Assert.Contains("browser", positioned.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("not a full PDF renderer", positioned.UnsupportedScope, StringComparison.Ordinal);
        Assert.Contains("image-placements", positioned.ReviewSignals);
        Assert.Contains("unsafe-link-sanitization", positioned.DiagnosticGuarantees);
        Assert.Contains("no-full-graphics-renderer", positioned.RendererBoundaries);
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfHtmlProfileContracts.Get((PdfHtmlProfile)99));
    }

    [Fact]
    public void PdfHtml_NamedProfiles_ApplyCoherentReviewDefaults() {
        PdfHtmlSaveOptions semantic = PdfHtmlSaveOptions.CreateSemanticProfile(OfficeVisualThemeKind.TechnicalDocument);
        PdfHtmlSaveOptions positioned = PdfHtmlSaveOptions.CreatePositionedReviewProfile(OfficeVisualThemeKind.Report);

        Assert.Equal(PdfHtmlProfile.Semantic, semantic.Profile);
        Assert.Equal(OfficeVisualThemeKind.TechnicalDocument, semantic.Theme);
        Assert.True(semantic.IncludeDefaultStyles);
        Assert.False(semantic.IncludeLinkAnnotations);
        Assert.False(semantic.IncludeFormWidgets);

        Assert.Equal(PdfHtmlProfile.PositionedReview, positioned.Profile);
        Assert.Equal(OfficeVisualThemeKind.Report, positioned.Theme);
        Assert.True(positioned.IncludeDefaultStyles);
        Assert.True(positioned.IncludeLinkAnnotations);
        Assert.True(positioned.IncludeFormWidgets);
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_ExportsLogicalStructure() {
        byte[] pdf = CreateLogicalSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        PdfHtmlSaveOptions options = PdfHtmlSaveOptions.CreateSemanticProfile(OfficeVisualThemeKind.TechnicalDocument);

        string html = PdfCore.PdfDocumentReadResult.Load(pdf, layoutOptions).ToHtml(options);

        Assert.Contains("<title>Logical PDF sample</title>", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">", html, StringComparison.Ordinal);
        Assert.Contains("class=\"officeimo-html officeimo-pdf-html officeimo-pdf-semantic\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-html-profile=\"pdf-html-semantic\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-html-theme=\"TechnicalDocument\"", html, StringComparison.Ordinal);
        Assert.Contains("body.officeimo-pdf-semantic .pdf-page", html, StringComparison.Ordinal);
        Assert.Contains("<h1>Logical Heading</h1>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Logical readback marker.</p>", html, StringComparison.Ordinal);
        Assert.Contains("<ul data-pdf-list-level=\"1\"><li>Detected logical bullet</li></ul>", html, StringComparison.Ordinal);
        Assert.Contains("<table>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Code</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th class=\"pdf-numeric\" style=\"text-align:right\">Qty</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>A-100</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td class=\"pdf-numeric\" style=\"text-align:right\">2</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td class=\"pdf-numeric\" style=\"text-align:right\">14</td>", html, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(html, "A-100"));
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPageGeometryAndTextBlocks() {
        byte[] pdf = CreateLogicalSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        PdfHtmlSaveOptions options = PdfHtmlSaveOptions.CreatePositionedReviewProfile();

        string html = PdfCore.PdfDocumentReadResult.Load(pdf, layoutOptions).ToHtml(options);

        Assert.Contains("body.officeimo-pdf-positioned .pdf-page", html, StringComparison.Ordinal);
        Assert.Contains("body.officeimo-pdf-positioned table.pdf-table", html, StringComparison.Ordinal);
        Assert.Contains("class=\"officeimo-html officeimo-pdf-html officeimo-pdf-positioned\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\" style=\"width:420pt;height:360pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("<table class=\"pdf-table\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(html, "A-100"));
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_PreservesMatchingTextOutsideDetectedTableBounds() {
        var pdfOptions = new PdfCore.PdfOptions {
            PageWidth = 460,
            PageHeight = 360,
            MarginLeft = 36,
            MarginRight = 36,
            MarginTop = 36,
            MarginBottom = 36,
            DefaultFontSize = 10
        };
        var tableRows = new[] {
            new[] { "Area", "Owner", "Status" },
            new[] { "HTML", "OfficeIMO.Html", "Ready" },
            new[] { "PDF", "OfficeIMO.Html.Pdf", "Stable" }
        };
        var tableStyle = new PdfCore.PdfTableStyle {
            ColumnWidthPoints = new List<double?> { 70, 170, 60 },
            HeaderRowCount = 1,
            CellPaddingX = 6,
            CellPaddingY = 4
        };
        byte[] baselinePdf = PdfCore.PdfDocument.Create(pdfOptions)
            .H1("Positioned geometry")
            .Table(tableRows, style: tableStyle)
            .ToBytes();
        PdfCore.PdfLogicalPage baselinePage = PdfCore.PdfDocumentReadResult.Load(baselinePdf).Pages[0];
        PdfCore.PdfLogicalTable baselineTable = Assert.Single(baselinePage.Tables);
        PdfCore.PdfLogicalTextBlock matchingRow = Assert.Single(
            baselinePage.TextBlocks,
            block => block.Text.Contains("Ready", StringComparison.Ordinal));
        double outsideX = baselineTable.Columns[baselineTable.Columns.Count - 1].To + 18D;
        double outsideY = baselinePage.Height - matchingRow.BaselineY - matchingRow.FontSize;

        byte[] pdf = PdfCore.PdfDocument.Create(pdfOptions)
            .H1("Positioned geometry")
            .Table(tableRows, style: tableStyle)
            .Canvas(canvas => canvas.Text("Ready", outsideX, outsideY, 50D, 18D, fontSize: matchingRow.FontSize))
            .ToBytes();
        PdfCore.PdfLogicalPage page = PdfCore.PdfDocumentReadResult.Load(pdf).Pages[0];
        PdfCore.PdfLogicalTable table = Assert.Single(page.Tables);
        double tableRight = table.Columns[table.Columns.Count - 1].To;
        Assert.Contains(page.TextBlocks, block =>
            block.XStart > tableRight + 1D && block.Text.Contains("Ready", StringComparison.Ordinal));

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(
            PdfHtmlSaveOptions.CreatePositionedReviewProfile());

        Assert.Equal(2, CountOccurrences(html, "Ready"));
    }

    [Fact]
    public void Pdf_ToHtmlResult_PositionedReviewProfile_ReportsExportSummary() {
        byte[] pdf = CreatePdfHtmlSummarySamplePdf("https://example.com/summary");
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(pdf, layoutOptions).ToHtmlResult(options);

        Assert.False(result.Report.HasWarnings);
        Assert.Contains("Logical Heading", result.Value, StringComparison.Ordinal);
        Assert.Equal(PdfHtmlProfile.PositionedReview, result.Summary.Profile);
        Assert.Equal("pdf-html-positioned-review", result.Summary.ProfileId);
        Assert.Equal(1, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 1 }, result.Summary.PageNumbers);
        Assert.True(result.Summary.TextBlockCount > 0);
        Assert.True(result.Summary.HeadingCount > 0);
        Assert.True(result.Summary.ListItemCount > 0);
        Assert.Equal(1, result.Summary.TableCount);
        Assert.True(result.Summary.ImageCount > 0);
        Assert.True(result.Summary.ImagePlacementCount > 0);
        Assert.True(result.Summary.LinkCount > 0);
        Assert.Equal(0, result.Summary.WarningCount);
        Assert.True(result.Summary.EmitsDocumentShell);
        Assert.True(result.Summary.UsesSharedDocumentStyles);
        Assert.Equal(OfficeVisualThemeKind.Report, result.Summary.Theme);
        Assert.Equal(PdfHtmlImageExportMode.EmbeddedDataUri, result.Summary.ImageExportMode);
        Assert.Contains("positioned", result.Summary.FidelityContract, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("not a full PDF renderer", result.Summary.UnsupportedScope, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_PositionedReviewProfile_RendersOutlinesAsNavigationMetadata() {
        byte[] pdf = CreateOutlineSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(pdf, layoutOptions).ToHtmlResult(options);

        Assert.Contains("class=\"pdf-outline\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("aria-label=\"PDF outline\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-outline-count=\"3\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-rendered-outline-count=\"3\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-outline-level=\"1\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-outline-level=\"2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-1\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("id=\"pdf-page-1\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("id=\"pdf-page-2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("Executive summary", result.Value, StringComparison.Ordinal);
        Assert.Contains("Risk posture", result.Value, StringComparison.Ordinal);
        Assert.Contains("Appendix", result.Value, StringComparison.Ordinal);
        Assert.Equal(3, result.Summary.OutlineCount);
        Assert.Equal(3, result.Summary.RenderedOutlineCount);
    }

    [Fact]
    public void Pdf_ToHtmlResult_CanSuppressOutlineNavigation() {
        byte[] pdf = CreateOutlineSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeOutlines = false
        };

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(pdf, layoutOptions).ToHtmlResult(options);

        Assert.DoesNotContain("class=\"pdf-outline\"", result.Value, StringComparison.Ordinal);
        Assert.Equal(3, result.Summary.OutlineCount);
        Assert.Equal(0, result.Summary.RenderedOutlineCount);
    }

    [Fact]
    public void Pdf_ToHtmlResult_ReportsAcroFormXfaAsInertReviewMetadata() {
        byte[] pdf = CreateAcroFormXfaPdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtmlResult(options);

        Assert.Contains("class=\"pdf-xfa-notice\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-xfa-packet-count=\"2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-xfa-packet-names=\"template,datasets\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("does not render or fill XFA", result.Value, StringComparison.Ordinal);
        Assert.True(result.Summary.HasAcroFormXfa);
        Assert.Equal(2, result.Summary.AcroFormXfaPacketCount);
        Assert.Equal(2, result.Summary.AcroFormXfaStreamCount);
        Assert.True(result.Summary.AcroFormXfaPayloadByteCount > 0);
        Assert.Equal(1, result.Summary.WarningCount);
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Report.Warnings, item => item.Code == "AcroFormXfaDetected");
        Assert.Equal("OfficeIMO.Html.Pdf", warning.Converter);
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Contains("does not render or fill XFA", warning.Message, StringComparison.Ordinal);
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void Pdf_ToHtmlResult_SnapshotsConversionReportWhenOptionsAreReused() {
        byte[] imagePdf = CreateImageSamplePdf();
        byte[] textPdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            MaxEmbeddedImageBytes = 0
        };

        PdfHtmlConversionResult imageResult = PdfCore.PdfDocumentReadResult.Load(imagePdf).ToHtmlResult(options);
        PdfCore.PdfConversionWarning warning = Assert.Single(imageResult.Report.Warnings, item => item.Code == "ImageDataTooLarge");
        Assert.Equal("OfficeIMO.Html.Pdf", warning.Converter);
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.True(imageResult.HasLoss);
        Assert.Throws<InvalidOperationException>(() => imageResult.RequireNoLoss());

        using var output = new MemoryStream();
        PdfCore.PdfConversionReport saveReport = PdfCore.PdfDocumentReadResult.Load(imagePdf).SaveAsHtml(output, options);
        Assert.True(saveReport.IsReadOnly);
        Assert.Throws<InvalidOperationException>(() => saveReport.Clear());
        Assert.True(saveReport.HasLoss);
        Assert.NotEmpty(output.ToArray());

        PdfHtmlConversionResult textResult = PdfCore.PdfDocumentReadResult.Load(textPdf).ToHtmlResult(options);

        Assert.Single(imageResult.Report.Warnings, item => item.Code == "ImageDataTooLarge");
        Assert.False(textResult.Report.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtmlResult_PageRanges_PreserveSourcePageCountAndSelectedFormFields() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .TextField("FirstPageField", width: 120, value: "first")
            .PageBreak()
            .TextField("SecondPageField", width: 120, value: "second")
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(pdf);
        PdfHtmlConversionResult result = logical.ToHtmlResult(options);

        Assert.Equal(2, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 2 }, result.Summary.PageNumbers);
        Assert.Equal(1, result.Summary.FormFieldCount);
        Assert.Equal(1, result.Summary.FormWidgetCount);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewFragment_IncludesPositioningCss() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            EmitDocumentShell = false
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.DoesNotContain("<!doctype html>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<style>", html, StringComparison.Ordinal);
        Assert.Contains("body.officeimo-pdf-positioned .pdf-page", html, StringComparison.Ordinal);
        Assert.Contains(".pdf-text {", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_RejectsInvalidNamedThemeWhenSharedStylesAreEnabled() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Theme = (OfficeVisualThemeKind)999,
            IncludeDefaultStyles = true
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options));
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewWithoutDefaultStyles_RetainsOnlyStructuralCss() {
        byte[] pdf = CreateLogicalSamplePdf();
        PdfHtmlSaveOptions options = PdfHtmlSaveOptions.CreatePositionedReviewProfile();
        options.IncludeDefaultStyles = false;

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtmlResult(options);

        Assert.Contains(".pdf-page {", result.Value, StringComparison.Ordinal);
        Assert.Contains(".pdf-text {", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(":root{--officeimo-accent:", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("body.officeimo-pdf-html {", result.Value, StringComparison.Ordinal);
        Assert.False(result.Summary.UsesSharedDocumentStyles);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPositionedImagePlaceholders() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(1, 1), 40, 50, 60, 30))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("body.officeimo-pdf-positioned figure.pdf-image-placeholder", html, StringComparison.Ordinal);
        Assert.Contains("style=\"position:absolute;left:40pt;top:50pt;width:60pt;height:30pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("data-matrix=\"60 0 0 30 40 140\"", html, StringComparison.Ordinal);
        Assert.Contains("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_CanForceImagePlaceholders() {
        byte[] pdf = CreateImageSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>Image:", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_LogicalDocumentPageRanges_UsesUniqueAnchorsForDuplicatePageSelections() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true,
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Repeated Page")
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(pdf);
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(1, 1),
                PdfCore.PdfPageRange.From(1, 1)
            }
        };

        string html = logical.ToHtml(options);

        Assert.Contains("id=\"pdf-page-1-1\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"pdf-page-1-2\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"pdf-page-1\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-1-1\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_EmbedsExtractedImageData() {
        byte[] pdf = CreateImageSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.Contains("<figure class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>Image:", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_ExportsSelectedPagesThroughSameBridgePackage() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("First PDF page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second PDF page"))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.DoesNotContain("First PDF page", html, StringComparison.Ordinal);
        Assert.Contains("Second PDF page", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_DoesNotReapplyRangesAfterLoadingSelection() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("Duplicated selected page"))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(1, 1),
                PdfCore.PdfPageRange.From(1, 1)
            }
        };

        string html = PdfCore.PdfDocument.Load(pdf).ToHtml(options);

        Assert.Equal(2, CountOrdinal(html, "<section class=\"pdf-page\""));
        Assert.Equal(2, CountOrdinal(html, "Duplicated selected page"));
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_FilterAlreadyLoadedLogicalDocument() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("First logical page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second logical page"))
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(pdf);
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        string html = logical.ToHtml(options);

        Assert.DoesNotContain("First logical page", html, StringComparison.Ordinal);
        Assert.Contains("Second logical page", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_AccountsForRotatedPages() {
        byte[] pdf = CreateRotatedLinkAnnotationPdf(90, "https://example.com/rotated");
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\" style=\"width:220pt;height:320pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:38pt;top:40pt;width:22pt;height:140pt\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://example.com/rotated\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_FlipsCoordinatesForRotated180Pages() {
        byte[] pdf = CreateRotatedLinkAnnotationPdf(180, "https://example.com/rotated-180");
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string html = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\" style=\"width:320pt;height:220pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:140pt;top:38pt;width:140pt;height:22pt\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://example.com/rotated-180\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_LinkAnnotations_RenderUnsafeUriAsInertText() {
        const string unsafeUri = "javascript:alert(1)";
        byte[] pdf = CreateLinkAnnotationPdf(unsafeUri);
        var semanticOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            IncludeLinkAnnotations = true
        };
        var positionedOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string semanticHtml = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(semanticOptions);
        string positionedHtml = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtml(positionedOptions);

        Assert.DoesNotContain("<a href=\"javascript:", semanticHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", semanticHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("<a href=\"javascript:", positionedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", positionedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_ReportsActiveActionDiagnosticsWithoutPayloads() {
        byte[] pdf = CreateActiveContentDiagnosticsPdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfCore.PdfDocumentReadResult.Load(pdf).ToHtmlResult(options);

        Assert.True(result.Summary.HasOpenAction);
        Assert.True(result.Summary.HasCatalogActions);
        Assert.True(result.Summary.HasPageActions);
        Assert.True(result.Summary.HasAnnotationActions);
        Assert.True(result.Summary.HasActiveContent);
        Assert.Equal(5, result.Summary.PotentiallyUnsafeActionCount);
        Assert.Equal(2, result.Summary.JavaScriptActionCount);
        Assert.Equal(1, result.Summary.LaunchActionCount);
        Assert.Equal(1, result.Summary.SubmitFormActionCount);
        Assert.Equal(1, result.Summary.CatalogActionCount);
        Assert.Equal(1, result.Summary.PageActionCount);
        Assert.Equal(1, result.Summary.SelectedPageActionCount);
        Assert.Equal(3, result.Summary.AnnotationActionCount);
        Assert.Equal(3, result.Summary.SelectedAnnotationActionCount);
        Assert.DoesNotContain("app.alert", result.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tool.exe", result.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.com/submit", result.Value, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlPdf_BaselineArtifacts_ExposeStableRoundTripShape() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.Pdf." + Guid.NewGuid().ToString("N"));
        string pdfPath = Path.Combine(directory, "practical-html.pdf");
        string htmlPath = Path.Combine(directory, "practical-html-review.html");
        string linkUri = "https://example.com/artifact";
        Directory.CreateDirectory(directory);

        try {
            OfficeIMO.Html.HtmlConversionDocument.Parse(CreatePracticalHtmlSample(linkUri)).SaveAsPdf(pdfPath, new HtmlPdfSaveOptions());
            PdfCore.PdfDocumentReadResult.Load(pdfPath).SaveAsHtml(htmlPath, new PdfHtmlSaveOptions {
                Profile = PdfHtmlProfile.PositionedReview,
                IncludeLinkAnnotations = true
            });

            byte[] pdf = File.ReadAllBytes(pdfPath);
            string html = File.ReadAllText(htmlPath);

            Assert.True(new FileInfo(pdfPath).Length > 0);
            Assert.True(new FileInfo(htmlPath).Length > 0);
            Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount >= 2);
            Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", html, StringComparison.Ordinal);
            Assert.Contains("class=\"pdf-link\"", html, StringComparison.Ordinal);
            Assert.Contains("href=\"" + linkUri + "\"", html, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", html, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static byte[] CreateAcroFormXfaPdf() {
        const string template = "<template/>";
        const string datasets = "<datasets/>";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [] /XFA [(template) 6 0 R (datasets) 7 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /Length " + template.Length + " >>",
            "stream",
            template,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Length " + datasets.Length + " >>",
            "stream",
            datasets,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateImageSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(1, 1), 40, 50, 60, 30))
            .ToBytes();
    }

    private static byte[] CreateLinkAnnotationPdf(string uri) {
        string escapedUri = uri.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Unsafe link) /A << /S /URI /URI ({escapedUri}) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateActiveContentDiagnosticsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /Fit] /Names << /JavaScript << /Names [(Open) 6 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /Annots [5 0 R 9 0 R] /AA << /O 7 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Action link) /A << /S /Launch /F (tool.exe) >> /AA << /E 8 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /S /JavaScript /JS (app.alert('catalog')) >>",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('page')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /SubmitForm /F (https://example.com/submit) >>",
            "endobj",
            "9 0 obj",
            "<< /Type /Annot /Subtype /Screen /Rect [40 110 180 150] /A << /S /RichMedia >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static int CountOrdinal(string value, string search) {
        int count = 0;
        int index = 0;
        while (true) {
            index = value.IndexOf(search, index, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            index += search.Length;
        }
    }

    private static byte[] CreateRotatedLinkAnnotationPdf(int rotationDegrees, string uri) {
        string escapedUri = uri.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Rotate {rotationDegrees.ToString(System.Globalization.CultureInfo.InvariantCulture)} /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Rotated link) /A << /S /URI /URI ({escapedUri}) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateLogicalSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical PDF sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
    }

    private static byte[] CreateOutlineSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true,
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Executive summary")
            .Paragraph(paragraph => paragraph.Text("Summary body."))
            .H2("Risk posture")
            .Paragraph(paragraph => paragraph.Text("Risk body."))
            .PageBreak()
            .H1("Appendix")
            .Paragraph(paragraph => paragraph.Text("Appendix body."))
            .ToBytes();
    }

    private static byte[] CreatePdfHtmlSummarySamplePdf(string linkUri) {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "PDF to HTML summary sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, PdfCore.PdfAlign.Left, null, null, 6, 0, null, linkUri, "Summary link")
            .ToBytes();
    }

    private static string CreatePracticalHtmlSample(string linkUri) {
        string pixel = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        return $$"""
<html>
<head>
  <style>
    table { border-collapse: collapse; }
    td, th { border: 1px solid #444; padding: 4px; }
    .page-two { break-before: page; }
  </style>
</head>
<body>
  <h1>Practical HTML</h1>
  <p><a href="{{linkUri}}">Report link</a></p>
  <p><img src="data:image/png;base64,{{pixel}}" alt="Embedded pixel" width="24" height="24"></p>
  <table>
    <tr><th>Area</th><th>Status</th></tr>
    <tr><td>Table marker</td><td>Ready</td></tr>
  </table>
  <section class="page-two"><h2>Second page marker</h2><p>Page break proof.</p></section>
</body>
</html>
""";
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

    private static string GetFieldAppearanceContent(byte[] pdf, string fieldName) {
        var (objects, _) = PdfCore.PdfSyntax.ParseObjects(pdf);
        PdfCore.PdfDictionary fieldObject = Assert.IsType<PdfCore.PdfDictionary>(objects.Values
            .Select(item => item.Value)
            .Single(item => item is PdfCore.PdfDictionary dictionary && dictionary.Get<PdfCore.PdfStringObj>("T")?.Value == fieldName));
        PdfCore.PdfDictionary appearance = Assert.IsType<PdfCore.PdfDictionary>(fieldObject.Items["AP"]);
        PdfCore.PdfReference normalAppearance = Assert.IsType<PdfCore.PdfReference>(appearance.Items["N"]);
        PdfCore.PdfStream stream = Assert.IsType<PdfCore.PdfStream>(objects[normalAppearance.ObjectNumber].Value);
        return Encoding.ASCII.GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
    }

    private static IEnumerable<HtmlRenderVisual> EnumeratePdfSceneVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clip
                ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip
                    ? pathClip.Visuals
                    : visual is HtmlRenderEffectGroup effect
                        ? effect.Visuals
                        : visual is HtmlRenderSemanticGroup semantic
                            ? semantic.Visuals
                            : visual is HtmlRenderLogicalTextGroup logical
                                ? logical.Visuals
                                : visual is HtmlRenderFormField form ? form.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumeratePdfSceneVisuals(children)) yield return child;
        }
    }

    private static string Hex(string text) {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2"));
        }

        return builder.ToString();
    }

    private sealed class PdfTextBounds {
        public PdfTextBounds(double left, double right) {
            Left = left;
            Right = right;
        }

        public double Left { get; }

        public double Right { get; }

        public double Center => (Left + Right) / 2D;
    }

    private static PdfTextBounds FindPdfTextBounds(byte[] pdf, params string[] texts) {
        using PdfPigDocument document = PdfPigDocument.Open(new MemoryStream(pdf));
        var lines = document.GetPage(1)
            .GetWords()
            .GroupBy(word => Math.Round(word.BoundingBox.Bottom, 1))
            .Select(group => group.OrderBy(word => word.BoundingBox.Left).ToList())
            .ToList();

        foreach (var line in lines) {
            for (int index = 0; index <= line.Count - texts.Length; index++) {
                bool matches = true;
                for (int offset = 0; offset < texts.Length; offset++) {
                    if (!string.Equals(line[index + offset].Text, texts[offset], StringComparison.Ordinal)) {
                        matches = false;
                        break;
                    }
                }

                if (matches) {
                    double left = line.Skip(index).Take(texts.Length).Min(word => word.BoundingBox.Left);
                    double right = line.Skip(index).Take(texts.Length).Max(word => word.BoundingBox.Right);
                    return new PdfTextBounds(left, right);
                }
            }
        }

        throw new InvalidOperationException("Could not find rendered PDF text '" + string.Join(" ", texts) + "'.");
    }
}
