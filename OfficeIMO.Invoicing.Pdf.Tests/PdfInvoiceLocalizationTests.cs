using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public class PdfInvoiceLocalizationTests {
    [Fact]
    public void BuiltInLanguagePacksCoverEveryGeneratedLabel() {
        foreach (string culture in new[] { "en-US", "de-DE", "pl-PL", "fr-FR", "es-ES", "it-IT", "nl-NL", "pt-PT", "cs-CZ", "sk-SK" }) {
            InvoicePdfLanguagePack pack = InvoicePdfLanguagePack.ForCulture(culture);
            foreach (InvoicePdfText text in Enum.GetValues(typeof(InvoicePdfText)))
                Assert.False(string.IsNullOrWhiteSpace(pack[text]));
        }
    }

    [Theory]
    [InlineData("es-MX", "Factura", "Importe pendiente")]
    [InlineData("it-CH", "Fattura", "Importo dovuto")]
    [InlineData("nl-BE", "Factuur", "Te betalen")]
    [InlineData("pt-BR", "Fatura", "Montante a pagar")]
    [InlineData("cs-CZ", "Faktura", "Částka k úhradě")]
    [InlineData("sk-SK", "Faktúra", "Suma na úhradu")]
    public void AddedBuiltInLanguagesSelectByLanguageAndExposeTranslatedCoreLabels(
        string culture, string invoice, string amountDue) {
        InvoicePdfLanguagePack pack = InvoicePdfLanguagePack.ForCulture(culture);

        Assert.Equal(invoice, pack[InvoicePdfText.Invoice]);
        Assert.Equal(amountDue, pack[InvoicePdfText.AmountDue]);
    }

    [Theory]
    [InlineData("es-ES", "Factura INV-2026-001", "Importe pendiente")]
    [InlineData("it-IT", "Fattura INV-2026-001", "Importo dovuto")]
    [InlineData("nl-NL", "Factuur INV-2026-001", "Te betalen")]
    [InlineData("pt-PT", "Fatura INV-2026-001", "Montante a pagar")]
    [InlineData("cs-CZ", "Faktura INV-2026-001", "Částka k úhradě")]
    [InlineData("sk-SK", "Faktúra INV-2026-001", "Suma na úhradu")]
    public void AddedBuiltInLanguagesRenderLocalizedInvoiceLabels(
        string culture, string heading, string amountDue) {
        InvoicePdfLayoutOptions layout = InvoicePdfLayoutOptions.ForCultures(culture);

        byte[] pdf = PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract(), layout)
            .ToPresentationPdfBytes(MultilingualOptions());
        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains(heading, text, StringComparison.Ordinal);
        Assert.Contains(amountDue, text, StringComparison.Ordinal);
    }

    [Fact]
    public void MultilingualLongMixedScriptLayoutPreservesExactXmlAndVisibleMeaning() {
        Invoice invoice = InvoiceFixture.Create();
        const string mixed = "Usługa wielojęzyczna — 日本語 — العربية — हिन्दी";
        invoice.Lines[0].Name = mixed;
        invoice.Lines[0].Description = string.Join(" ", Enumerable.Repeat(mixed, 45)) + " FINAL-MIXED-SCRIPT-MARKER";
        invoice.Lines[0].UnitPrice = 1234.56m;

        InvoicePdfLayoutOptions layout = InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB");
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract(), layout);
        byte[] xml = snapshot.ToXmlBytes();

        layout.Languages.Clear();
        layout.DateFormat = "yyyy";
        byte[] pdf = snapshot.ToPdfBytes(MultilingualOptions());

        PdfDocument loaded = PdfDocument.Load(pdf);
        Assert.Equal(xml, Assert.Single(loaded.Attachments.Extract()).Bytes);
        InvoiceReadResult embedded = InvoiceParser.Read(xml);
        Assert.Equal(invoice.Lines[0].Name, embedded.Invoice.Lines[0].Name);
        Assert.Equal(invoice.Lines[0].Description, embedded.Invoice.Lines[0].Description);

        PdfReadDocument visible = PdfReadDocument.Open(pdf);
        string text = visible.ExtractText();
        Assert.True(visible.Pages.Count > 1);
        Assert.Contains("Faktura / Invoice " + invoice.Number, text, StringComparison.Ordinal);
        Assert.Contains("Sprzedawca / Seller", text, StringComparison.Ordinal);
        Assert.Contains("1234,56", text, StringComparison.Ordinal);
        Assert.Contains("Usługa wielojęzyczna", text, StringComparison.Ordinal);
        Assert.Contains("日本語", text, StringComparison.Ordinal);
        Assert.Contains("العربية", text, StringComparison.Ordinal);
        Assert.Contains("हिन्दी", text, StringComparison.Ordinal);
        Assert.Contains("FINAL-MIXED-SCRIPT-MARKER", text, StringComparison.Ordinal);
        Assert.True(loaded.Diagnostics().EmbeddedFontCount >= 4);

        WriteEvidence("multilingual-mixed-script", pdf, xml);
    }

    [Fact]
    public void CustomLanguagePackOverridesOwnedLabelsAndFallsBackForTheRest() {
        InvoicePdfLanguagePack custom = InvoicePdfLanguagePack.Create("sv-SE",
            new Dictionary<InvoicePdfText, string> { [InvoicePdfText.Invoice] = "Faktura" });
        var layout = new InvoicePdfLayoutOptions();
        layout.Languages.Clear();
        layout.Languages.Add(custom);
        layout.FormattingCulture = custom.Culture;

        byte[] pdf = PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract(), layout).ToPdfBytes(MultilingualOptions());
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Faktura INV-2026-001", text, StringComparison.Ordinal);
        Assert.Contains("Seller", text, StringComparison.Ordinal);
    }

    private static PdfOptions MultilingualOptions() {
        string fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!;
        Assert.NotNull(fontPath);
        byte[] baseFont = File.ReadAllBytes(fontPath);
        return new PdfOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, baseFont, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, baseFont, "OfficeIMO Source Serif")
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Japanese", File.ReadAllBytes(Asset("Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "NotoSansJP-OfficeIMO-Common.ttf"))),
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Arabic", File.ReadAllBytes(Asset("Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "NotoSansArabic-Regular.ttf"))),
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Devanagari", File.ReadAllBytes(Asset("OfficeIMO.Drawing.Tests", "TestAssets", "NotoSansDevanagari-Regular.ttf")))
            }));
    }

    private static string Asset(params string[] parts) {
        DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) directory = directory.Parent;
        Assert.NotNull(directory);
        return Path.Combine(new[] { directory!.FullName }.Concat(parts).ToArray());
    }

    private static void WriteEvidence(string name, byte[] pdf, byte[] xml) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output!);
        File.WriteAllBytes(Path.Combine(output!, name + ".pdf"), pdf);
        File.WriteAllBytes(Path.Combine(output!, name + ".xml"), xml);
    }

    private static InvoiceXmlOptions Contract() => new(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
        InvoiceSyntax.Cii,
        InvoiceProfile.En16931);
}
