using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public class InvoiceCorpusTests {
    [Fact]
    public void NationalCiusFixturesHavePinnedIdentityAndBothDeclaredSyntaxes() {
        CorpusManifest manifest = LoadManifest();
        Assert.Equal(1, manifest.SchemaVersion);
        Assert.Equal("XRechnung 3.0.2 2026-08-31", manifest.XRechnung.Release);
        Assert.Equal(new[] { "Cii", "Ubl" }, manifest.XRechnung.Syntaxes);
        foreach (CorpusFixture fixture in manifest.XRechnung.Fixtures) {
            byte[] xml = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "KoSIT", fixture.File));
            Assert.Equal(fixture.Sha256, Convert.ToHexString(SHA256.HashData(xml)));
            InvoiceProfileDeclaration declaration = InvoiceProfileDeclaration.Read(xml);
            Assert.Equal(InvoiceProfile.XRechnung, declaration.Profile);
        }
    }

    [InvoiceStandardsTheory]
    [InlineData("minimum.xml", InvoiceProfile.Minimum)]
    [InlineData("basic_wl.xml", InvoiceProfile.BasicWithoutLines)]
    [InlineData("basic.xml", InvoiceProfile.Basic)]
    [InlineData("en16931.xml", InvoiceProfile.En16931)]
    [InlineData("extended.xml", InvoiceProfile.Extended)]
    public async Task IndependentFacturXProducerFixturesMatchTheirPinnedCorpusAndRules(string file, InvoiceProfile profile) {
        CorpusManifest manifest = LoadManifest();
        CorpusFixture fixture = Assert.Single(manifest.FacturX.Fixtures, candidate => candidate.File == file);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_FACTURX_PRODUCER_CORPUS")
            ?? throw new InvalidOperationException("Factur-X producer corpus path is required.");
        byte[] xml = File.ReadAllBytes(Path.Combine(root, file));
        Assert.Equal(fixture.Sha256, Convert.ToHexString(SHA256.HashData(xml)));
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(xml).Profile);
        var validator = new InvoiceValidator(InvoiceRuleBundle.Load(
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE")!,
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PEPPOL_RULES"),
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_FACTURX_RULE_BUNDLE")),
            new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!,
                Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java"));
        InvoiceValidationReport report = await validator.ValidateAsync(xml, InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2);
        Assert.True(report.IsValid, report.SchemaStatus + "/" + report.BusinessRulesStatus + "\n" +
            string.Join("\n", report.Diagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
    }

    private static CorpusManifest LoadManifest() => JsonSerializer.Deserialize<CorpusManifest>(
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "invoice-corpus.json")),
        new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? throw new InvalidDataException("Invoice corpus manifest is empty.");

    private sealed class CorpusManifest {
        public int SchemaVersion { get; set; }
        public CorpusSource FacturX { get; set; } = new();
        public CorpusSource XRechnung { get; set; } = new();
    }
    private sealed class CorpusSource {
        public string? Release { get; set; }
        public string[] Syntaxes { get; set; } = Array.Empty<string>();
        public CorpusFixture[] Fixtures { get; set; } = Array.Empty<CorpusFixture>();
    }
    private sealed class CorpusFixture {
        public string File { get; set; } = string.Empty;
        public string Sha256 { get; set; } = string.Empty;
    }
}
