using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Validation.Tests;

public class Fa3SchemaValidatorTests {
    private static string SchemaDirectory => Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "Schemas");
    private static byte[] Example(int number) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "example-" + number.ToString("D2") + ".xml"));
    public static IEnumerable<object[]> OfficialExamples => Enumerable.Range(1, 26).Select(number => new object[] { number });

    [Fact]
    public void CheckedInCorpusKeepsTheAuthorityArchiveBytesAcrossPlatforms() {
        using var manifest = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "manifest.json")));
        Assert.Equal("41EBD3C57144951C65D68A36FBE433285B5791A86A8BD46CB059503E3F8B1E10", manifest.RootElement.GetProperty("archiveSha256").GetString());
        var records = manifest.RootElement.GetProperty("fixtures").EnumerateArray().ToArray();
        Assert.Equal(26, records.Length);
        Assert.Equal(26, records.Select(record => record.GetProperty("file").GetString()).Distinct().Count());
        foreach (var record in records) {
            byte[] xml = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", record.GetProperty("file").GetString()!));
            Assert.Equal(record.GetProperty("sha256").GetString(), Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(xml)));
        }
    }

    [Theory]
    [MemberData(nameof(OfficialExamples))]
    public void OfficialIndependentExamplesPassTheExactPinnedSchema(int number) {
        var validator = new Fa3SchemaValidator(Fa3SchemaBundle.LoadDirectory(SchemaDirectory));
        Fa3SchemaValidationResult result = validator.Validate(Example(number));
        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Diagnostics.Select(item => item.Message)));
        Assert.Equal(InvoiceValidationStatus.Passed, result.Status);
        Assert.Equal(Fa3SchemaBundle.SchemaSha256, result.SchemaSha256);
        Assert.Equal(Example(number).Length, result.InputByteCount);
        Assert.Equal(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(Example(number))), result.InputSha256);
    }

    [Fact]
    public void RequiredFieldErrorsAndDocumentControlledSchemaLocationCannotBypassPinnedValidation() {
        var validator = new Fa3SchemaValidator(Fa3SchemaBundle.LoadDirectory(SchemaDirectory));
        XNamespace ns = Fa3InvoiceReader.NamespaceUri;
        XDocument document = XDocument.Parse(System.Text.Encoding.UTF8.GetString(Example(1)));
        document.Root!.SetAttributeValue(XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance") + "schemaLocation", Fa3InvoiceReader.NamespaceUri + " https://unreachable.invalid/accept.xsd");
        document.Descendants(ns + "P_2").Single().Remove();
        Fa3SchemaValidationResult result = validator.Validate(System.Text.Encoding.UTF8.GetBytes(document.ToString()));
        Assert.Equal(InvoiceValidationStatus.Invalid, result.Status);
        Assert.Contains(result.Diagnostics, item => item.Code == "XSD" && item.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.False(validator.Validate(System.Text.Encoding.UTF8.GetBytes("<unknown />")).IsValid);
        Assert.False(validator.Validate(System.Text.Encoding.UTF8.GetBytes("<!DOCTYPE Faktura [<!ENTITY x 'expanded'>]><Faktura>&x;</Faktura>")).IsValid);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => validator.Validate(Example(1), cancelled.Token));
    }

    [Fact]
    public void TamperedImportsAreRejectedAndExistingBundleKeepsItsVerifiedSnapshot() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-FA3-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            foreach (string path in Directory.GetFiles(SchemaDirectory, "*.xsd")) File.Copy(path, Path.Combine(directory, Path.GetFileName(path)));
            Fa3SchemaBundle bundle = Fa3SchemaBundle.LoadDirectory(directory);
            File.AppendAllText(Path.Combine(directory, "KodyKrajow_v10-0E.xsd"), " ");
            Assert.Throws<InvalidDataException>(() => Fa3SchemaBundle.LoadDirectory(directory));
            Assert.True(new Fa3SchemaValidator(bundle).Validate(Example(1)).IsValid);
        } finally { Directory.Delete(directory, true); }
    }
}
