using System.Text.Json;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Invoice;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class InvoiceCommandTests {
    private static readonly string[] Target = ["--release", "En16931_1_3_16", "--syntax", "Ubl", "--profile", "En16931"];

    [Fact]
    public async Task SourceEditPublishesPreservedXmlAndReportsBlockedSignedInput() {
        using var scope = new Files(); string input = scope.Invoice("input.xml"), output = scope.Path("edited.xml");
        byte[] original = File.ReadAllBytes(input);
        var result = await Run(["invoice", "edit", input, "--output", output, "--number", "EDITED-123", "--issue-date", "2028-02-29"]);
        Assert.Equal(0, result.Code);
        Assert.Equal(original, File.ReadAllBytes(input));
        var edited = InvoiceSourceDocument.Load(output);
        Assert.Equal("EDITED-123", edited.Number); Assert.Equal(new DateTime(2028, 2, 29), edited.IssueDate);
        Assert.Equal(InvoiceSourceDocument.Load(input).PaymentReference, edited.PaymentReference);
        using var json = JsonDocument.Parse(result.Output);
        Assert.Equal("EditSource", json.RootElement.GetProperty("operation").GetString());
        Assert.True(json.RootElement.GetProperty("published").GetBoolean());
        var signed = System.Xml.Linq.XDocument.Parse(Encoding.UTF8.GetString(original));
        signed.Root!.Add(new System.Xml.Linq.XElement("{http://www.w3.org/2000/09/xmldsig#}Signature"));
        File.WriteAllText(input, signed.ToString());
        string blockedOutput = scope.Path("blocked.xml");
        var blocked = await Run(["invoice", "edit", input, "--output", blockedOutput, "--number", "BLOCKED"]);
        Assert.Equal(1, blocked.Code); Assert.False(File.Exists(blockedOutput));
        using var blockedJson = JsonDocument.Parse(blocked.Output);
        Assert.Contains(blockedJson.RootElement.GetProperty("diagnostics").EnumerateArray(), d => d.GetProperty("location").GetString() == "Source.Signature");
    }

    [Fact]
    public async Task BatchSourceEditCreatesXmlOutputsUsingSharedFailurePolicy() {
        using var scope = new Files(); string first = scope.Invoice("first.xml"), second = scope.Invoice("second.xml");
        string outputRoot = scope.Path("outputs");
        var result = await Run(["invoice", "batch", "edit", first, second, "--output-directory", outputRoot, "--buyer-reference", "NEW-BUYER"]);
        Assert.Equal(0, result.Code);
        Assert.Equal("NEW-BUYER", InvoiceSourceDocument.Load(System.IO.Path.Combine(outputRoot, "first.invoice.xml")).BuyerReference);
        Assert.Equal("NEW-BUYER", InvoiceSourceDocument.Load(System.IO.Path.Combine(outputRoot, "second.invoice.xml")).BuyerReference);
        Assert.Equal(2, Directory.GetFiles(outputRoot).Length);
    }

    [Theory]
    [InlineData("--issue-date", "2028-02-30")]
    [InlineData("--number", "invalid\0number")]
    [InlineData("--number", "")]
    [InlineData("--syntax", "Ubl")]
    public async Task InvalidSourceEditOptionsFailBeforeInputCapture(string option, string value) {
        var result = await Run(["invoice", "edit", "absent.xml", "--output", "absent-output.xml", option, value]);
        Assert.Equal(2, result.Code); Assert.Equal(string.Empty, result.Output);
    }

    [Fact]
    public async Task MissingEditsAndEditsOnOtherOperationsAreUsageErrors() {
        Assert.Equal(2, (await Run(["invoice", "edit", "absent.xml", "--output", "absent-output.xml"])).Code);
        Assert.Equal(2, (await Run(["invoice", "inspect", "absent.xml", "--number", "EDITED"])).Code);
    }

    [Fact]
    public async Task InspectionJsonSeparatesCompletedInspectionFromUnrunStandards() {
        using var scope = new Files(); string input = scope.Invoice("input.xml");
        var result = await Run(["invoice", "inspect", input]);
        Assert.Equal(0, result.Code);
        using var json = JsonDocument.Parse(result.Output);
        Assert.Equal("officeimo.invoice.result.v1", json.RootElement.GetProperty("schema").GetString());
        Assert.Equal("NotRun", json.RootElement.GetProperty("standards").GetProperty("businessRulesStatus").GetString());
        Assert.Equal(119m, json.RootElement.GetProperty("model").GetProperty("payableAmount").GetDecimal());
        Assert.False(json.RootElement.GetProperty("published").GetBoolean());
        Assert.Equal(string.Empty, result.Error);
    }

    [Fact]
    public async Task ConversionCreatesParseableOutputAndRefusesExistingDestinationAndSource() {
        using var scope = new Files(); string input = scope.Invoice("input.xml"), output = scope.Path("output.xml");
        byte[] source = File.ReadAllBytes(input);
        var converted = await Run(["invoice", "convert", input, "--output", output, .. Target]);
        Assert.Equal(0, converted.Code);
        Assert.Equal(InvoiceSyntax.Ubl, InvoiceParser.Read(File.ReadAllBytes(output)).Declaration.Syntax);
        byte[] originalOutput = File.ReadAllBytes(output);
        Assert.Equal(6, (await Run(["invoice", "convert", input, "--output", output, .. Target])).Code);
        Assert.Equal(originalOutput, File.ReadAllBytes(output));
        Assert.NotEqual(0, (await Run(["invoice", "convert", input, "--output", input, .. Target])).Code);
        Assert.Equal(source, File.ReadAllBytes(input));
        Assert.Equal(2, Directory.GetFiles(scope.Root).Length);
    }

    [Fact]
    public async Task BatchPreflightProtectsAllInputsBeforePublishingAnyOutput() {
        using var scope = new Files(); string input = scope.Invoice("invoice.xml");
        string collided = scope.Invoice("invoice.invoice.xml"); byte[] source = File.ReadAllBytes(collided);
        var result = await Run(["invoice", "batch", "convert", input, collided, "--output-directory", scope.Root, .. Target]);
        Assert.NotEqual(0, result.Code);
        Assert.Equal(source, File.ReadAllBytes(collided));
        Assert.False(File.Exists(scope.Path("invoice.invoice.invoice.xml")));
    }

    [Fact]
    public async Task BatchReportsIndividualValidationFailuresAndEnforcesCaptureLimits() {
        using var scope = new Files(); string good = scope.Invoice("good.xml"), bad = scope.Path("bad.xml");
        File.WriteAllText(bad, "<broken>");
        var result = await Run(["invoice", "batch", "validate", bad, good]);
        Assert.Equal(1, result.Code);
        using var json = JsonDocument.Parse(result.Output);
        var reports = json.RootElement.GetProperty("results");
        Assert.Equal(2, reports.GetArrayLength()); Assert.False(reports[0].GetProperty("succeeded").GetBoolean());
        Assert.True(reports[1].GetProperty("succeeded").GetBoolean());
        var stopped = await Run(["invoice", "batch", "validate", bad, good, "--stop-on-failure"]);
        using var stoppedJson = JsonDocument.Parse(stopped.Output);
        Assert.Equal(1, stoppedJson.RootElement.GetProperty("results").GetArrayLength());
        Assert.NotEqual(0, (await Run(["invoice", "validate", good, "--max-input-bytes", "1"])).Code);
    }

    [Fact]
    public async Task HybridCliEmbedsConvertedXmlAndRendersLocalizedInvoice() {
        using var scope = new Files(); string input = scope.Invoice("input.xml"), output = scope.Path("hybrid.pdf");
        var result = await Run(["invoice", "hybrid", input, "--output", output,
            "--release", "FacturX_1_09_2_Zugferd_2_5_2", "--syntax", "Cii", "--profile", "En16931",
            "--language", "pl-PL", "--font", System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "SourceSerif4-Regular.otf"),
            "--columns", "Item,Quantity,Unit,NetPrice,Vat,NetAmount", "--unit-display", "Description", "--payment-display", "Description",
            "--modern", "--compact-details", "--page-identity"]);
        Assert.Equal(0, result.Code);
        byte[] pdf = File.ReadAllBytes(output);
        byte[] xml = Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes;
        Assert.Equal("INV-2026-001", InvoiceParser.Read(xml).Invoice.Number);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("119,00 EUR", text);
        Assert.Contains("sztuka", text);
        Assert.Contains("Przelew SEPA", text);
        Assert.Contains("Strona 1 /", text);
    }

    [Theory]
    [InlineData("--release", "0")]
    [InlineData("--release", " 0")]
    [InlineData("--profile", "")]
    [InlineData("--unknown", "value")]
    public async Task InvalidOptionsReturnUsageWithoutExecuting(string option, string value) {
        var result = await Run(["invoice", "convert", "absent.xml", option, value]);
        Assert.Equal(2, result.Code); Assert.Equal(string.Empty, result.Output);
    }

    [Theory]
    [InlineData("--columns", "Item,NetAmount,Item")]
    [InlineData("--columns", "Quantity,NetAmount")]
    [InlineData("--unit-display", "Code,Description")]
    public async Task InvalidPresentationOptionsFailBeforeInputCapture(string option, string value) {
        var result = await Run(["invoice", "render", "absent.xml", "--output", "absent.pdf",
            "--release", "En16931_1_3_16", "--syntax", "Cii", "--profile", "En16931", option, value]);
        Assert.Equal(2, result.Code); Assert.Equal(string.Empty, result.Output);
    }

    [Fact]
    public async Task CancellationIsReportedWithoutPublishing() {
        using var scope = new Files(); string input = scope.Invoice("input.xml"), output = scope.Path("output.xml");
        var result = await Run(["invoice", "convert", input, "--output", output, .. Target], new(true));
        Assert.Equal(130, result.Code); Assert.False(File.Exists(output));
    }

    private static async Task<(int Code, string Output, string Error)> Run(string[] args, CancellationToken token = default) {
        using var input = new MemoryStream(); using var output = new MemoryStream(); using var error = new StringWriter();
        int code = await OfficeImoToolApp.RunAsync(args, input, output, error, token);
        return (code, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }
    private sealed class Files : IDisposable {
        internal string Root { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OfficeIMO-Invoice-CLI-" + Guid.NewGuid().ToString("N"));
        internal Files() => Directory.CreateDirectory(Root);
        internal string Path(string name) => System.IO.Path.Combine(Root, name);
        internal string Invoice(string name) {
            string path = Path(name); File.WriteAllBytes(path, InvoiceSerializer.Write(InvoiceFixture.Create(),
                new(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931))); return path;
        }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
