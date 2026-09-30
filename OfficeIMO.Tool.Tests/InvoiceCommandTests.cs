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
            "--language", "pl-PL", "--font", System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "SourceSerif4-Regular.otf")]);
        Assert.Equal(0, result.Code);
        byte[] pdf = File.ReadAllBytes(output);
        byte[] xml = Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes;
        Assert.Equal("INV-2026-001", InvoiceParser.Read(xml).Invoice.Number);
        Assert.Contains("119,00 EUR", PdfReadDocument.Open(pdf).ExtractText());
    }

    [Theory]
    [InlineData("--release", "0")]
    [InlineData("--profile", "")]
    [InlineData("--unknown", "value")]
    public async Task InvalidOptionsReturnUsageWithoutExecuting(string option, string value) {
        var result = await Run(["invoice", "convert", "absent.xml", option, value]);
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
