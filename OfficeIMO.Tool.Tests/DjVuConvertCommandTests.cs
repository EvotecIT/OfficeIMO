using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class DjVuConvertCommandTests {
    [Theory]
    [InlineData(".djvu")]
    [InlineData(".djv")]
    public async Task ConvertPreservesSearchableUnicodeAndSourcePageSize(string extension) {
        string root = CreateRoot();
        try {
            string source = Path.Combine(root, "scan" + extension), destination = Path.Combine(root, "scan.pdf");
            File.Copy(Fixture("unicode.djvu"), source);
            byte[] original = File.ReadAllBytes(source);

            var run = await Run(["convert", source, destination]);

            Assert.Equal((int)OfficeImoToolExitCode.Success, run.Code);
            Assert.Contains(destination, run.Output, StringComparison.Ordinal);
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(destination));
            Assert.Equal((23.04D, 30.72D), Assert.Single(pdf.Pages).GetPageSize());
            Assert.Contains("Zażółć 😀", pdf.ExtractText(), StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(source));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("--require-no-loss")]
    [InlineData("--max-output-bytes")]
    public async Task FailedAcceptanceLeavesExistingDestinationIntact(string option) {
        string root = CreateRoot();
        try {
            string source = Path.Combine(root, "scan.djvu"), destination = Path.Combine(root, "scan.pdf");
            File.Copy(Fixture("noise-small.djvu"), source);
            byte[] original = [1, 2, 3, 4];
            File.WriteAllBytes(destination, original);
            string[] acceptance = option == "--max-output-bytes" ? [option, "16"] : [option];

            var run = await Run(["convert", source, destination, "--force", .. acceptance]);

            Assert.NotEqual((int)OfficeImoToolExitCode.Success, run.Code);
            Assert.Contains(option == "--require-no-loss" ? "djvu.render.short-iw44-edge" : "limit", run.Error, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Empty(Directory.GetFiles(root, ".officeimo-*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task DiagramLayerOptionsAreRejectedForDjVu() {
        string root = CreateRoot();
        try {
            string source = Path.Combine(root, "scan.djvu");
            File.Copy(Fixture("unicode.djvu"), source);
            var run = await Run(["convert", source, "--diagram-layers", "screen"]);
            Assert.Equal((int)OfficeImoToolExitCode.Usage, run.Code);
            Assert.False(File.Exists(Path.ChangeExtension(source, ".pdf")));
        } finally { Directory.Delete(root, true); }
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "DjVu", name);
    private static string CreateRoot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-djvu-cli-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }
    private static async Task<(int Code, string Output, string Error)> Run(string[] args) {
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int code = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (code, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }
}
