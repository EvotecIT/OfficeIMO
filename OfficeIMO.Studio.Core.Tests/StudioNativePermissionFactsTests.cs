using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioNativePermissionFactsTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-permissions-" + Guid.NewGuid().ToString("N"));
    public StudioNativePermissionFactsTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }
    public void Dispose() => Directory.Delete(_root, true);

    private string Source(PdfStandardPermissions permissions) {
        string path = Path.Combine(_root, "restricted.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)))
            .Security.Encrypt(new PdfStandardEncryptionOptions("reader") { OwnerPassword = "owner", AllowedPermissions = permissions }).Pdf);
        return path;
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task PrintingFactsDescribeTheAuthenticatedResolutionLimit(bool highQuality, bool owner) {
        var permissions = PdfStandardPermissions.Print | (highQuality ? PdfStandardPermissions.HighQualityPrint : PdfStandardPermissions.None);
        using var document = await StudioNativeDocument.OpenAsync(Source(permissions), owner ? "owner" : "reader", default);
        var facts = await document.GetInfoAsync(default);
        Assert.True(facts.CanPrint);
        Assert.Equal(owner || highQuality ? null : (double?)150, facts.MaxPrintDpi);
        Assert.Equal(!owner, facts.RequiresOwnerPasswordForProtection);
    }

    [Fact]
    public async Task AssemblyOnlyDocumentCanRotateAndInsertBlankWithoutCopyPermission() {
        using var document = await StudioNativeDocument.OpenAsync(Source(PdfStandardPermissions.AssembleDocument), "reader", default);
        var facts = await document.GetInfoAsync(default);
        Assert.True(facts.CanRotatePages);
        Assert.True(facts.CanInsertBlankPages);
        Assert.False(facts.CanExtractPages);
        await document.InsertBlankPageAsync(1, 300, 400, default);
        Assert.Equal(2, (await document.GetInfoAsync(default)).Pages.Count);
        await document.RotatePagesAsync([1, 2], 90, default);
        Assert.All((await document.GetInfoAsync(default)).Pages, page => Assert.Equal(90, page.Rotation));
    }
}
