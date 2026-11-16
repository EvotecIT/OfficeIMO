using System.Security.Cryptography;
using System.Text;
using OfficeIMO.JavaScript.Assets;
using Xunit;

namespace OfficeIMO.JavaScript.Assets.Tests;

public sealed class JavaScriptAssetsTests {
    private static readonly JavaScriptAsset[] Readable = {
        JavaScriptAssets.Script, JavaScriptAssets.XlsxScript, JavaScriptAssets.CsvScript,
        JavaScriptAssets.Module, JavaScriptAssets.XlsxModule, JavaScriptAssets.CsvModule,
        JavaScriptAssets.PdfScript, JavaScriptAssets.PdfModule, JavaScriptAssets.ExecutionScript, JavaScriptAssets.ExecutionModule,
        JavaScriptAssets.DataTablesScript, JavaScriptAssets.DataTablesModule, JavaScriptAssets.DataTablesAdapterScript, JavaScriptAssets.DataTablesAdapterModule,
        JavaScriptAssets.CanopyXScript, JavaScriptAssets.CanopyXModule, JavaScriptAssets.CanopyXAdapterScript, JavaScriptAssets.CanopyXAdapterModule };

    private static IEnumerable<JavaScriptAsset> All => Readable.Concat(Readable.Select(asset => asset.Minified));

    public static TheoryData<string> Names() {
        var data = new TheoryData<string>();
        foreach (JavaScriptAsset asset in All) data.Add(asset.FileName);
        return data;
    }

    [Theory]
    [MemberData(nameof(Names))]
    public void EmbeddedAssetAndHashedNameDescribeExactShippedBytes(string name) {
        JavaScriptAsset asset = All.Single(a => a.FileName == name);
        byte[] expected = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", name));
        byte[] actual = new UTF8Encoding(false).GetBytes(asset.Content);
        Assert.Equal(expected, actual);
        using var sha = SHA256.Create();
        string hash = BitConverter.ToString(sha.ComputeHash(expected)).Replace("-", "").ToLowerInvariant().Substring(0, 16);
        Assert.Equal(hash, asset.ContentHash);
        Assert.Equal(Path.GetFileNameWithoutExtension(name) + "." + hash + Path.GetExtension(name), asset.HashedFileName);
    }

    [Fact]
    public void EveryShippedBundleHasOneAssetAndMinifiedVariantsAreTheirOwn() {
        string[] shipped = Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Assets")).Select(Path.GetFileName).OrderBy(n => n, StringComparer.Ordinal).ToArray()!;
        Assert.Equal(shipped, All.Select(a => a.FileName).OrderBy(n => n, StringComparer.Ordinal).ToArray());
        foreach (JavaScriptAsset asset in Readable) {
            Assert.Same(asset.Minified, asset.Minified.Minified);
            Assert.True(asset.Minified.Content.Length < asset.Content.Length, asset.FileName);
        }
    }
}
