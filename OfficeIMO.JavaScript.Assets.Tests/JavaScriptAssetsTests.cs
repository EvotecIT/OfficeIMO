using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.JavaScript.Assets;
using Xunit;

namespace OfficeIMO.JavaScript.Assets.Tests;

public sealed class JavaScriptAssetsTests {
    [Theory]
    [InlineData("officeimo.js")]
    [InlineData("officeimo-xlsx.js")]
    [InlineData("officeimo-csv.js")]
    [InlineData("officeimo.mjs")]
    [InlineData("officeimo-xlsx.mjs")]
    [InlineData("officeimo-csv.mjs")]
    [InlineData("officeimo-pdf.js")]
    [InlineData("officeimo-pdf.mjs")]
    [InlineData("officeimo-execution.js")]
    [InlineData("officeimo-execution.mjs")]
    [InlineData("officeimo-datatables.js")]
    [InlineData("officeimo-datatables.mjs")]
    [InlineData("officeimo-canopyx.js")]
    [InlineData("officeimo-canopyx.mjs")]
    public void EmbeddedAssetAndHashedNameDescribeExactShippedBytes(string name) {
        JavaScriptAsset asset = new[] { JavaScriptAssets.Script, JavaScriptAssets.XlsxScript, JavaScriptAssets.CsvScript,
            JavaScriptAssets.Module, JavaScriptAssets.XlsxModule, JavaScriptAssets.CsvModule,
            JavaScriptAssets.DataTablesScript, JavaScriptAssets.DataTablesModule, JavaScriptAssets.PdfScript, JavaScriptAssets.PdfModule,
            JavaScriptAssets.CanopyXScript, JavaScriptAssets.CanopyXModule,
            JavaScriptAssets.ExecutionScript, JavaScriptAssets.ExecutionModule }.Single(a => a.FileName == name);
        byte[] expected = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", name));
        byte[] actual = new UTF8Encoding(false).GetBytes(asset.Content);
        Assert.Equal(expected, actual);
        using var sha = SHA256.Create();
        string hash = BitConverter.ToString(sha.ComputeHash(expected)).Replace("-", "").ToLowerInvariant().Substring(0, 16);
        Assert.Equal(hash, asset.ContentHash);
        Assert.Equal(Path.GetFileNameWithoutExtension(name) + "." + hash + Path.GetExtension(name), asset.HashedFileName);
    }
}
