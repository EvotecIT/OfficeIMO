using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Browser;
using Xunit;

namespace OfficeIMO.Browser.Tests;

public sealed class BrowserAssetsTests {
    [Theory]
    [InlineData("officeimo.js")]
    [InlineData("officeimo-xlsx.js")]
    [InlineData("officeimo-csv.js")]
    [InlineData("officeimo.mjs")]
    [InlineData("officeimo-xlsx.mjs")]
    [InlineData("officeimo-csv.mjs")]
    [InlineData("officeimo-datatables.js")]
    [InlineData("officeimo-datatables.mjs")]
    public void EmbeddedAssetAndHashedNameDescribeExactShippedBytes(string name) {
        BrowserAsset asset = new[] { BrowserAssets.Script, BrowserAssets.XlsxScript, BrowserAssets.CsvScript,
            BrowserAssets.Module, BrowserAssets.XlsxModule, BrowserAssets.CsvModule,
            BrowserAssets.DataTablesScript, BrowserAssets.DataTablesModule }.Single(a => a.FileName == name);
        byte[] expected = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", name));
        byte[] actual = new UTF8Encoding(false).GetBytes(asset.Content);
        Assert.Equal(expected, actual);
        using var sha = SHA256.Create();
        string hash = BitConverter.ToString(sha.ComputeHash(expected)).Replace("-", "").ToLowerInvariant().Substring(0, 16);
        Assert.Equal(hash, asset.ContentHash);
        Assert.Equal(Path.GetFileNameWithoutExtension(name) + "." + hash + Path.GetExtension(name), asset.HashedFileName);
    }
}
