#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.Text.Json;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvBrowserExportVectorsTests {
    [Fact]
    public void BrowserAndDotNetShareQuotingDialectBomAndInjectionVectors() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "TestAssets", "browser-exports.json")));
        foreach (JsonElement vector in document.RootElement.GetProperty("cases").EnumerateArray())
            Assert.Equal(BrowserCsvVectorContract.Expected(vector), BrowserCsvVectorContract.Write(vector));
    }
}
#endif
