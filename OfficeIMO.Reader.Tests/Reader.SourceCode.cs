using System.Text;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderSourceCodeTests {
    private static OfficeDocumentReader CreateAllReader() =>
        new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

    [Fact]
    public void SourceFilesKeepVerbatimTextAndGetALanguageTaggedFenceAndMetadata() {
        // A UTF-8 BOM, CRLF line breaks and an embedded triple-backtick string all occur in real PowerShell files.
        string script = "function Get-Fence {\r\n    '```'\r\n}\r\n";
        byte[] bytes = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes(script)).ToArray();

        OfficeDocumentReadResult result = CreateAllReader().ReadDocument(bytes, "Get-Fence.ps1");

        Assert.Equal(ReaderInputKind.Text, result.Kind);
        Assert.Contains(OfficeDocumentReaderBuilderSourceCodeExtensions.HandlerId, result.CapabilitiesUsed);
        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal("function Get-Fence {\n    '```'\n}\n", chunk.Text);
        Assert.Equal("````powershell\nfunction Get-Fence {\n    '```'\n}\n````", chunk.Markdown);
        Assert.Equal((1, 3), (chunk.Location.StartLine, chunk.Location.EndLine));
        OfficeDocumentMetadataEntry language = Assert.Single(result.Metadata, entry => entry.Id == "source-code-language");
        Assert.Equal("powershell", language.Value);
        Assert.Equal("PowerShell", language.Attributes["displayName"]);
        Assert.Equal("extension", language.Attributes["detectedBy"]);
    }

    [Fact]
    public void LongSourceIsChunkedOnLineBoundariesWithContinuousLineNumbers() {
        string line = "int value = 42; // " + new string('x', 60) + "\n";
        string code = string.Concat(Enumerable.Repeat(line, 40));

        OfficeDocumentReadResult result = CreateAllReader().ReadDocument(Encoding.UTF8.GetBytes(code), "Values.cs",
            new ReaderOptions { MaxChars = 300 });

        Assert.True(result.Chunks.Count > 1);
        Assert.Equal(code, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.All(result.Chunks, chunk => Assert.EndsWith("\n", chunk.Text));
        Assert.All(result.Chunks, chunk => Assert.StartsWith("```csharp\n", chunk.Markdown));
        for (int index = 1; index < result.Chunks.Count; index++) {
            Assert.Equal(result.Chunks[index - 1].Location.EndLine + 1, result.Chunks[index].Location.StartLine);
        }
        Assert.Equal(40, result.Chunks[result.Chunks.Count - 1].Location.EndLine);
    }

    [Fact]
    public void ExtensionlessScriptsAreRoutedByTheirShebangLine() {
        // Without the shebang route the '# ' comment line makes content detection call this Markdown.
        string script = "#!/usr/bin/env python3\n# Release helper\nprint('ok')\n";
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Reader-SourceCode-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "release");
            File.WriteAllText(path, script, new UTF8Encoding(false));
            OfficeDocumentReader reader = CreateAllReader();

            foreach (OfficeDocumentReadResult result in new[] {
                         reader.ReadDocument(path),
                         reader.ReadDocument(Encoding.UTF8.GetBytes(script), "release")
                     }) {
                Assert.Contains(OfficeDocumentReaderBuilderSourceCodeExtensions.HandlerId, result.CapabilitiesUsed);
                Assert.StartsWith("```python\n#!/usr/bin/env python3", Assert.Single(result.Chunks).Markdown);
                OfficeDocumentMetadataEntry language = Assert.Single(result.Metadata, entry => entry.Id == "source-code-language");
                Assert.Equal("python", language.Value);
                Assert.Equal("shebang", language.Attributes["detectedBy"]);
            }
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Utf16SourceIsDecodedFromItsByteOrderMark() {
        byte[] bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("SELECT N'Zażółć' AS Name;\n")).ToArray();

        OfficeDocumentReadResult result = CreateAllReader().ReadDocument(bytes, "query.sql");

        Assert.Equal("SELECT N'Zażółć' AS Name;\n", Assert.Single(result.Chunks).Text);
        Assert.Equal("sql", Assert.Single(result.Metadata, entry => entry.Id == "source-code-language").Value);
    }

    [Fact]
    public void BinaryPayloadWithASourceExtensionIsRefused() {
        byte[] bytes = { 0x4D, 0x5A, 0x90, 0x00, 0x03, 0x00, 0x00, 0x00, 0x04, 0x00 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateAllReader().ReadDocument(bytes, "payload.cs"));
        Assert.Contains("binary", exception.Message, StringComparison.Ordinal);
    }
}
