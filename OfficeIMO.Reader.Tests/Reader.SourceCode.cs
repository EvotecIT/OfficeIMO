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
        Assert.Equal("````powershell\nfunction Get-Fence {\n    '```'\n}\n````\n", chunk.Markdown);
        Assert.Equal((1, 3), (chunk.Location.StartLine, chunk.Location.EndLine));
        OfficeDocumentMetadataEntry language = Assert.Single(result.Metadata, entry => entry.Id == "source-code-language");
        Assert.Equal("powershell", language.Value);
        Assert.Equal("PowerShell", language.Attributes["displayName"]);
        Assert.Equal("extension", language.Attributes["detectedBy"]);
    }

    [Fact]
    public void LongSourceIsChunkedOnLineBoundariesWithContinuousLineNumbers() {
        // Forty short lines with one 700-character line in the middle, which must be split inside the line.
        string line = "int value = 42; // " + new string('x', 60) + "\n";
        string code = string.Concat(Enumerable.Repeat(line, 20)) + "string data = \"" + new string('y', 700) + "\";\n"
            + string.Concat(Enumerable.Repeat(line, 20));

        OfficeDocumentReadResult result = CreateAllReader().ReadDocument(Encoding.UTF8.GetBytes(code), "Values.cs",
            new ReaderOptions { MaxChars = 300 });

        Assert.True(result.Chunks.Count > 3);
        Assert.Equal(code, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.All(result.Chunks, chunk => Assert.StartsWith("```csharp\n", chunk.Markdown));
        for (int index = 1; index < result.Chunks.Count; index++) {
            ReaderChunk previous = result.Chunks[index - 1];
            int expectedStart = previous.Location.EndLine!.Value + (previous.Text.EndsWith("\n", StringComparison.Ordinal) ? 1 : 0);
            Assert.Equal(expectedStart, result.Chunks[index].Location.StartLine);
        }
        Assert.Contains(result.Chunks, chunk => !chunk.Text.EndsWith("\n", StringComparison.Ordinal) && chunk.Location.StartLine == 21);
        Assert.Equal(41, result.Chunks[result.Chunks.Count - 1].Location.EndLine);

        // The document Markdown keeps one closed fence per chunk; no fence line runs into the next one.
        string[] fenceLines = result.Markdown!.Split('\n').Where(text => text.StartsWith("```", StringComparison.Ordinal)).ToArray();
        Assert.Equal(result.Chunks.Count * 2, fenceLines.Length);
        Assert.All(fenceLines.Where((_, index) => index % 2 == 0), text => Assert.Equal("```csharp", text));
        Assert.All(fenceLines.Where((_, index) => index % 2 == 1), text => Assert.Equal("```", text));
    }

    [Fact]
    public void SourceExtensionsKeepTheirHandlerWhenContentDetectionIsPreferred() {
        // Python comments look like Markdown headings to content detection.
        string script = "# Release helper\n# Usage: python tool.py\nprint('ok')\n";

        OfficeDocumentReadResult result = CreateAllReader().ReadDocument(Encoding.UTF8.GetBytes(script), "tool.py",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

        Assert.Contains(OfficeDocumentReaderBuilderSourceCodeExtensions.HandlerId, result.CapabilitiesUsed);
        Assert.StartsWith("```python\n# Release helper", Assert.Single(result.Chunks).Markdown);
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

    [Theory]
    [InlineData("Dockerfile", "# Build stage\nFROM mcr.microsoft.com/dotnet/sdk:10.0\nRUN dotnet build\n", "dockerfile")]
    [InlineData("Makefile", "# Targets\nall:\n\tdotnet build\n", "makefile")]
    [InlineData("Jenkinsfile", "pipeline {\n    agent any\n}\n", "groovy")]
    [InlineData("Gemfile", "source 'https://rubygems.org'\ngem 'rake'\n", "ruby")]
    public void ExtensionlessConventionalFileNamesAreReadAsTheirLanguage(string fileName, string code, string languageId) {
        // Leading '# ' comments make content detection call several of these Markdown, so the name must route them.
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Reader-SourceCode-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, fileName);
            File.WriteAllText(path, code, new UTF8Encoding(false));
            OfficeDocumentReader reader = CreateAllReader();

            foreach (OfficeDocumentReadResult result in new[] {
                         reader.ReadDocument(path),
                         reader.ReadDocument(Encoding.UTF8.GetBytes(code), fileName),
                         reader.ReadDocument(Encoding.UTF8.GetBytes(code), fileName,
                             new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent })
                     }) {
                Assert.Contains(OfficeDocumentReaderBuilderSourceCodeExtensions.HandlerId, result.CapabilitiesUsed);
                Assert.StartsWith("```" + languageId + "\n", Assert.Single(result.Chunks).Markdown);
                OfficeDocumentMetadataEntry language = Assert.Single(result.Metadata, entry => entry.Id == "source-code-language");
                Assert.Equal(languageId, language.Value);
                Assert.Equal("file-name", language.Attributes["detectedBy"]);
            }

            ReaderSourceDocument folderDocument = Assert.Single(reader.EnumerateFolderDocuments(root));
            Assert.True(folderDocument.Parsed);
            Assert.Equal(code, Assert.Single(folderDocument.Chunks).Text);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Utf16SourceIsDecodedFromItsByteOrderMarkOrConfiguredEncoding(bool byteOrderMark) {
        byte[] text = Encoding.Unicode.GetBytes("SELECT N'Zażółć' AS Name;\n");
        byte[] bytes = byteOrderMark ? Encoding.Unicode.GetPreamble().Concat(text).ToArray() : text;

        OfficeDocumentReadResult result = CreateAllReader().ReadDocument(bytes, "query.sql",
            new ReaderOptions { TextEncoding = byteOrderMark ? null : Encoding.Unicode });

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
