using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentHardeningTests {
    [Fact]
    public void RejectsDuplicateAndCaseAmbiguousArchiveEntries() {
        using OdtDocument created = OdtDocument.Create();
        byte[] valid = created.ToBytes();
        byte[] duplicate = RewritePackage(valid, additions: new[] { new ArchiveItem("content.xml", Encoding.UTF8.GetBytes("<duplicate/>")) });
        byte[] ambiguous = RewritePackage(valid, additions: new[] { new ArchiveItem("Content.xml", Encoding.UTF8.GetBytes("<ambiguous/>")) });

        Assert.Throws<InvalidDataException>(() => OdtDocument.Open(new MemoryStream(duplicate)));
        Assert.Throws<InvalidDataException>(() => OdtDocument.Open(new MemoryStream(ambiguous)));
        Assert.Throws<InvalidDataException>(() => OdtDocument.Open(new MemoryStream(new byte[12])));
    }

    [Fact]
    public void RejectsDtdEntitiesAndXmlBeyondConfiguredDepth() {
        using OdtDocument created = OdtDocument.Create();
        byte[] valid = created.ToBytes();
        const string entityXml = "<?xml version=\"1.0\"?><!DOCTYPE office:document-content [<!ENTITY x \"expanded\">]><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.4\"><office:body><office:text><text:p>&x;</text:p></office:text></office:body></office:document-content>";
        string nested = string.Concat(Enumerable.Repeat("<text:span>", 20)) + "value" + string.Concat(Enumerable.Repeat("</text:span>", 20));
        string deepXml = "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.4\"><office:body><office:text><text:p>" + nested + "</text:p></office:text></office:body></office:document-content>";

        using OdtDocument entity = OdtDocument.Open(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new ArchiveItem("content.xml", Encoding.UTF8.GetBytes(entityXml)) })));
        Assert.Throws<InvalidDataException>(() => entity.ContentBlocks.ToArray());

        using OdtDocument deep = OdtDocument.Open(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new ArchiveItem("content.xml", Encoding.UTF8.GetBytes(deepXml)) })),
            new OdfOpenOptions { MaxXmlDepth = 8 });
        Assert.Throws<InvalidDataException>(() => deep.ContentBlocks.ToArray());
    }

    [Fact]
    public void RejectsBrokenAndEncryptedManifestsBeforeEditing() {
        using OdtDocument created = OdtDocument.Create();
        byte[] valid = created.ToBytes();
        XDocument broken = ReadPackageXml(valid, "META-INF/manifest.xml");
        XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        broken.Root!.Elements(manifest + "file-entry").First(element => (string?)element.Attribute(manifest + "full-path") == "/")
            .SetAttributeValue(manifest + "media-type", OdfMediaTypes.Spreadsheet);
        Assert.Throws<InvalidDataException>(() => OdtDocument.Open(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new ArchiveItem("META-INF/manifest.xml", SaveXml(broken)) }))));

        XDocument encrypted = ReadPackageXml(valid, "META-INF/manifest.xml");
        XElement contentEntry = encrypted.Root!.Elements(manifest + "file-entry")
            .First(element => (string?)element.Attribute(manifest + "full-path") == "content.xml");
        contentEntry.Add(new XElement(manifest + "encryption-data",
            new XElement(manifest + "algorithm", new XAttribute(manifest + "algorithm-name", "urn:example:cipher"))));
        Assert.Throws<OdfEncryptedPackageException>(() => OdtDocument.Open(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new ArchiveItem("META-INF/manifest.xml", SaveXml(encrypted)) }))));
    }

    [Fact]
    public void PreservesUnchangedSignatureEntriesAndRequiresExplicitInvalidation() {
        using OdtDocument created = OdtDocument.Create();
        byte[] signedBytes = RewritePackage(created.ToBytes(), additions: new[] {
            new ArchiveItem("META-INF/documentsignatures.xml", Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><signatures/>"))
        });

        using OdtDocument signed = OdtDocument.Open(new MemoryStream(signedBytes));
        Assert.Contains("META-INF/documentsignatures.xml", signed.PackageEntries);
        byte[] unchanged = signed.ToBytes();
        Assert.True(ContainsEntry(unchanged, "META-INF/documentsignatures.xml"));

        signed.Metadata.Title = "Changed";
        Assert.Throws<InvalidOperationException>(() => signed.ToBytes());
        byte[] unsigned = signed.ToBytes(new OdfSaveOptions { SignatureHandling = OdfSignatureHandling.RemoveInvalidated });
        Assert.False(ContainsEntry(unsigned, "META-INF/documentsignatures.xml"));
    }

    [Fact]
    public void MalformedFormulaCorpusReturnsErrorsWithinConfiguredBounds() {
        using OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data");
        string[] malformed = {
            "of:=", "of:=(", "of:=1+", "of:=1+@", "of:=[.A0]", "of:=[.A999999999999999999999999999999999]",
            "of:=SUM([.A1:.A999999])", "of:=UNKNOWN(1)", "of:=\"unterminated", "of:=((((((((1))))))))"
        };
        var options = new OdsFormulaEvaluationOptions {
            MaximumFormulaCharacters = 64,
            MaximumRangeCells = 10,
            MaximumDependencyDepth = 4,
            MaximumOperations = 50
        };
        foreach (string formula in malformed) {
            OdsFormulaEvaluationResult result = OdsFormulaEvaluator.EvaluateExpression(document, "Data", formula, options);
            Assert.False(result.Success, formula);
        }
        OdsFormulaEvaluationResult oversized = OdsFormulaEvaluator.EvaluateExpression(document, "Data", "of:=" + new string('1', 100), options);
        Assert.False(oversized.Success);
        Assert.Contains("character limit", oversized.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DetectsInvalidRepeatsAndStyleCyclesWithoutUnboundedTraversal() {
        using OdsDocument created = OdsDocument.Create();
        created.AddSheet("Data");
        XDocument content = XDocument.Parse(Encoding.UTF8.GetString(created.Package.GetRequiredEntry("content.xml").GetBytesForSave()));
        XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        content.Descendants(table + "table-row").First().SetAttributeValue(table + "number-rows-repeated", "0");
        using OdsDocument invalid = OdsDocument.Open(new MemoryStream(RewritePackage(created.ToBytes(),
            replacements: new[] { new ArchiveItem("content.xml", SaveXml(content)) })));
        Assert.Contains(invalid.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODS100");

        using OdtDocument styled = OdtDocument.Create();
        OdfStyle first = styled.Styles.CreateNamed("First", OdfStyleFamily.Paragraph, "Second");
        styled.Styles.CreateNamed("Second", OdfStyleFamily.Paragraph, "First");
        Assert.Equal(2, styled.Styles.Resolve(first).Count);
        Assert.Contains(styled.Diagnostics, diagnostic => diagnostic.Id == "ODF101" && diagnostic.Message.Contains("cycle", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void TruncatedMediaRemainsOpaqueAndNeverExecutesOrFetchesContent() {
        byte[] truncated = { 0x89, 0x50, 0x4E, 0x47, 0x00 };
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Image").AddImage(truncated, "broken.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        Assert.Equal(truncated, Assert.Single(reopened.ContentBlocks[0].Paragraph!.Images).GetImageBytes());
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void DeterministicTextCorpusRoundTripsWhitespaceAndUnicode() {
        var random = new Random(20260710);
        char[] alphabet = "abc XYZ\t\né中".ToCharArray();
        for (int iteration = 0; iteration < 64; iteration++) {
            int length = random.Next(0, 160);
            var builder = new StringBuilder(length);
            for (int index = 0; index < length; index++) builder.Append(alphabet[random.Next(alphabet.Length)]);
            string expected = builder.ToString();
            using OdtDocument document = OdtDocument.Create();
            document.AddParagraph(expected);
            using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));
            Assert.Equal(expected, reopened.ContentBlocks[0].Paragraph!.Text);
        }
    }

    private static byte[] RewritePackage(byte[] sourceBytes, IReadOnlyList<ArchiveItem>? replacements = null,
        IReadOnlyList<ArchiveItem>? additions = null) {
        var replacementMap = (replacements ?? Array.Empty<ArchiveItem>()).ToDictionary(item => item.Name, StringComparer.Ordinal);
        using var output = new MemoryStream();
        using (var sourceStream = new MemoryStream(sourceBytes, writable: false))
        using (var source = new ZipArchive(sourceStream, ZipArchiveMode.Read))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries) {
                byte[] bytes = replacementMap.TryGetValue(entry.FullName, out ArchiveItem? replacement)
                    ? replacement.Bytes : ReadEntry(entry);
                WriteEntry(target, entry.FullName, bytes);
            }
            foreach (ArchiveItem item in additions ?? Array.Empty<ArchiveItem>()) WriteEntry(target, item.Name, item.Bytes);
        }
        return output.ToArray();
    }

    private static XDocument ReadPackageXml(byte[] package, string path) {
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using Stream entry = archive.GetEntry(path)!.Open();
        return XDocument.Load(entry);
    }

    private static byte[] ReadEntry(ZipArchiveEntry entry) {
        using Stream input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output); return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(name, name == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
        using Stream output = entry.Open(); output.Write(bytes, 0, bytes.Length);
    }

    private static byte[] SaveXml(XDocument document) {
        using var output = new MemoryStream(); document.Save(output); return output.ToArray();
    }

    private static bool ContainsEntry(byte[] package, string path) {
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.GetEntry(path) != null;
    }

    private sealed class ArchiveItem {
        internal ArchiveItem(string name, byte[] bytes) { Name = name; Bytes = bytes; }
        internal string Name { get; }
        internal byte[] Bytes { get; }
    }
}
