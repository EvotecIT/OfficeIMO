using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentHardeningTests {
    [Fact]
    public void RejectsDuplicateAndCaseAmbiguousArchiveEntries() {
        OdtDocument created = OdtDocument.Create();
        byte[] valid = created.ToBytes();
        byte[] duplicate = RewritePackage(valid, additions: new[] { new OdfTestPackageEntry("content.xml", Encoding.UTF8.GetBytes("<duplicate/>")) });
        byte[] ambiguous = RewritePackage(valid, additions: new[] { new OdfTestPackageEntry("Content.xml", Encoding.UTF8.GetBytes("<ambiguous/>")) });

        Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(duplicate)));
        Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(ambiguous)));
        Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(new byte[12])));
    }

    [Fact]
    public void RejectsDtdEntitiesAndXmlBeyondConfiguredDepth() {
        OdtDocument created = OdtDocument.Create();
        byte[] valid = created.ToBytes();
        const string entityXml = "<?xml version=\"1.0\"?><!DOCTYPE office:document-content [<!ENTITY x \"expanded\">]><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.4\"><office:body><office:text><text:p>&x;</text:p></office:text></office:body></office:document-content>";
        string nested = string.Concat(Enumerable.Repeat("<text:span>", 20)) + "value" + string.Concat(Enumerable.Repeat("</text:span>", 20));
        string deepXml = "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.4\"><office:body><office:text><text:p>" + nested + "</text:p></office:text></office:body></office:document-content>";

        OdtDocument entity = OdtDocument.Load(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new OdfTestPackageEntry("content.xml", Encoding.UTF8.GetBytes(entityXml)) })));
        Assert.Throws<InvalidDataException>(() => entity.ContentBlocks.ToArray());

        OdtDocument deep = OdtDocument.Load(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new OdfTestPackageEntry("content.xml", Encoding.UTF8.GetBytes(deepXml)) })),
            new OdfLoadOptions { MaxXmlDepth = 8 });
        Assert.Throws<InvalidDataException>(() => deep.ContentBlocks.ToArray());
    }

    [Fact]
    public void RejectsBrokenAndEncryptedManifestsBeforeEditing() {
        OdtDocument created = OdtDocument.Create();
        byte[] valid = created.ToBytes();
        XDocument broken = ReadPackageXml(valid, "META-INF/manifest.xml");
        XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        broken.Root!.Elements(manifest + "file-entry").First(element => (string?)element.Attribute(manifest + "full-path") == "/")
            .SetAttributeValue(manifest + "media-type", OdfMediaTypes.Spreadsheet);
        Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new OdfTestPackageEntry("META-INF/manifest.xml", SaveXml(broken)) }))));

        XDocument encrypted = ReadPackageXml(valid, "META-INF/manifest.xml");
        XElement contentEntry = encrypted.Root!.Elements(manifest + "file-entry")
            .First(element => (string?)element.Attribute(manifest + "full-path") == "content.xml");
        contentEntry.Add(new XElement(manifest + "encryption-data",
            new XElement(manifest + "algorithm", new XAttribute(manifest + "algorithm-name", "urn:example:cipher"))));
        Assert.Throws<OdfEncryptedPackageException>(() => OdtDocument.Load(new MemoryStream(RewritePackage(valid,
            replacements: new[] { new OdfTestPackageEntry("META-INF/manifest.xml", SaveXml(encrypted)) }))));
    }

    [Fact]
    public void PreservesUnchangedSignatureEntriesAndRequiresExplicitInvalidation() {
        OdtDocument created = OdtDocument.Create();
        byte[] signedBytes = RewritePackage(created.ToBytes(), additions: new[] {
            new OdfTestPackageEntry("META-INF/documentsignatures.xml", Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><signatures/>"))
        });

        OdtDocument signed = OdtDocument.Load(new MemoryStream(signedBytes));
        Assert.Contains("META-INF/documentsignatures.xml", signed.PackageEntries);
        byte[] unchanged = signed.ToBytes();
        Assert.True(ContainsEntry(unchanged, "META-INF/documentsignatures.xml"));

        signed.Metadata.Title = "Changed";
        Assert.Throws<InvalidOperationException>(() => signed.ToBytes());
        byte[] unsigned = signed.ToBytes(new OdfSaveOptions { SignatureHandling = OdfSignatureHandling.RemoveInvalidated });
        Assert.False(ContainsEntry(unsigned, "META-INF/documentsignatures.xml"));
    }

    [Fact]
    public void CompatibilityProfileRewriteInvalidatesSignaturesBeforeSerialization() {
        OdtDocument created = OdtDocument.Create();
        byte[] odf13 = created.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
        byte[] signedBytes = RewritePackage(odf13, additions: new[] {
            new OdfTestPackageEntry("META-INF/documentsignatures.xml", Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><signatures/>"))
        });

        OdtDocument signed = OdtDocument.Load(new MemoryStream(signedBytes));

        Assert.Throws<InvalidOperationException>(() => signed.ToBytes());
        byte[] unsigned = signed.ToBytes(new OdfSaveOptions { SignatureHandling = OdfSignatureHandling.RemoveInvalidated });
        Assert.False(ContainsEntry(unsigned, "META-INF/documentsignatures.xml"));
        Assert.Equal(OdfVersion.V1_4, signed.Version);
    }

    [Fact]
    public void MalformedFormulaCorpusReturnsErrorsWithinConfiguredBounds() {
        OdsDocument document = OdsDocument.Create();
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

    [Theory]
    [InlineData("of:=----------------1")]
    [InlineData("of:=1^1^1^1^1^1^1^1^1")]
    public void FormulaSyntaxDepthBoundsRecursiveUnaryAndPowerOperators(string formula) {
        OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data");

        OdsFormulaEvaluationResult result = OdsFormulaEvaluator.EvaluateExpression(document, "Data", formula,
            new OdsFormulaEvaluationOptions { MaximumDependencyDepth = 4 });

        Assert.False(result.Success);
        Assert.Contains("syntax depth", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DetectsInvalidRepeatsAndStyleCyclesWithoutUnboundedTraversal() {
        OdsDocument created = OdsDocument.Create();
        created.AddSheet("Data");
        XDocument content = XDocument.Parse(Encoding.UTF8.GetString(created.Package.GetRequiredEntry("content.xml").GetBytesForSave()));
        XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        content.Descendants(table + "table-row").First().SetAttributeValue(table + "number-rows-repeated", "0");
        OdsDocument invalid = OdsDocument.Load(new MemoryStream(RewritePackage(created.ToBytes(),
            replacements: new[] { new OdfTestPackageEntry("content.xml", SaveXml(content)) })));
        Assert.Contains(invalid.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODS100");

        OdtDocument styled = OdtDocument.Create();
        OdfStyle first = styled.Styles.CreateNamed("First", OdfStyleFamily.Paragraph, "Second");
        styled.Styles.CreateNamed("Second", OdfStyleFamily.Paragraph, "First");
        Assert.Equal(2, styled.Styles.Resolve(first).Count);
        Assert.Contains(styled.Diagnostics, diagnostic => diagnostic.Id == "ODF203" && diagnostic.Message.IndexOf("cycle", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void TruncatedMediaRemainsOpaqueAndNeverExecutesOrFetchesContent() {
        byte[] truncated = { 0x89, 0x50, 0x4E, 0x47, 0x00 };
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Image").AddImage(truncated, "broken.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));

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
            OdtDocument document = OdtDocument.Create();
            document.AddParagraph(expected);
            OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.Equal(expected, reopened.ContentBlocks[0].Paragraph!.Text);
        }
    }

    private static byte[] RewritePackage(byte[] sourceBytes, IReadOnlyList<OdfTestPackageEntry>? replacements = null,
        IReadOnlyList<OdfTestPackageEntry>? additions = null) =>
        OdfTestPackageRewriter.Rewrite(sourceBytes, replacements, additions);

    private static XDocument ReadPackageXml(byte[] package, string path) {
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using Stream entry = archive.GetEntry(path)!.Open();
        return XDocument.Load(entry);
    }

    private static byte[] SaveXml(XDocument document) {
        using var output = new MemoryStream(); document.Save(output); return output.ToArray();
    }

    private static bool ContainsEntry(byte[] package, string path) {
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.GetEntry(path) != null;
    }
}
