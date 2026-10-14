using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingOfficePackageRelationshipIdentityTests {
    [Theory]
    [InlineData("word")]
    [InlineData("powerpoint")]
    public void SerializedPackageRootPayloadRelationshipCountsItsPartOnce(string format) {
        byte[] bytes = SerializedPackage(format);
        using (var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)) {
            Assert.Single(archive.Entries.Where(entry => entry.FullName.Contains("/embeddings/")));
            var relationships = archive.Entries.Where(entry => entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
                .SelectMany(entry => { using Stream stream = entry.Open(); return XDocument.Load(stream).Root!.Elements().ToArray(); });
            string target = Assert.Single(relationships, element => (string?)element.Attribute("Id") == "rPayload")
                .Attribute("Target")!.Value;
            Assert.StartsWith("/", target);
        }
        Assert.Equal(1, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
        using var source = new MemoryStream(bytes); source.Position = 7;
        Assert.Equal(1, OfficePackageSecurityInspector.Inspect(source).EmbeddedPayloadPartCount);
        Assert.Equal(7, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("embedded", "/payload/item.bin", 1, false)]
    [InlineData("macro", "/payload/item.bin", 1, false)]
    [InlineData("activex", "/payload/item.bin", 1, false)]
    [InlineData("embedded", "../payload/item.bin", 1, false)]
    [InlineData("embedded", "./../payload/../payload/item.bin?query#fragment", 1, false)]
    [InlineData("embedded", "/payload/./item.bin", 1, false)]
    [InlineData("embedded", "../..//payload/item.bin", 2, false)]
    [InlineData("embedded", "//payload/item.bin", 2, false)]
    [InlineData("embedded", "\\\\payload\\item.bin", 2, false)]
    [InlineData("embedded", "/\\payload/item.bin", 2, false)]
    [InlineData("embedded", "https://example.test/payload/item.bin", 2, false)]
    [InlineData("embedded", "file:///payload/item.bin", 2, false)]
    [InlineData("embedded", "C:/payload/item.bin", 2, false)]
    [InlineData("embedded", "/payload/item.bin", 1, true)]
    public void CanonicalClassificationDeduplicatesOnlyResolvableInternalPackageTargets(
        string kind, string target, int expected, bool external) {
        string contentType = kind == "macro" ? "application/vnd.ms-office.vbaProject"
            : kind == "activex" ? "application/vnd.ms-office.activeX" : "application/vnd.ms-office.oleObject";
        string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
            (kind == "macro" ? "vbaProject" : kind == "activex" ? "control" : "oleObject");
        byte[] bytes;
        using (var memory = new MemoryStream()) {
            using (var archive = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true)) {
                Write(archive, "[Content_Types].xml", new XElement("Types",
                    new XElement("Override", new XAttribute("PartName", "/payload/item.bin"),
                        new XAttribute("ContentType", contentType))).ToString());
                Write(archive, "payload/item.bin", "inert");
                var relationship = new XElement("Relationship", new XAttribute("Id", "rPayload"),
                    new XAttribute("Type", relationshipType), new XAttribute("Target", target));
                if (external) relationship.Add(new XAttribute("TargetMode", "External"));
                Write(archive, "word/_rels/document.xml.rels", new XElement("Relationships", relationship).ToString());
            }
            bytes = memory.ToArray();
        }
        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(bytes);
        int observed = kind == "macro" ? report.MacroPartCount
            : kind == "activex" ? report.ActiveXPartCount : report.EmbeddedPayloadPartCount;
        // Unresolved internal references retain their conservative separate identity;
        // this contract does not turn an external/escaping target into a package part.
        Assert.Equal(expected, observed);
        Assert.Equal(external ? 1 : 0, report.ExternalRelationshipCount);
    }

    [Theory]
    [InlineData("item one.bin", "item one.bin", "item%20one.bin", 1)]
    [InlineData("item%20one.bin", "item one.bin", "item%20one.bin", 1)]
    [InlineData("item one.bin", "item%20one.bin", "item one.bin", 1)]
    [InlineData("item%20one.bin", "item%20one.bin", "item%2520one.bin", 1)]
    [InlineData("item%2520one.bin", "item%2520one.bin", "item%252520one.bin", 1)]
    [InlineData("item one.bin", "item one.bin", "item%2520one.bin", 2)]
    [InlineData("item%20one.bin", "item%20one.bin", "item%252520one.bin", 2)]
    [InlineData("café.bin", "caf%C3%A9.bin", "caf%C3%A9.bin", 1)]
    [InlineData("item%2fone.bin", "item%252fone.bin", "item%252fone.bin", 1)]
    [InlineData("item%name.bin", "item%25name.bin", "item%25name.bin", 1)]
    public void EncodedMetadataUsesOneUniquePhysicalIdentityWithoutRepeatedDecoding(
        string physical, string contentTypeName, string target, int expected) {
        byte[] bytes = IdentityPackage(new[] { "word/embeddings/" + physical },
            "/word/embeddings/" + contentTypeName, "/word/embeddings/" + target);
        Assert.Equal(expected, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
        using var source = new MemoryStream(bytes); source.Position = 7;
        Assert.Equal(expected, OfficePackageSecurityInspector.Inspect(source).EmbeddedPayloadPartCount);
        Assert.Equal(7, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("item one.bin", "item%20one.bin", null, "item%20one.bin", 3)]
    [InlineData("item%20one.bin", "item%2520one.bin", null, "item%2520one.bin", 3)]
    [InlineData("item one.bin", "item%20one.bin", "item%20one.bin", null, 3)]
    [InlineData("item one.bin", "item%2520one.bin", null, "item%252520one.bin", 2)]
    public void PhysicalPercentSpellingsRemainDistinctAndAmbiguousAliasesStayUnresolved(
        string first, string second, string? contentTypeName, string? target, int expected) {
        byte[] bytes = IdentityPackage(new[] { "word/embeddings/" + first, "word/embeddings/" + second },
            contentTypeName == null ? null : "/word/embeddings/" + contentTypeName,
            target == null ? null : "/word/embeddings/" + target);
        Assert.Equal(expected, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
    }

    [Theory]
    [InlineData("embedded")]
    [InlineData("macro")]
    [InlineData("activex")]
    public void EncodedRelationshipIdentityIsSharedByAllContentClassifications(string kind) {
        byte[] bytes = IdentityPackage(new[] { "payload/item one.bin" },
            "/payload/item one.bin", "/payload/item%20one.bin", kind);
        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(bytes);
        Assert.Equal(1, kind == "macro" ? report.MacroPartCount
            : kind == "activex" ? report.ActiveXPartCount : report.EmbeddedPayloadPartCount);
    }

    [Theory]
    [InlineData("/payload//item.bin", 2)]
    [InlineData("/payload/item.bin/", 2)]
    [InlineData("/payload/item.bin/.", 2)]
    [InlineData("../payload//item.bin", 2)]
    [InlineData("../payload/item.bin/", 2)]
    [InlineData("/payload/item.bin//../item.bin", 2)]
    [InlineData("/payload/./item.bin", 1)]
    [InlineData("/payload/../payload/item.bin", 1)]
    [InlineData("/payload/%2fitem.bin", 2)]
    [InlineData("/payload%2fitem.bin", 2)]
    [InlineData("/payload%5citem.bin", 2)]
    [InlineData("/payload/item.bin%00", 2)]
    [InlineData("/payload/%2e%2e/payload/item.bin", 2)]
    [InlineData("/payload/item.bin/..", 2)]
    public void InvalidSegmentsCannotBecomeKnownPhysicalPartsThroughNormalization(string target, int expected) {
        byte[] bytes = IdentityPackage(new[] { "payload/item.bin" }, "/payload/item.bin", target);
        Assert.Equal(expected, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
    }

    [Theory]
    [InlineData("/word/embeddings/missing.bin", "/word/embeddings/missing.bin")]
    [InlineData("/word/embeddings/item.bin", "/word/embeddings/missing.bin")]
    [InlineData("/word/embeddings/missing.bin", "/word/embeddings/item.bin")]
    public void MissingMetadataReferencesRetainTheirSeparateConservativeIdentity(string contentTypeName, string target) {
        byte[] bytes = IdentityPackage(new[] { "word/embeddings/item.bin" }, contentTypeName, target);
        Assert.Equal(2, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
    }

    [Theory]
    [InlineData("/payload/%2f/../item.bin", 2)]
    [InlineData("/payload/%5c/../item.bin", 2)]
    [InlineData("/payload/%2e/../item.bin", 2)]
    [InlineData("/payload/%2e%2e/../item.bin", 2)]
    [InlineData("/payload/%00/../item.bin", 2)]
    [InlineData("/payload/%2/../item.bin", 2)]
    [InlineData("../payload/%2f/../item.bin", 2)]
    [InlineData("../payload/%5c/../item.bin", 2)]
    [InlineData("../payload/%2e/../item.bin", 2)]
    [InlineData("../payload/%2e%2e/../item.bin", 2)]
    [InlineData("../payload/%00/../item.bin", 2)]
    [InlineData("../payload/%2/../item.bin", 2)]
    [InlineData("/payload/./temporary/../item.bin", 1)]
    [InlineData("/payload/%252f/../item.bin", 1)]
    [InlineData("/payload/%25/../item.bin", 1)]
    public void OriginalEncodedSegmentsRemainQualifiedBeforeParentTraversal(string target, int expected) {
        byte[] bytes = IdentityPackage(new[] { "payload/item.bin" }, "/payload/item.bin", target);
        Assert.Equal(expected, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
    }

    [Theory]
    [InlineData("%2f", 2)]
    [InlineData("%5c", 2)]
    [InlineData("%2e", 2)]
    [InlineData("%2e%2e", 2)]
    [InlineData("%00", 2)]
    [InlineData("%2", 2)]
    [InlineData("%252f", 1)]
    [InlineData("temporary", 1)]
    public void OriginalSourceDirectorySegmentsCannotDisappearBeforeQualification(string segment, int expected) {
        byte[] bytes = IdentityPackage(new[] { "payload/item.bin" }, "/payload/item.bin", "../item.bin",
            relationshipPartName: "payload/" + segment + "/_rels/document.xml.rels");
        Assert.Equal(expected, OfficePackageSecurityInspector.Inspect(bytes).EmbeddedPayloadPartCount);
    }

    [Theory]
    [InlineData("macro")]
    [InlineData("activex")]
    public void MalformedTraversalSegmentsRemainUnresolvedAcrossRelationshipClassifications(string kind) {
        byte[] bytes = IdentityPackage(new[] { "payload/item.bin" }, "/payload/item.bin",
            "/payload/%2f/../item.bin", kind);
        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(bytes);
        Assert.Equal(2, kind == "macro" ? report.MacroPartCount : report.ActiveXPartCount);
    }

    [Theory]
    [InlineData("embedded", "word/embeddings/item.bin", "/word/embeddings/%2f/../item.bin")]
    [InlineData("macro", "word/vbaProject.bin", "/word/%2f/../vbaProject.bin")]
    [InlineData("activex", "word/activeX/item.bin", "/word/activeX/%2f/../item.bin")]
    public void OriginalContentTypeSegmentsRetainConservativeClassification(
        string kind, string physicalName, string contentTypeName) {
        byte[] bytes = IdentityPackage(new[] { physicalName }, contentTypeName, target: null, kind: kind);
        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(bytes);
        Assert.Equal(2, kind == "macro" ? report.MacroPartCount
            : kind == "activex" ? report.ActiveXPartCount : report.EmbeddedPayloadPartCount);
    }

    private static byte[] IdentityPackage(string[] physicalNames, string? contentTypeName, string? target,
        string kind = "embedded", string relationshipPartName = "word/_rels/document.xml.rels") {
        using var memory = new MemoryStream();
        using (var archive = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true)) {
            var types = new XElement("Types");
            if (contentTypeName != null) types.Add(new XElement("Override",
                new XAttribute("PartName", contentTypeName), new XAttribute("ContentType", kind == "macro"
                    ? "application/vnd.ms-office.vbaProject" : kind == "activex"
                    ? "application/vnd.ms-office.activeX" : "application/vnd.ms-office.oleObject")));
            Write(archive, "[Content_Types].xml", types.ToString());
            foreach (string name in physicalNames) Write(archive, name, "inert");
            if (target != null) Write(archive, relationshipPartName, new XElement("Relationships",
                new XElement("Relationship", new XAttribute("Id", "rPayload"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
                        (kind == "macro" ? "vbaProject" : kind == "activex" ? "control" : "oleObject")),
                    new XAttribute("Target", target))).ToString());
        }
        return memory.ToArray();
    }

    private static byte[] SerializedPackage(string format) {
        if (format == "word") {
            using WordDocument document = WordDocument.Create(); document.AddParagraph("Inert payload");
            AddPayload(document.OpenXmlDocument.MainDocumentPart!);
            return document.ToBytes();
        }
        using var memory = new MemoryStream();
        using (PresentationDocument presentation = PresentationDocument.Create(memory,
            DocumentFormat.OpenXml.PresentationDocumentType.Presentation)) {
            PresentationPart main = presentation.AddPresentationPart();
            main.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
            SlidePart slide = main.AddNewPart<SlidePart>();
            slide.Slide = new DocumentFormat.OpenXml.Presentation.Slide(
                new DocumentFormat.OpenXml.Presentation.CommonSlideData(new DocumentFormat.OpenXml.Presentation.ShapeTree()));
            AddPayload(slide);
        }
        return memory.ToArray();
    }

    private static void AddPayload(OpenXmlPart owner) {
        EmbeddedPackagePart payload = owner.AddNewPart<EmbeddedPackagePart>("application/octet-stream", "rPayload");
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 }, writable: false); payload.FeedData(stream);
    }

    private static void Write(ZipArchive archive, string name, string text) {
        using Stream stream = archive.CreateEntry(name).Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false)); writer.Write(text);
    }
}
