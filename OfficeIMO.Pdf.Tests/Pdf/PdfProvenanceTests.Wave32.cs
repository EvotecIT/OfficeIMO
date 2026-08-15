using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ParserAndAttachmentDecodingShareOneExpandedByteBudget() {
        byte[] pdf = CreatePdfWithObjectStreamAndCandidate();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(
            pdf,
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 350 }));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void StructTreeRoleMapDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var roleMap = new PdfDictionary();
            roleMap.Items["AF"] = ArrayWith(candidate);
            var structureTree = new PdfDictionary();
            structureTree.Items["Type"] = new PdfName("StructTreeRoot");
            structureTree.Items["RoleMap"] = AddObject(objects, roleMap);
            catalog.Items["StructTreeRoot"] = AddObject(objects, structureTree);
        });
    }

    [Fact]
    public void ActivePageGroupDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var group = new PdfDictionary();
            group.Items["AF"] = ArrayWith(candidate);
            page.Items["Group"] = AddObject(objects, group);
        });
    }

    [Fact]
    public void ActiveAnnotationAppearanceDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var appearance = new PdfDictionary();
            appearance.Items["AF"] = ArrayWith(candidate);
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Text");
            annotation.Items["AP"] = AddObject(objects, appearance);
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Fact]
    public void CatalogExtensionsDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var extensions = new PdfDictionary();
            extensions.Items["AF"] = ArrayWith(candidate);
            catalog.Items["Extensions"] = AddObject(objects, extensions);
        });
    }

    private static void AssertStructuralOwnerRejected(
        Action<Dictionary<int, PdfIndirectObject>, PdfDictionary, PdfReference> mutate) {
        byte[] pdf = RewriteCandidateAssociation(mutate);

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    private static byte[] CreatePdfWithObjectStreamAndCandidate() {
        byte[] manifest = CreateManifestStore();
        string objectStream = "9 0 << /Pad (" + new string('a', 100) + ") >>";
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AF [6 0 R] /Names << /EmbeddedFiles << /Names [(content-credential.c2pa) 6 0 R] >> >> >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>\nendobj\n" +
            "4 0 obj\n<< /Type /ObjStm /N 1 /First 4 /Length " + objectStream.Length + " >>\nstream\n" + objectStream + "\nendstream\nendobj\n" +
            "5 0 obj\n<< /Type /EmbeddedFile /Subtype /application#2Fc2pa /Length " + manifest.Length + " >>\nstream\n");
        output.Write(manifest, 0, manifest.Length);
        WriteAscii(output, "\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /Filespec /F (content-credential.c2pa) /UF (content-credential.c2pa) /AFRelationship /C2PA_Manifest /EF << /F 5 0 R /UF 5 0 R >> >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}
