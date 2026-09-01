using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void RequiredCacheRevalidationConsumesAdditionalExpandedByteBudget() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("ASCIIHexDecode");
        var stream = new PdfStream(dictionary, Encoding.ASCII.GetBytes("010203>"));
        var budget = new PdfDecodedStreamBudget(new PdfReadLimits {
            MaxDecodedStreamBytes = 16,
            MaxTotalDecodedStreamBytes = 5
        });
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.Equal(new byte[] { 1, 2, 3 }, budget.Decode(stream, objects, maximumRequestedBytes: 16));
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            budget.DecodeRequired(stream, objects, maximumRequestedBytes: 16));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
        Assert.Equal(3, budget.UsedBytes);
    }

    [Fact]
    public void ActiveDocumentSecurityStoreDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => catalog.Items["DSS"] = candidate);
    }

    [Fact]
    public void ActiveEncryptionDictionaryCannotMasqueradeAsAProvenanceFileSpecification() {
        byte[] pdf = CreateEncryptedPdfWithEncryptionDictionaryCandidate();
        var readOptions = new PdfLoadOptions { Password = "open" };

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf, readOptions: readOptions);
        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);

        Assert.False(evidence.IsStructurallyValid);
    }

    private static byte[] CreateEncryptedPdfWithEncryptionDictionaryCandidate() {
        byte[] manifest = CreateManifestStore();
        byte[] embeddedFile = PdfObjectBytes.WrapStreamBody(
            "<< /Type /EmbeddedFile /Subtype /application#2Fc2pa /Length " +
            manifest.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            manifest);
        var sourceObjects = new[] {
            PdfObjectBytes.WrapIndirectObject(1,
                "<< /Type /Catalog /Pages 2 0 R /AF [5 0 R] /Names << /EmbeddedFiles << /Names [(content-credential.c2pa) 5 0 R] >> >> >>\n"),
            PdfObjectBytes.WrapIndirectObject(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>\n"),
            PdfObjectBytes.WrapIndirectObject(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>\n"),
            PdfObjectBytes.WrapIndirectObject(4, embeddedFile)
        };
        var encryptionOptions = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.LegacyRc4
        };

        using PdfEncryptionAssembly encryption = PdfStandardSecurityWriter.Encrypt(sourceObjects, encryptionOptions);
        Assert.Equal(5, encryption.EncryptionObjectNumber);
        var encryptedObjects = encryption.Objects.Select(static value => (byte[])value.Clone()).ToList();
        string encryptionObject = Encoding.ASCII.GetString(encryptedObjects[encryption.EncryptionObjectNumber - 1]);
        int dictionaryEnd = encryptionObject.LastIndexOf(">>", StringComparison.Ordinal);
        Assert.True(dictionaryEnd > 0);
        encryptionObject = encryptionObject.Insert(
            dictionaryEnd,
            " /F (content-credential.c2pa) /UF (content-credential.c2pa) /AFRelationship /C2PA_Manifest /EF << /F 4 0 R /UF 4 0 R >> ");
        encryptedObjects[encryption.EncryptionObjectNumber - 1] = Encoding.ASCII.GetBytes(encryptionObject);

        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
        var offsets = new long[encryptedObjects.Count + 1];
        for (int index = 0; index < encryptedObjects.Count; index++) {
            offsets[index + 1] = output.Position;
            output.Write(encryptedObjects[index], 0, encryptedObjects[index].Length);
        }
        long xrefOffset = output.Position;
        WriteAscii(output, "xref\n0 " + offsets.Length.ToString(CultureInfo.InvariantCulture) + "\n0000000000 65535 f \n");
        for (int objectNumber = 1; objectNumber < offsets.Length; objectNumber++) {
            WriteAscii(output, offsets[objectNumber].ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n \n");
        }
        string fileId = BitConverter.ToString(encryption.FileId).Replace("-", string.Empty);
        WriteAscii(output,
            "trailer\n<< /Size " + offsets.Length.ToString(CultureInfo.InvariantCulture) +
            " /Root 1 0 R /Encrypt " + encryption.EncryptionObjectNumber.ToString(CultureInfo.InvariantCulture) +
            " 0 R /ID [<" + fileId + "> <" + fileId + ">] >>\nstartxref\n" +
            xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return output.ToArray();
    }
}
