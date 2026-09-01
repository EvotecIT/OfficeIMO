using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void RepeatedReferencesToOneFileSpecificationCountAsOneCarrier() {
        byte[] source = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] duplicatedReference = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            entries.Items.Add(new PdfStringObj("duplicate-content-credential.c2pa"));
            entries.Items.Add(candidate);
            return security.InfoObjectNumber;
        });
        var inspectOptions = new OfficeProvenanceOptions {
            MaxAssetBytes = duplicatedReference.LongLength + 1,
            MaxManifestBytes = 512,
            MaxExpandedContainerBytes = 1024 * 1024,
            MaxCarriers = 1
        };

        OfficeProvenanceReport report = PdfProvenance.Inspect(duplicatedReference, inspectOptions);
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxAssetBytes = duplicatedReference.LongLength + 1;
        removalOptions.Limits.MaxManifestBytes = 512;
        removalOptions.Limits.MaxExpandedContainerBytes = 1024 * 1024;
        removalOptions.Limits.MaxCarriers = 1;
        OfficeProvenanceRemovalResult removal = PdfProvenance.Remove(duplicatedReference, removalOptions);

        Assert.Single(report.Evidence);
        Assert.Equal(2, PdfAttachmentExtractor.ExtractAttachments(duplicatedReference).Count);
        Assert.Single(removal.Before.Evidence);
        Assert.Single(removal.Changes);
        Assert.Empty(removal.After.Evidence);
        Assert.Equal("keep.txt", Assert.Single(PdfAttachmentExtractor.ExtractAttachments(removal.ToArray())).FileName);
    }

    [Fact]
    public void EncryptedObjectStreamConsumesDecodedBudgetOnlyAfterDecryption() {
        (byte[] pdf, long requiredDecodedBytes) = CreateEncryptedObjectStreamPdf();
        var options = new PdfLoadOptions {
            Password = "open",
            Limits = new PdfReadLimits {
                MaxDecodedStreamBytes = 1024,
                MaxTotalDecodedStreamBytes = requiredDecodedBytes
            }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        Assert.Equal(requiredDecodedBytes, document.DecodedStreamBytes);
        Assert.Single(document.Pages);
    }

    private static (byte[] Pdf, long RequiredDecodedBytes) CreateEncryptedObjectStreamPdf() {
        const string packedPage = "3 0 << /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>";
        byte[] objectStreamData = PdfEncoding.Latin1GetBytes(packedPage);
        byte[] objectStreamBody = PdfObjectBytes.WrapStreamBody(
            "<< /Type /ObjStm /N 1 /First 4 /Length " + objectStreamData.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            objectStreamData);
        var sourceObjects = new[] {
            PdfObjectBytes.WrapIndirectObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n"),
            PdfObjectBytes.WrapIndirectObject(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>\n"),
            PdfObjectBytes.WrapIndirectObject(3, "null\n"),
            PdfObjectBytes.WrapIndirectObject(4, objectStreamBody)
        };
        var encryptionOptions = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.LegacyRc4
        };

        using PdfEncryptionAssembly encryption = PdfStandardSecurityWriter.Encrypt(sourceObjects, encryptionOptions);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n%\xE2\xE3\xCF\xD3\n");
        var offsets = new long[7];
        for (int i = 0; i < encryption.Objects.Count; i++) {
            int objectNumber = i + 1;
            offsets[objectNumber] = output.Position;
            byte[] value = encryption.Objects[i];
            output.Write(value, 0, value.Length);
        }

        const int xrefObjectNumber = 6;
        offsets[xrefObjectNumber] = output.Position;
        byte[] xrefData = BuildEncryptedObjectStreamXref(offsets);
        string fileId = BitConverter.ToString(encryption.FileId).Replace("-", string.Empty);
        WriteAscii(output,
            xrefObjectNumber.ToString(CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 7 /W [1 4 2] /Index [0 7] /Root 1 0 R /Encrypt " +
            encryption.EncryptionObjectNumber.ToString(CultureInfo.InvariantCulture) +
            " 0 R /ID [<" + fileId + "> <" + fileId + ">] /Length " +
            xrefData.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(xrefData, 0, xrefData.Length);
        WriteAscii(output, "\nendstream\nendobj\nstartxref\n" + offsets[xrefObjectNumber].ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return (output.ToArray(), xrefData.LongLength + objectStreamData.LongLength);
    }

    private static byte[] BuildEncryptedObjectStreamXref(long[] offsets) {
        using var data = new MemoryStream();
        WriteXrefEntry(data, 0, 0, ushort.MaxValue);
        WriteXrefEntry(data, 1, offsets[1], 0);
        WriteXrefEntry(data, 1, offsets[2], 0);
        WriteXrefEntry(data, 2, 4, 0);
        WriteXrefEntry(data, 1, offsets[4], 0);
        WriteXrefEntry(data, 1, offsets[5], 0);
        WriteXrefEntry(data, 1, offsets[6], 0);
        return data.ToArray();
    }

    private static void WriteXrefEntry(Stream output, byte type, long field1, ushort field2) {
        output.WriteByte(type);
        output.WriteByte((byte)(field1 >> 24));
        output.WriteByte((byte)(field1 >> 16));
        output.WriteByte((byte)(field1 >> 8));
        output.WriteByte((byte)field1);
        output.WriteByte((byte)(field2 >> 8));
        output.WriteByte((byte)field2);
    }
}
