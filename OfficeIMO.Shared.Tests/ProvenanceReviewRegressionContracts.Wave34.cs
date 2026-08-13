using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void OdfMimetypeValidationRejectsLocalExtraFields() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                ZipArchiveEntry mimetype = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
                using Stream target = mimetype.Open();
                WriteAll(target, Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text"));
            }
            package = AddEntryExtraFields(stream.ToArray(), "mimetype", new byte[] { 0xFE, 0xCA, 0, 0 }, Array.Empty<byte>());
        }

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceZip.ValidateMimetypeEntry(
            package, "application/vnd.oasis.opendocument.text", 100));
    }
}
