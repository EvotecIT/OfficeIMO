using OfficeIMO.Email;
using OfficeIMO.Reader;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderOftTests {
    [Fact]
    public void EmailReaderProjectsOutlookTemplateThroughExistingEmailSurface() {
        var source = new EmailDocument { Subject = "Reader template" };
        source.Body.Text = "Reusable reader body";
        byte[] bytes = source.ToBytes(EmailFileFormat.OutlookTemplate);

        OfficeDocumentReadResult result = OfficeDocumentReader.Default.ReadDocument(bytes, "reader-template.oft");

        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Contains("officeimo.email.outlooktemplate", result.CapabilitiesUsed);
        Assert.Contains(result.Metadata, item => item.Name == "Format" && item.Value == "OutlookTemplate");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Reusable reader body", StringComparison.Ordinal));
    }

    [Fact]
    public void BuiltInEmailCapabilityAdvertisesOft() {
        ReaderHandlerCapability capability = Assert.Single(OfficeDocumentReader.Default.GetCapabilities(), item =>
            item.Id == "officeimo.reader.email");

        Assert.Contains(".oft", capability.Extensions);
        Assert.Equal(ReaderInputKind.Email, OfficeDocumentReader.Default.DetectKind("template.oft"));
    }
}
