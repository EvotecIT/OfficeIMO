using OfficeIMO.Email;

namespace OfficeIMO.Email.Tests;

public sealed class ArchiveLossPolicyTests {
    [Theory]
    [InlineData(EmailConversionLossPolicy.Block)]
    [InlineData(EmailConversionLossPolicy.Warn)]
    [InlineData(EmailConversionLossPolicy.Allow)]
    public async Task SourceMetadataAndPartialContentHonorPolicyBeforeWriting(EmailConversionLossPolicy policy) {
        foreach (string kind in new[] { "Emlx", "Olm", "Mapi", "Partial" }) {
            var document = new EmailDocument { Subject = "Retained archive properties" };
            document.Body.Text = "Body";
            if (kind == "Emlx") document.Properties["Emlx:RawMetadata"] = new byte[] { 1 };
            else if (kind == "Olm") document.Properties["Olm:RawXml"] = "<source/>";
            else if (kind == "Mapi") document.MessageMetadata.IconIndex = 1;
            else document.Properties["Emlx:IsPartial"] = true;
            var writer = new EmailDocumentWriter(new EmailWriterOptions(conversionLossPolicy: policy));
            bool blocked = policy == EmailConversionLossPolicy.Block;
            Assert.Equal(!blocked, writer.AnalyzeConversion(document, EmailFileFormat.Eml).CanWrite);
            foreach (bool asynchronous in new[] { false, true }) {
                using var output = new MemoryStream();
                output.WriteByte(123);
                EmailWriteResult result = asynchronous ? await writer.WriteAsync(document, output) : writer.Write(document, output);
                Assert.Equal(blocked ? EmailConversionLossDisposition.Blocked : EmailConversionLossDisposition.Accepted, result.LossDisposition);
                Assert.Contains(result.Diagnostics, item => item.Severity == (blocked ? EmailDiagnosticSeverity.Error :
                    policy == EmailConversionLossPolicy.Warn ? EmailDiagnosticSeverity.Warning : EmailDiagnosticSeverity.Information));
                if (blocked) Assert.Equal(new byte[] { 123 }, output.ToArray());
                else Assert.True(output.Length > 1);
            }
        }
    }
}
