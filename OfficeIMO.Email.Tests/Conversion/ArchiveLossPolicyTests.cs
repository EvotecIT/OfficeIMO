using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed class ArchiveLossPolicyTests {
    [Theory]
    [InlineData(EmailConversionLossPolicy.Block)]
    [InlineData(EmailConversionLossPolicy.Warn)]
    [InlineData(EmailConversionLossPolicy.Allow)]
    public async Task RegeneratedEmbeddedMessagesHonorPolicyInEveryArtifact(EmailConversionLossPolicy policy) {
        foreach (EmailFileFormat format in new[] { EmailFileFormat.Eml, EmailFileFormat.Emlx,
                     EmailFileFormat.OutlookMsg, EmailFileFormat.OutlookTemplate, EmailFileFormat.Tnef }) {
            foreach (string kind in new[] { "Emlx", "Olm", "Partial", "Mapi", "TnefMessage", "TnefAttachment" }) {
                if (kind == "Mapi" && format != EmailFileFormat.Eml && format != EmailFileFormat.Emlx) continue;
                if (kind.StartsWith("Tnef", StringComparison.Ordinal) &&
                    format != EmailFileFormat.OutlookMsg && format != EmailFileFormat.OutlookTemplate) continue;
                var child = new EmailDocument { Subject = "Nested source" };
                child.Body.Text = "Nested content";
                if (kind == "Emlx") child.Properties["Emlx:Metadata:remote-id"] = "42";
                else if (kind == "Olm") child.Properties["Olm:RawXml"] = "<source/>";
                else if (kind == "Partial") child.Properties["Emlx:IsPartial"] = true;
                else if (kind == "Mapi") child.MessageMetadata.IconIndex = 1;
                else if (kind == "TnefMessage") child.TnefAttributes.Add(new TnefAttribute(
                    TnefAttributeLevel.Message, 0x0006F001, new byte[] { 7, 8 }));
                else {
                    var payload = new EmailAttachment { FileName = "payload.bin", Content = new byte[] { 1 }, Length = 1 };
                    payload.TnefAttributes.Add(new TnefAttribute(TnefAttributeLevel.Attachment, 0x0006F001, new byte[] { 7, 8 }));
                    child.Attachments.Add(payload);
                }
                var parent = new EmailDocument { Subject = "Parent" };
                parent.Attachments.Add(new EmailAttachment { EmbeddedDocument = child, FileName = "child.eml" });
                var root = new EmailDocument { Subject = "Root" };
                root.Attachments.Add(new EmailAttachment { EmbeddedDocument = parent, FileName = "parent.eml" });
                if (format == EmailFileFormat.Emlx) root.Properties["Emlx:Metadata:remote-id"] = "root";
                var options = new EmailWriterOptions(conversionLossPolicy: policy);
                bool blocked = policy == EmailConversionLossPolicy.Block;
                foreach (bool asynchronous in new[] { false, true }) {
                    using var output = new MemoryStream();
                    output.WriteByte(123);
                    EmailWriteResult result;
                    if (format == EmailFileFormat.Emlx) {
                        var writer = new EmailStoreEmlxWriter(new EmailStoreEmlxWriterOptions(options));
                        result = asynchronous ? await writer.WriteAsync(root, output) : writer.Write(root, output);
                    } else {
                        var writer = new EmailDocumentWriter(options);
                        Assert.Equal(!blocked, writer.AnalyzeConversion(root, format).CanWrite);
                        result = asynchronous ? await writer.WriteAsync(root, output, format) : writer.Write(root, output, format);
                    }
                    Assert.Equal(blocked ? EmailConversionLossDisposition.Blocked : EmailConversionLossDisposition.Accepted, result.LossDisposition);
                    Assert.Contains(result.Diagnostics, item => item.Location?.Contains("attachment/0/attachment/0") == true);
                    if (blocked) Assert.Equal(new byte[] { 123 }, output.ToArray());
                    else Assert.True(output.Length > 1);
                }
            }
        }
    }

    [Fact]
    public void ExactEmbeddedMimePayloadReuseDoesNotAnalyzeAnUnusedProjection() {
        var child = new EmailDocument { Subject = "Unused projection" };
        child.Properties["Olm:RawXml"] = "<source/>";
        var root = new EmailDocument { Subject = "Exact embedded bytes" };
        byte[] payload = System.Text.Encoding.ASCII.GetBytes("Subject: Preserved payload\r\n\r\nBody\r\n");
        var attachment = new EmailAttachment {
            EmbeddedDocument = child, Content = payload, Length = payload.Length,
            PreserveMimeHeadersOnWrite = true
        };
        attachment.MimeHeaders.Add(new EmailHeader("Content-Type", "message/rfc822"));
        root.Attachments.Add(attachment);
        using var output = new MemoryStream();
        EmailWriteResult result = new EmailDocumentWriter().Write(root, output);
        Assert.False(result.HasErrors);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_OLM_METADATA_NOT_REPRESENTED");
        Assert.Contains(System.Text.Encoding.ASCII.GetString(payload), System.Text.Encoding.ASCII.GetString(output.ToArray()));
    }

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
