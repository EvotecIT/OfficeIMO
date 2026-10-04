namespace OfficeIMO.Email.Store.Tests;

public sealed class AppleSourceLossTests {
    [Theory]
    [InlineData("Olm:Vendor", "EMAIL_OLM_METADATA_NOT_REPRESENTED")]
    [InlineData("Emlx:RawMetadata", "EMAIL_EMLX_METADATA_NOT_REPRESENTED")]
    [InlineData("Emlx:IsPartial", "EMAIL_EMLX_PARTIAL_CONTENT")]
    public void PstWriterReportsArchiveLossAndStrictModeKeepsDestinationAbsent(string property, string code) {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Email.AppleLoss." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var document = new EmailDocument { Subject = "Retained message" };
        document.Properties[property] = property == "Emlx:IsPartial" ? true : "opaque value";
        try {
            string accepted = Path.Combine(root, "accepted.pst");
            using (var writer = EmailStorePstWriter.Create(accepted)) {
                writer.AddItem(writer.AddFolder("Inbox"), document);
                Assert.Contains(writer.Complete().Diagnostics, diagnostic => diagnostic.Code == code);
            }
            using (var session = EmailStoreSession.Open(accepted))
                Assert.Equal(document.Subject, session.ReadSummary(Assert.Single(session.EnumerateItems())).Subject);
            string blocked = Path.Combine(root, "blocked.pst");
            using (var writer = EmailStorePstWriter.Create(blocked, new EmailStorePstWriterOptions(failOnDataLoss: true))) {
                writer.AddItem(writer.AddFolder("Inbox"), document);
                Assert.Throws<InvalidOperationException>(() => writer.Complete());
                Assert.False(File.Exists(blocked));
            }
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void PartialEmlxDeclaredMissingAttachmentsRemainIndeterminate() {
        var document = new EmailDocument();
        document.Properties["Emlx:IsPartial"] = true;
        document.Properties["Emlx:Flag:AttachmentCount"] = 2;
        var item = new EmailStoreItem("partial", "folder", document, format: EmailStoreFormat.Emlx);
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.AttachmentMetadata));
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        Assert.False(item.ContentAvailability.AvailableParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(document, EmailFileFormat.Eml);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_EMLX_PARTIAL_CONTENT" && diagnostic.DataLossRisk == EmailDataLossRisk.Possible);
    }
}
