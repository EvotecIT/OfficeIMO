using OfficeIMO.Email.Store.Tests.Olm;
using System.Globalization;

namespace OfficeIMO.Email.Store.Tests;

public sealed class StoreMigrationResumeMatrixTests {
    [Theory]
    [InlineData(EmailStoreFormat.Pst)]
    [InlineData(EmailStoreFormat.Ost)]
    [InlineData(EmailStoreFormat.Olm)]
    [InlineData(EmailStoreFormat.Emlx)]
    [InlineData(EmailStoreFormat.Mbox)]
    [InlineData(EmailStoreFormat.MailboxDirectory)]
    public void Every_supported_store_input_resumes_through_one_verified_migration_contract(
        EmailStoreFormat format) {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-migration-matrix-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string source = CreateSource(directory, format);
        string destination = Path.Combine(directory, "destination.pst");
        string checkpoint = Path.Combine(directory, "migration.checkpoint");
        try {
            using (var cancellation = new CancellationTokenSource()) {
                var progress = new InlineProgress<EmailStorePstMigrationProgress>(value => {
                    if (value.InspectedItems == 1) cancellation.Cancel();
                });
                Assert.ThrowsAny<OperationCanceledException>(() =>
                    EmailStoreConverter.ConvertToPst(source, destination,
                        conversionOptions: new EmailStorePstConversionOptions(
                            checkpointPath: checkpoint,
                            checkpointIntervalItems: 1,
                            progress: progress),
                        cancellationToken: cancellation.Token));
            }

            Assert.True(File.Exists(checkpoint));
            Assert.True(File.Exists(checkpoint + ".verify-map"));
            EmailStorePstConversionReport report = EmailStoreConverter.ConvertToPst(
                source, destination,
                conversionOptions: new EmailStorePstConversionOptions(
                    checkpointPath: checkpoint,
                    checkpointIntervalItems: 1));

            Assert.True(report.WasResumed);
            Assert.Equal(format, report.SourceFormat);
            Assert.Equal(format, report.SourceIdentity.Format);
            Assert.True(report.ConvertedItems >= 1);
            Assert.Equal(0, report.SkippedItems);
            Assert.Equal(report.HasDataLoss
                    ? EmailStoreMigrationDisposition.CompletedWithAcceptedLoss
                    : EmailStoreMigrationDisposition.Completed,
                report.Disposition);
            EmailStorePstVerificationReport verification = Assert.IsType<EmailStorePstVerificationReport>(
                report.Verification);
            Assert.Equal(report.ConvertedItems, verification.AttemptedItems);
            Assert.Equal(verification.AttemptedItems,
                verification.MatchedItems + verification.MismatchedItems + verification.FailedItems);
            if (!verification.IsSuccessful) Assert.NotEmpty(verification.Issues);
            Assert.False(File.Exists(checkpoint));
            Assert.False(File.Exists(checkpoint + ".verify-map"));
            using EmailStoreSession result = EmailStoreSession.Open(destination);
            Assert.Equal(report.ConvertedItems, result.EnumerateItems().Count());
        } finally {
            if (File.Exists(checkpoint)) {
                EmailStoreConverter.DeletePstConversionCheckpoint(checkpoint);
            }
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static string CreateSource(string directory, EmailStoreFormat format) {
        switch (format) {
            case EmailStoreFormat.Pst:
                string pst = Path.Combine(directory, "source.pst");
                using (EmailStorePstWriter writer = EmailStorePstWriter.Create(pst)) {
                    string folder = writer.AddFolder("Inbox");
                    writer.AddItem(folder, CreateDocument("pst-one"));
                    writer.AddItem(folder, CreateDocument("pst-two"));
                    writer.Complete();
                }
                return pst;
            case EmailStoreFormat.Ost:
                string ost = Path.Combine(directory, "source.ost");
                File.WriteAllBytes(ost, PstTestFileBuilder.Create(ost: true,
                    attachmentContent: Encoding.UTF8.GetBytes("ost attachment")));
                return ost;
            case EmailStoreFormat.Olm:
                string olm = Path.Combine(directory, "source.olm");
                const string xml = "<emails>" +
                    "<email><OPFMessageCopySubject>olm-one</OPFMessageCopySubject><OPFMessageCopyBody>body-one</OPFMessageCopyBody></email>" +
                    "<email><OPFMessageCopySubject>olm-two</OPFMessageCopySubject><OPFMessageCopyBody>body-two</OPFMessageCopyBody></email>" +
                    "</emails>";
                using (var builder = new OlmTestArchiveBuilder()) {
                    File.WriteAllBytes(olm, builder.AddText(
                        "Local/com.microsoft.__Messages/Inbox/message_00000.xml", xml).Build());
                }
                return olm;
            case EmailStoreFormat.Emlx:
                string emlx = Path.Combine(directory, "source.emlx");
                byte[] message = Encoding.ASCII.GetBytes(
                    "From: source@example.test\r\nSubject: emlx-one\r\n\r\nbody-emlx\r\n");
                byte[] prefix = Encoding.ASCII.GetBytes(message.Length.ToString(
                    CultureInfo.InvariantCulture) + "\n");
                File.WriteAllBytes(emlx, prefix.Concat(message).ToArray());
                return emlx;
            case EmailStoreFormat.Mbox:
                string mbox = Path.Combine(directory, "source.mbox");
                File.WriteAllText(mbox,
                    "From sender@example.test Tue Jul 14 08:30:00 2026\nSubject: mbox-one\n\nbody-one\n" +
                    "From sender@example.test Tue Jul 14 08:31:00 2026\nSubject: mbox-two\n\nbody-two\n",
                    Encoding.ASCII);
                return mbox;
            case EmailStoreFormat.MailboxDirectory:
                string mailbox = Path.Combine(directory, "mailbox");
                Directory.CreateDirectory(Path.Combine(mailbox, "Inbox"));
                File.WriteAllText(Path.Combine(mailbox, "Inbox", "one.eml"),
                    "From: source@example.test\r\nSubject: directory-one\r\n\r\nbody-one\r\n");
                File.WriteAllText(Path.Combine(mailbox, "Inbox", "two.eml"),
                    "From: source@example.test\r\nSubject: directory-two\r\n\r\nbody-two\r\n");
                return mailbox;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }

    private static EmailDocument CreateDocument(string subject) {
        var document = new EmailDocument { Subject = subject, MessageClass = "IPM.Note" };
        document.Body.Text = "body-" + subject;
        return document;
    }

    private sealed class InlineProgress<T> : IProgress<T> {
        private readonly Action<T> _action;
        internal InlineProgress(Action<T> action) { _action = action; }
        public void Report(T value) { _action(value); }
    }
}
