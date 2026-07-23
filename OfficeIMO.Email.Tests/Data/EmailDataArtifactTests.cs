using OfficeIMO.Email.AddressBook.Tests;
using OfficeIMO.Email.Data;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Data.Tests;

public sealed class EmailDataArtifactTests {
    [Fact]
    public void DirectoryWithUnsupportedOabComponentFallsBackToMailboxStore() {
        string directory = TemporaryDirectory();
        try {
            File.WriteAllBytes(Path.Combine(directory, "unsupported.oab"), new byte[] { 7, 0, 0, 0 });
            File.WriteAllText(Path.Combine(directory, "message.eml"),
                "From: sender@example.test\r\nSubject: mailbox fallback\r\n\r\nBody\r\n");

            using EmailDataOpenResult result = EmailDataArtifact.Open(directory);

            Assert.Equal(EmailDataArtifactKind.Store, result.Kind);
            Assert.Single(result.Store!.EnumerateItems());
        } finally {
            TryDeleteDirectory(directory);
        }
    }

    [Fact]
    public void Opens_individual_email_calendar_and_contact_through_existing_owners() {
        string directory = TemporaryDirectory();
        try {
            string emailPath = Path.Combine(directory, "message.eml");
            string calendarPath = Path.Combine(directory, "event.data");
            string contactPath = Path.Combine(directory, "contact.data");
            File.WriteAllText(emailPath,
                "From: sender@example.test\r\nTo: recipient@example.test\r\nSubject: Facade\r\n\r\nBody",
                new UTF8Encoding(false));
            File.WriteAllText(calendarPath,
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//OfficeIMO//Test//EN\r\nEND:VCALENDAR\r\n",
                new UTF8Encoding(false));
            File.WriteAllText(contactPath,
                "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Ada Lovelace\r\nEND:VCARD\r\n",
                new UTF8Encoding(false));

            using (EmailDataOpenResult email = EmailDataArtifact.Open(emailPath)) {
                Assert.Equal(EmailDataArtifactKind.EmailDocument, email.Kind);
                Assert.Equal("Facade", email.EmailDocument!.Subject);
                Assert.Same(email.EmailDocument, email.Artifact);
            }
            using (EmailDataOpenResult calendar = EmailDataArtifact.Open(calendarPath)) {
                Assert.Equal(EmailDataArtifactKind.Calendar, calendar.Kind);
                Assert.Single(calendar.Calendar!.Calendars);
            }
            using (EmailDataOpenResult contact = EmailDataArtifact.Open(contactPath)) {
                Assert.Equal(EmailDataArtifactKind.Contact, contact.Kind);
                Assert.Single(contact.Contact!.Cards);
            }
        } finally {
            TryDeleteDirectory(directory);
        }
    }

    [Fact]
    public void Opens_real_pst_and_oab_sessions_and_owns_their_lifetime() {
        string directory = TemporaryDirectory();
        try {
            string pstPath = Path.Combine(directory, "archive.pst");
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(pstPath)) {
                string inbox = writer.AddFolder("Inbox", EmailStoreSpecialFolderKind.Inbox,
                    containerClass: "IPF.Note");
                writer.AddItem(inbox, new EmailDocument { MessageClass = "IPM.Note", Subject = "Stored" });
                writer.Complete();
            }
            string oabPath = Path.Combine(directory, "udetails.oab");
            File.WriteAllBytes(oabPath, new OabV4Fixture().Build());

            EmailDataOpenResult store = EmailDataArtifact.Open(pstPath);
            Assert.Equal(EmailDataArtifactKind.Store, store.Kind);
            Assert.Equal(EmailStoreFormat.Pst, store.Store!.Format);
            Assert.Single(store.Store.EnumerateItems());
            store.Dispose();
            Assert.Throws<ObjectDisposedException>(() => _ = store.Store.Format);

            using EmailDataOpenResult addressBook = EmailDataArtifact.Open(oabPath);
            Assert.Equal(EmailDataArtifactKind.OfflineAddressBook, addressBook.Kind);
            Assert.Equal(3, addressBook.AddressBook!.DeclaredEntryCount);
        } finally {
            TryDeleteDirectory(directory);
        }
    }

    [Fact]
    public void Explicit_kind_resolves_extension_free_artifacts_without_fallback_parsers() {
        string directory = TemporaryDirectory();
        try {
            string path = Path.Combine(directory, "calendar.bin");
            File.WriteAllText(path,
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//OfficeIMO//Test//EN\r\nEND:VCALENDAR\r\n",
                new UTF8Encoding(false));
            var options = new EmailDataOpenOptions(expectedKind: EmailDataArtifactKind.Calendar);

            using EmailDataOpenResult result = EmailDataArtifact.Open(path, options);

            Assert.Equal(EmailDataArtifactKind.Calendar, result.Kind);
            Assert.Null(result.Email);
            Assert.NotNull(result.Calendar);
        } finally {
            TryDeleteDirectory(directory);
        }
    }

    private static string TemporaryDirectory() {
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-email-data-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(path);
        return path;
    }

    private static void TryDeleteDirectory(string path) {
        if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
    }
}
