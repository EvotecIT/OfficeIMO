using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreNativeDirectoryExportTests {
    [Theory]
    [InlineData(EmailStoreNativeDirectoryFormat.Maildir)]
    [InlineData(EmailStoreNativeDirectoryFormat.Emlx)]
    public void NativeDirectoryExportCanBeReadBackThroughMailboxDirectory(
        EmailStoreNativeDirectoryFormat format) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");

            EmailStoreExportReport report = session.ExportToNativeDirectory(root,
                new EmailStoreNativeDirectoryExportOptions(format));
            using EmailStoreSession reopened = EmailStoreSession.Open(root);
            string[] subjects = reopened.EnumerateItems().Select(reference =>
                reopened.ReadSummary(reference).Subject ?? string.Empty).OrderBy(value => value).ToArray();

            Assert.False(report.HasErrors, string.Join(Environment.NewLine,
                report.Diagnostics.Concat(report.Entries.SelectMany(entry => entry.Diagnostics))
                    .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
            Assert.False(report.WasTruncated);
            Assert.Equal(2, report.SucceededCount);
            Assert.Equal(new[] { "Native first", "Native second" }, subjects);
            Assert.All(report.Entries, entry => Assert.True(File.Exists(entry.DestinationPath)));
            Assert.NotNull(report.ManifestPath);
            if (format == EmailStoreNativeDirectoryFormat.Maildir) {
                Assert.Equal("S", report.Entries[0].MaildirFlags);
                string manifest = File.ReadAllText(report.ManifestPath!);
                Assert.Contains("\tMaildirFlags\t", manifest);
                Assert.Contains("\tS\t", manifest);
                Assert.Contains("\tDiagnosticCount", manifest);
                Assert.DoesNotContain("DiagnosticCodes", manifest, StringComparison.Ordinal);
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void MaildirExportKeepsStoredUnreadMessagesInCurWhenSuffixesAreSupported() {
        if (Array.IndexOf(Path.GetInvalidFileNameChars(), ':') >= 0) return;
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-native-maildir-cur-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");

            EmailStoreExportReport report = session.ExportToNativeDirectory(root,
                new EmailStoreNativeDirectoryExportOptions(EmailStoreNativeDirectoryFormat.Maildir));

            EmailStoreExportEntry unread = Assert.Single(report.Entries,
                entry => string.Equals(entry.MaildirFlags, string.Empty, StringComparison.Ordinal));
            Assert.Equal("cur", Path.GetFileName(Path.GetDirectoryName(unread.DestinationPath!)));
            Assert.EndsWith(":2,", Path.GetFileName(unread.DestinationPath!), StringComparison.Ordinal);
            Assert.All(report.Entries, entry => Assert.Equal(
                "cur", Path.GetFileName(Path.GetDirectoryName(entry.DestinationPath!))));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void FlatMaildirExportDoesNotRecreateSourceFolderHierarchy() {
        string sourceRoot = Path.Combine(Path.GetTempPath(),
            "oims-flat-source-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        string destinationRoot = Path.Combine(Path.GetTempPath(),
            "oims-flat-maildir-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        try {
            string inbox = Path.Combine(sourceRoot, "Inbox");
            string archive = Path.Combine(sourceRoot, "Archive");
            Directory.CreateDirectory(inbox);
            Directory.CreateDirectory(archive);
            File.WriteAllText(Path.Combine(inbox, "first.eml"),
                "Subject: First\r\n\r\nBody\r\n");
            File.WriteAllText(Path.Combine(archive, "second.eml"),
                "Subject: Second\r\n\r\nBody\r\n");
            using EmailStoreSession session = EmailStoreSession.Open(sourceRoot);

            EmailStoreExportReport report = session.ExportToNativeDirectory(destinationRoot,
                new EmailStoreNativeDirectoryExportOptions(
                    EmailStoreNativeDirectoryFormat.Maildir,
                    preserveFolderHierarchy: false));

            Assert.False(report.HasErrors);
            Assert.Equal(2, report.SucceededCount);
            Assert.All(report.Entries, entry => {
                string maildirDirectory = Path.GetDirectoryName(entry.DestinationPath!)!;
                string exportRoot = Path.GetDirectoryName(maildirDirectory)!;
                Assert.True(EmailStorePathIdentity.AreEquivalent(destinationRoot, exportRoot));
            });
            Assert.DoesNotContain(Directory.EnumerateDirectories(destinationRoot), path =>
                string.Equals(Path.GetFileName(path), "Inbox", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(Path.GetFileName(path), "Archive", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(sourceRoot)) Directory.Delete(sourceRoot, recursive: true);
            if (Directory.Exists(destinationRoot)) Directory.Delete(destinationRoot, recursive: true);
        }
    }

    [Fact]
    public void EmlxExportDoesNotReplaceAnExistingDestinationByDefault() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");
            var options = new EmailStoreNativeDirectoryExportOptions(
                EmailStoreNativeDirectoryFormat.Emlx, maxItems: 1);
            EmailStoreExportReport first = session.ExportToNativeDirectory(root, options);
            string path = Assert.Single(first.Entries).DestinationPath!;
            byte[] sentinel = Encoding.ASCII.GetBytes("existing destination");
            File.WriteAllBytes(path, sentinel);

            EmailStoreExportReport second = session.ExportToNativeDirectory(root, options);

            Assert.True(second.HasErrors);
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            Assert.DoesNotContain(Directory.EnumerateFiles(Path.GetDirectoryName(path)!),
                candidate => candidate.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void NativeDirectoryExportDoesNotReplaceAnExistingManifestByDefault() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            string manifestPath = Path.Combine(root, "officeimo-email-store-export.tsv");
            byte[] sentinel = Encoding.ASCII.GetBytes("existing manifest");
            File.WriteAllBytes(manifestPath, sentinel);
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");

            EmailStoreExportReport report = session.ExportToNativeDirectory(root,
                new EmailStoreNativeDirectoryExportOptions(
                    EmailStoreNativeDirectoryFormat.Maildir, maxItems: 1));

            Assert.Null(report.ManifestPath);
            Assert.Equal(sentinel, File.ReadAllBytes(manifestPath));
            Assert.Contains(report.Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_MANIFEST_EXISTS");
            Assert.DoesNotContain(Directory.EnumerateFiles(root),
                path => Path.GetFileName(path).StartsWith(".officeimo-email-store-export.",
                    StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void EmlxExportPathsRemainUniqueWhenLongSourceIdsShareTheirTruncatedPrefix() {
        string sourceRoot = Path.Combine(Path.GetTempPath(),
            "officeimo-native-source-" + Guid.NewGuid().ToString("N"));
        string destinationRoot = Path.Combine(Path.GetTempPath(),
            "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(sourceRoot);
            string prefix = new string('a', 72);
            string message = "Subject: Same subject\r\n\r\nBody\r\n";
            File.WriteAllText(Path.Combine(sourceRoot, prefix + "-first.eml"), message);
            File.WriteAllText(Path.Combine(sourceRoot, prefix + "-second.eml"), message);
            using EmailStoreSession session = EmailStoreSession.Open(sourceRoot);

            EmailStoreExportReport report = session.ExportToNativeDirectory(destinationRoot,
                new EmailStoreNativeDirectoryExportOptions(
                    EmailStoreNativeDirectoryFormat.Emlx,
                    preserveFolderHierarchy: false));

            Assert.False(report.HasErrors);
            Assert.Equal(2, report.SucceededCount);
            Assert.Equal(2, report.Entries.Select(entry => entry.DestinationPath)
                .Distinct(StringComparer.OrdinalIgnoreCase).Count());
        } finally {
            if (Directory.Exists(sourceRoot)) Directory.Delete(sourceRoot, recursive: true);
            if (Directory.Exists(destinationRoot)) Directory.Delete(destinationRoot, recursive: true);
        }
    }

    [Fact]
    public void FolderPathsRemainUniqueWhenSanitizedNamesAndTruncatedIdsMatch() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-paths");
        string sharedIdPrefix = "mailbox/" + new string('a', 64);
        var folders = new[] {
            new EmailStoreFolderInfo(sharedIdPrefix + "/first", null, "Projects:2026"),
            new EmailStoreFolderInfo(sharedIdPrefix + "/second", null, "Projects?2026")
        };
        var paths = new EmailStoreExportPathBuilder(root, folders, preserveHierarchy: true);

        string first = paths.GetFolderPath(folders[0].Id);
        string second = paths.GetFolderPath(folders[1].Id);

        Assert.NotEqual(first, second);
        Assert.StartsWith(root, first, StringComparison.Ordinal);
        Assert.StartsWith(root, second, StringComparison.Ordinal);
    }

    [Fact]
    public void EmlxExportContinuesAfterXmlForbiddenMetadataText() {
        string sourceRoot = Path.Combine(Path.GetTempPath(),
            "officeimo-native-source-" + Guid.NewGuid().ToString("N"));
        string destinationRoot = Path.Combine(Path.GetTempPath(),
            "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(sourceRoot);
            File.WriteAllText(Path.Combine(sourceRoot, "bad.eml"),
                "Subject: Bad\u0001subject\r\n\r\nBody\r\n");
            File.WriteAllText(Path.Combine(sourceRoot, "good.eml"),
                "Subject: Good subject\r\n\r\nBody\r\n");
            using EmailStoreSession session = EmailStoreSession.Open(sourceRoot);

            EmailStoreExportReport report = session.ExportToNativeDirectory(destinationRoot,
                new EmailStoreNativeDirectoryExportOptions(
                    EmailStoreNativeDirectoryFormat.Emlx,
                    preserveFolderHierarchy: false,
                    continueOnError: true));

            Assert.Equal(2, report.Entries.Count);
            Assert.Equal(1, report.SucceededCount);
            Assert.Contains(report.Entries, entry => entry.Diagnostics.Any(diagnostic =>
                diagnostic.Code == "EMAIL_STORE_EMLX_EXPORT_FAILED"));
        } finally {
            if (Directory.Exists(sourceRoot)) Directory.Delete(sourceRoot, recursive: true);
            if (Directory.Exists(destinationRoot)) Directory.Delete(destinationRoot, recursive: true);
        }
    }

    [Fact]
    public void MaildirExportContinuesAfterAMessageExceedsTheWriterLimit() {
        string sourceRoot = Path.Combine(Path.GetTempPath(),
            "oims-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        string destinationRoot = Path.Combine(Path.GetTempPath(),
            "oimd-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        try {
            Directory.CreateDirectory(sourceRoot);
            File.WriteAllText(Path.Combine(sourceRoot, "01-oversized.eml"),
                "Subject: Oversized\r\n\r\n" + new string('x', 65_536));
            File.WriteAllText(Path.Combine(sourceRoot, "02-valid.eml"),
                "Subject: Valid\r\n\r\nSmall body\r\n");
            using EmailStoreSession session = EmailStoreSession.Open(sourceRoot);
            var messageOptions = new EmailWriterOptions(maxOutputBytes: 4_096);

            EmailStoreExportReport report = session.ExportToNativeDirectory(destinationRoot,
                new EmailStoreNativeDirectoryExportOptions(
                    EmailStoreNativeDirectoryFormat.Maildir,
                    preserveFolderHierarchy: false,
                    continueOnError: true,
                    messageOptions: messageOptions));

            Assert.Equal(2, report.Entries.Count);
            Assert.True(report.SucceededCount == 1, string.Join(Environment.NewLine,
                report.Entries.SelectMany(entry => entry.Diagnostics)
                    .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
            Assert.Contains(report.Entries, entry => entry.Diagnostics.Any(diagnostic =>
                diagnostic.Code == "EMAIL_STORE_MAILDIR_EXPORT_FAILED" &&
                diagnostic.Message.Contains(nameof(EmailWriterOptions.MaxOutputBytes),
                    StringComparison.Ordinal)));
            EmailStoreExportEntry succeeded = Assert.Single(report.Entries, entry => entry.Succeeded);
            Assert.True(File.Exists(succeeded.DestinationPath));
            Assert.Equal("Valid", EmailDocument.Load(succeeded.DestinationPath!).Subject);
        } finally {
            if (Directory.Exists(sourceRoot)) Directory.Delete(sourceRoot, recursive: true);
            if (Directory.Exists(destinationRoot)) Directory.Delete(destinationRoot, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDirectoryExportRejectsItsMailboxDirectorySourceTree(bool useDescendant) {
        string sourceRoot = Path.Combine(Path.GetTempPath(),
            "oims-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        try {
            Directory.CreateDirectory(sourceRoot);
            string sourceMessage = Path.Combine(sourceRoot, "source.eml");
            File.WriteAllText(sourceMessage, "Subject: Source\r\n\r\nBody\r\n");
            using EmailStoreSession session = EmailStoreSession.Open(sourceRoot);
            string destination = useDescendant ? Path.Combine(sourceRoot, "export") : sourceRoot;

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                session.ExportToNativeDirectory(destination,
                    new EmailStoreNativeDirectoryExportOptions(
                        EmailStoreNativeDirectoryFormat.Maildir)));

            Assert.Contains("source tree", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(new[] { sourceMessage }, Directory.EnumerateFiles(
                sourceRoot, "*", SearchOption.AllDirectories).ToArray());
            if (useDescendant) Assert.False(Directory.Exists(destination));
        } finally {
            if (Directory.Exists(sourceRoot)) Directory.Delete(sourceRoot, recursive: true);
        }
    }

#if NET8_0_OR_GREATER
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void NativeDirectoryExportRejectsLinkedAliasesOfItsSource(int scenario) {
        string container = Path.Combine(Path.GetTempPath(),
            "oims-link-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        string sourceRoot = Path.Combine(container, "source");
        string alias = Path.Combine(container, "alias");
        bool linkCreated = false;
        try {
            Directory.CreateDirectory(sourceRoot);
            string sourceMessage = Path.Combine(sourceRoot, "source.eml");
            File.WriteAllText(sourceMessage, "Subject: Source\r\n\r\nBody\r\n");
            try {
                Directory.CreateSymbolicLink(alias, sourceRoot);
                linkCreated = true;
            } catch (UnauthorizedAccessException) when (
                System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(
                    System.Runtime.InteropServices.OSPlatform.Windows)) {
                return;
            } catch (IOException) when (
                System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(
                    System.Runtime.InteropServices.OSPlatform.Windows)) {
                return;
            }

            string source = scenario == 2 ? alias : sourceRoot;
            string destination = scenario switch {
                0 => alias,
                1 => Path.Combine(alias, "export"),
                _ => sourceRoot
            };
            using EmailStoreSession session = EmailStoreSession.Open(source);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                session.ExportToNativeDirectory(destination,
                    new EmailStoreNativeDirectoryExportOptions(
                        EmailStoreNativeDirectoryFormat.Maildir)));

            Assert.Contains("source tree", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(new[] { sourceMessage }, Directory.EnumerateFiles(
                sourceRoot, "*", SearchOption.AllDirectories).ToArray());
        } finally {
            if (linkCreated && Directory.Exists(alias)) Directory.Delete(alias);
            if (Directory.Exists(container)) Directory.Delete(container, recursive: true);
        }
    }
#endif

    [Theory]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(int.MaxValue)]
    public void NativeDirectoryExportOptionsRejectUndefinedFormats(int value) {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new EmailStoreNativeDirectoryExportOptions(
                (EmailStoreNativeDirectoryFormat)value));
    }

    private static byte[] CreateMailboxBytes() {
        var mailbox = new EmailMailbox();
        var first = new EmailDocument { Subject = "Native first" };
        first.Body.Text = "First body";
        first.MessageMetadata.IsRead = true;
        mailbox.Messages.Add(new EmailMailboxEntry(first) { EnvelopeSender = "first@example.com" });
        var second = new EmailDocument { Subject = "Native second" };
        second.Body.Text = "Second body";
        mailbox.Messages.Add(new EmailMailboxEntry(second) { EnvelopeSender = "second@example.com" });
        return mailbox.ToBytes();
    }
}
