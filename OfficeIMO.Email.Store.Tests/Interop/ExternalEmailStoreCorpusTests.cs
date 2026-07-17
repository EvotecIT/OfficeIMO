using System.Globalization;

namespace OfficeIMO.Email.Store.Tests;

public sealed class ExternalEmailStoreCorpusTests {
    [EnvironmentFact("OFFICEIMO_EMAIL_STORE_CORPUS", requireDirectory: true)]
    public void ReadsBoundedPrivatePstAndOstCorpusWithoutTableTraversalFailures() {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_STORE_CORPUS")!;
        int maximumItems = ReadPositiveEnvironmentInteger(
            "OFFICEIMO_EMAIL_STORE_CORPUS_MAX_ITEMS", defaultValue: 10_000);

        string[] paths = Directory.GetFiles(root, "*.*", SearchOption.TopDirectoryOnly)
            .Where(path => path.EndsWith(".pst", StringComparison.OrdinalIgnoreCase) ||
                           path.EndsWith(".ost", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.NotEmpty(paths);

        var failures = new List<string>();
        long totalItems = 0;
        for (int storeIndex = 0; storeIndex < paths.Length; storeIndex++) {
            string path = paths[storeIndex];
            var before = new FileInfo(path);
            long beforeLength = before.Length;
            DateTime beforeWriteTimeUtc = before.LastWriteTimeUtc;
            using EmailStoreSession session = EmailStoreSession.Open(path,
                new EmailStoreReaderOptions(includeAssociatedItems: true, includeOrphanedItems: true));
            var references = new List<EmailStoreItemReference>();
            foreach (EmailStoreItemReference reference in session.EnumerateItems(new EmailStoreEnumerationOptions(
                maxItems: maximumItems,
                includeAssociatedItems: true,
                includeOrphanedItems: true))) {
                references.Add(reference);
            }
            int itemCount = references.Count;
            totalItems = checked(totalItems + itemCount);

            var safeProjection = new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.Bodies |
                EmailStoreItemReadParts.Recipients | EmailStoreItemReadParts.AttachmentMetadata,
                maxDecodedPropertyBytes: 32L * 1024 * 1024,
                preferStreamingAttachmentContent: true);
            foreach (EmailStoreItemReference reference in references.Take(8)) {
                _ = session.ReadItem(reference, safeProjection);
            }

            EmailStoreDiagnostic[] tableFailures = session.Diagnostics.Where(diagnostic =>
                diagnostic.Code == "EMAIL_STORE_PST_TABLE_CONTEXT").ToArray();
            if (tableFailures.Length > 0) {
                failures.Add(string.Concat(
                    "store-", (storeIndex + 1).ToString(CultureInfo.InvariantCulture),
                    " format=", session.Format,
                    " bytes=", beforeLength.ToString(CultureInfo.InvariantCulture),
                    " items=", itemCount.ToString(CultureInfo.InvariantCulture),
                    " table-errors=", tableFailures.Length.ToString(CultureInfo.InvariantCulture),
                    Environment.NewLine,
                    string.Join(Environment.NewLine, tableFailures
                        .Select(diagnostic => string.Concat(diagnostic.Location, ": ", diagnostic.Message))
                        .Distinct(StringComparer.Ordinal))));
            }

            before.Refresh();
            Assert.Equal(beforeLength, before.Length);
            Assert.Equal(beforeWriteTimeUtc, before.LastWriteTimeUtc);
        }

        Assert.True(totalItems > 0, "The external PST/OST corpus contained no enumerable items.");
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    [EnvironmentFact("OFFICEIMO_EMAIL_STORE_CORPUS_CONVERT", "1")]
    public void ConvertsBoundedPrivateStoresToTemporaryPstWithoutRetainingCorpusContent() {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_STORE_CORPUS");
        Assert.True(!string.IsNullOrWhiteSpace(root) && Directory.Exists(root),
            "OFFICEIMO_EMAIL_STORE_CORPUS must name a private local PST/OST directory.");
        int maximumItems = ReadPositiveEnvironmentInteger(
            "OFFICEIMO_EMAIL_STORE_CORPUS_CONVERT_MAX_ITEMS", defaultValue: 8);
        int maximumStores = ReadPositiveEnvironmentInteger(
            "OFFICEIMO_EMAIL_STORE_CORPUS_MAX_STORES", defaultValue: 1);
        string[] paths = Directory.GetFiles(root!, "*.*", SearchOption.TopDirectoryOnly)
            .Where(path => path.EndsWith(".pst", StringComparison.OrdinalIgnoreCase) ||
                           path.EndsWith(".ost", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Take(maximumStores)
            .ToArray();
        Assert.NotEmpty(paths);

        int convertedItems = 0;
        foreach (string sourcePath in paths) {
            var source = new FileInfo(sourcePath);
            long sourceLength = source.Length;
            DateTime sourceWriteTimeUtc = source.LastWriteTimeUtc;
            string destinationPath = Path.Combine(Path.GetTempPath(),
                string.Concat("officeimo-private-conversion-", Guid.NewGuid().ToString("N"), ".pst"));
            try {
                var readerOptions = new EmailStoreReaderOptions(
                    maxItemCount: maximumItems,
                    maxAttachmentBytes: 64L * 1024 * 1024,
                    maxTotalAttachmentBytes: 256L * 1024 * 1024,
                    retainAttachmentContent: false,
                    includeAssociatedItems: true,
                    includeOrphanedItems: true);
                EmailStorePstConversionReport report = EmailStoreConverter.ConvertToPst(
                    sourcePath, destinationPath, readerOptions,
                    new EmailStorePstConversionOptions(
                        continueOnItemError: true,
                        includeAssociatedItems: true,
                        includeOrphanedItems: true,
                        maxItems: maximumItems));
                convertedItems = checked(convertedItems + report.ConvertedItems);
                Assert.Equal(report.ConvertedItems, report.WriteReport.ItemCount);
                Assert.NotNull(report.Verification);
                Assert.True(report.Verification!.IsSuccessful);
                Assert.Equal(report.ConvertedItems, report.Verification.AttemptedItems);
                Assert.Equal(report.ConvertedItems, report.Verification.MatchedItems);
                Assert.Empty(report.Verification.Issues);
                Assert.Null(report.Verification.ManifestPath);
                using EmailStoreSession converted = EmailStoreSession.Open(destinationPath,
                    new EmailStoreReaderOptions(includeAssociatedItems: true,
                        includeOrphanedItems: true));
                Assert.Equal(report.ConvertedItems, converted.EnumerateItems(
                    new EmailStoreEnumerationOptions(
                        maxItems: maximumItems,
                        includeAssociatedItems: true,
                        includeOrphanedItems: true)).Count());

                source.Refresh();
                Assert.Equal(sourceLength, source.Length);
                Assert.Equal(sourceWriteTimeUtc, source.LastWriteTimeUtc);
            } finally {
                try { if (File.Exists(destinationPath)) File.Delete(destinationPath); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
        Assert.True(convertedItems > 0,
            "The bounded private-corpus conversion did not produce an item.");
    }

    private static int ReadPositiveEnvironmentInteger(string name, int defaultValue) {
        string? value = Environment.GetEnvironmentVariable(name);
        return int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) && parsed > 0
            ? parsed
            : defaultValue;
    }
}
