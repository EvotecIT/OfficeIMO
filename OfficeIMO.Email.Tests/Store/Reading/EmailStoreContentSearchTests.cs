using OfficeIMO.Email;
using OfficeIMO.Rtf;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreContentSearchTests {
    [Fact]
    public void SearchesSemanticHtmlAndReportsMatchedFieldAndSnippet() {
        string root = CreateCorpus();
        try {
            using EmailStoreSession session = EmailStoreSession.Open(root);
            var progress = new CaptureProgress();
            EmailStoreItem htmlItem = session.EnumerateItems()
                .Select(reference => session.ReadItem(reference))
                .Single(item => item.Document.Subject == "Second");
            Assert.Contains("visible target", htmlItem.Document.Body.Html!, StringComparison.OrdinalIgnoreCase);

            EmailStoreContentSearchReport report = session.SearchContent(
                new EmailStoreContentQuery(
                    new[] { "visible target", "more" },
                    fields: EmailStoreContentSearchFields.HtmlBody,
                    maxItemsScanned: 10,
                    maxResults: 10,
                    progressInterval: 1),
                progress);

            EmailStoreContentSearchResult result = Assert.Single(report.Results);
            Assert.Equal(EmailStoreContentSearchFields.HtmlBody, result.MatchedFields);
            Assert.Contains("visible target", result.Snippet!, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("& more", result.Snippet!, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<p>", result.Snippet!, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("hidden script target", result.Snippet!, StringComparison.OrdinalIgnoreCase);
            Assert.True(report.IsComplete);
            Assert.Equal(4, report.ItemsScanned);
            Assert.Equal(4, progress.Last!.ItemsScanned);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ResumesWithoutRepeatingProcessedItems() {
        string root = CreateCorpus();
        try {
            using EmailStoreSession session = EmailStoreSession.Open(root);
            var firstQuery = new EmailStoreContentQuery(
                new[] { "common needle" },
                fields: EmailStoreContentSearchFields.TextBody | EmailStoreContentSearchFields.HtmlBody,
                matchMode: EmailStoreContentMatchMode.AnyTerm,
                maxItemsScanned: 1,
                maxResults: 10);

            EmailStoreContentSearchReport first = session.SearchContent(firstQuery);
            EmailStoreContentSearchReport second = session.SearchContent(new EmailStoreContentQuery(
                new[] { "common needle" },
                fields: EmailStoreContentSearchFields.TextBody | EmailStoreContentSearchFields.HtmlBody,
                matchMode: EmailStoreContentMatchMode.AnyTerm,
                maxItemsScanned: 1,
                maxResults: 10,
                resumeFrom: first.NextCheckpoint));

            Assert.Single(first.Results);
            Assert.Single(second.Results);
            Assert.NotNull(first.NextCheckpoint);
            Assert.NotNull(second.NextCheckpoint);
            Assert.Equal(1, first.NextCheckpoint!.ItemOffset);
            Assert.Equal(2, second.NextCheckpoint!.ItemOffset);
            Assert.NotEqual(first.Results[0].Reference.Id, second.Results[0].Reference.Id);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ValidatesContentSearchBounds() {
        Assert.Throws<ArgumentException>(() => new EmailStoreContentQuery(Array.Empty<string>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreContentQuery(
            new[] { "value" }, maxItemsScanned: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreContentQuery(
            new[] { "value" }, fields: EmailStoreContentSearchFields.None));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new EmailStoreContentSearchCheckpoint(-1));
    }

    [Fact]
    public void SearchesSemanticRtfTextInsteadOfControlSource() {
        string root = CreateCorpus();
        try {
            using EmailStoreSession session = EmailStoreSession.Open(root);

            EmailStoreContentSearchReport report = session.SearchContent(
                new EmailStoreContentQuery(
                    new[] { "Unicode snowman ☃" },
                    fields: EmailStoreContentSearchFields.RtfBody,
                    maxItemsScanned: 10,
                    maxResults: 10));

            EmailStoreContentSearchResult result = Assert.Single(report.Results);
            Assert.Equal("Fourth", result.Summary.Subject);
            Assert.Contains("Unicode snowman ☃", result.Snippet!, StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static string CreateCorpus() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-content-search-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "01.eml"),
            "From: first@example.test\r\nTo: recipient@example.test\r\n" +
            "Subject: First\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n" +
            "common needle in the first body\r\n");
        File.WriteAllText(Path.Combine(root, "02.eml"),
            "From: second@example.test\r\nSubject: Second\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n\r\n" +
            "<style>.common-needle { color: red; }</style>" +
            "<script>hidden script target</script><p>common needle and visible target &amp; more</p>\r\n");
        File.WriteAllText(Path.Combine(root, "03.eml"),
            "From: third@example.test\r\nSubject: Third\r\n" +
            "Content-Type: text/plain; charset=utf-8\r\n\r\n" +
            "common needle in the third body\r\n");
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddParagraph("Unicode snowman ☃ searchable RTF body");
        var rtfMessage = new EmailDocument { Subject = "Fourth" };
        rtfMessage.Body.Rtf = rtf.ToRtf();
        File.WriteAllBytes(Path.Combine(root, "04.eml"),
            new EmailDocumentWriter().ToBytes(rtfMessage, EmailFileFormat.Eml));
        return root;
    }

    private sealed class CaptureProgress : IProgress<EmailStoreContentSearchProgress> {
        internal EmailStoreContentSearchProgress? Last { get; private set; }
        public void Report(EmailStoreContentSearchProgress value) => Last = value;
    }
}
