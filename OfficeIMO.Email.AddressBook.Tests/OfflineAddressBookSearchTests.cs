namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OfflineAddressBookSearchTests {
    [Fact]
    public void SearchesSemanticFieldsAndReturnsBoundedSummaries() {
        using (var stream = new MemoryStream(new OabV4Fixture().Build(), writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            var query = new OfflineAddressBookSearchQuery(
                new[] { "Grace", "Engineering" },
                fields: OfflineAddressBookSearchFields.Names | OfflineAddressBookSearchFields.Organization);

            OfflineAddressBookSearchReport report = session.Search(query);

            OfflineAddressBookSearchResult result = Assert.Single(report.Results);
            Assert.Equal("Grace Hopper", result.Summary.DisplayName);
            Assert.Equal(
                OfflineAddressBookSearchFields.Names | OfflineAddressBookSearchFields.Organization,
                result.MatchedFields);
            Assert.Contains("Grace", result.Snippet, StringComparison.OrdinalIgnoreCase);
            Assert.True(report.IsComplete);
            Assert.Equal(3, report.EntriesScanned);
        }
    }

    [Fact]
    public void ResumesAtExactRecordOffsetsWithoutDuplicatingMatches() {
        using (var stream = new MemoryStream(new OabV4Fixture().Build(), writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            var names = new List<string>();
            OfflineAddressBookSearchCheckpoint? checkpoint = null;
            do {
                var query = new OfflineAddressBookSearchQuery(
                    new[] { "example" },
                    matchMode: OfflineAddressBookSearchMatchMode.AnyTerm,
                    maxEntriesScanned: 1,
                    maxResults: 1,
                    resumeFrom: checkpoint);
                OfflineAddressBookSearchReport report = session.Search(query);
                names.AddRange(report.Results.Select(result => result.Summary.DisplayName!));
                checkpoint = report.NextCheckpoint;
            } while (checkpoint != null);

            Assert.Equal(new[] { "Ada Lovelace", "Grace Hopper", "All Example" }, names);
        }
    }

    [Fact]
    public void RejectsCheckpointsCreatedByAnotherSessionSnapshot() {
        byte[] oab = new OabV4Fixture().Build();
        using (var firstStream = new MemoryStream(oab, writable: false))
        using (var secondStream = new MemoryStream(oab, writable: false))
        using (OfflineAddressBookSession first = OfflineAddressBookSession.Open(firstStream, "synthetic.oab"))
        using (OfflineAddressBookSession second = OfflineAddressBookSession.Open(secondStream, "synthetic.oab")) {
            OfflineAddressBookSearchReport firstPage = first.Search(new OfflineAddressBookSearchQuery(
                new[] { "example" }, maxEntriesScanned: 1));
            Assert.NotNull(firstPage.NextCheckpoint);
            var resumed = new OfflineAddressBookSearchQuery(
                new[] { "example" }, maxEntriesScanned: 1, resumeFrom: firstPage.NextCheckpoint);

            Assert.Throws<ArgumentException>(() => second.Search(resumed));
        }
    }

    [Fact]
    public void AppliesObjectFilterProgressAndCancellation() {
        using (var stream = new MemoryStream(new OabV4Fixture().Build(), writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            var progress = new CapturingProgress();
            var query = new OfflineAddressBookSearchQuery(
                new[] { "example" },
                objectType: OfflineAddressBookObjectType.DistributionList,
                progressInterval: 1);

            OfflineAddressBookSearchReport report = session.Search(query, progress);

            Assert.Single(report.Results);
            Assert.True(report.Results[0].Summary.IsDistributionList);
            Assert.Equal(3, progress.Reports.Last().EntriesScanned);

            using (var source = new CancellationTokenSource()) {
                source.Cancel();
                Assert.Throws<OperationCanceledException>(() => {
                    session.Search(query, cancellationToken: source.Token);
                });
            }
        }
    }

    private sealed class CapturingProgress : IProgress<OfflineAddressBookSearchProgress> {
        internal List<OfflineAddressBookSearchProgress> Reports { get; } =
            new List<OfflineAddressBookSearchProgress>();

        public void Report(OfflineAddressBookSearchProgress value) => Reports.Add(value);
    }
}
