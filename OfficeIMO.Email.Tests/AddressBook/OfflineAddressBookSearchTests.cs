namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OfflineAddressBookSearchTests {
    [Fact]
    public void QuerySignatureDistinguishesTermsContainingLegacyDelimiter() {
        var singleTerm = new OfflineAddressBookSearchQuery(new[] { "alpha\u001fbeta" });
        var separateTerms = new OfflineAddressBookSearchQuery(new[] { "alpha", "beta" });

        Assert.NotEqual(singleTerm.Signature, separateTerms.Signature);
    }

    [Fact]
    public void QueryTermsCannotBeReplacedAfterTheCheckpointSignatureIsCaptured() {
        var query = new OfflineAddressBookSearchQuery(new[] { "needle" });

        Assert.Throws<NotSupportedException>(() =>
            ((IList<string>)query.Terms)[0] = "different");
        Assert.Equal("needle", Assert.Single(query.Terms));
    }

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
    public void ResumesCheckpointsAcrossSessionsForTheSameSource() {
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

            OfflineAddressBookSearchReport secondPage = second.Search(resumed);
            Assert.NotEmpty(secondPage.Results);
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

    [Fact]
    public void SearchRejectsSameLengthSourceMutationBeforePublishingCheckpoint() {
        using var stream = new MutatingAfterFullReadStream(new OabV4Fixture().Build());
        using OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab");
        stream.Arm();

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            session.Search(new OfflineAddressBookSearchQuery(
                new[] { "example" }, maxEntriesScanned: 1, maxResults: 1)));

        Assert.Contains("source changed", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class CapturingProgress : IProgress<OfflineAddressBookSearchProgress> {
        internal List<OfflineAddressBookSearchProgress> Reports { get; } =
            new List<OfflineAddressBookSearchProgress>();

        public void Report(OfflineAddressBookSearchProgress value) => Reports.Add(value);
    }

    private sealed class MutatingAfterFullReadStream : MemoryStream {
        private bool _armed;
        private bool _mutated;

        internal MutatingAfterFullReadStream(byte[] bytes)
            : base(bytes, 0, bytes.Length, writable: true, publiclyVisible: true) { }

        internal void Arm() => _armed = true;

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (_armed && !_mutated && read > 0 && Position == Length) {
                byte[] content = GetBuffer();
                int index = checked((int)Length - 1);
                content[index] ^= 1;
                _mutated = true;
            }
            return read;
        }
    }
}
