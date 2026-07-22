namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OfflineAddressBookSessionTests {
    [Fact]
    public void OpensStreamFromCurrentPositionAndRestoresIt() {
        byte[] oab = new OabV4Fixture().Build();
        using (var stream = new MemoryStream()) {
            stream.Write(new byte[17], 0, 17);
            stream.Write(oab, 0, oab.Length);
            stream.Position = 17;

            using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
                Assert.Equal(17, stream.Position);
                Assert.Equal(3, session.DeclaredEntryCount);
                Assert.Single(session.AddressLists);
                Assert.Equal("Synthetic Global Address List", session.AddressLists[0].Name);

                OfflineAddressBookEntry entry = session.EnumerateEntries().First();
                Assert.Equal(17, stream.Position);
                Assert.Equal("Ada Lovelace", entry.DisplayName);
            }
            Assert.Equal(17, stream.Position);
        }
    }

    [Fact]
    public void ProjectsPeopleAndDistributionListsWithoutLosingRawProperties() {
        using (var stream = new MemoryStream(new OabV4Fixture().Build(), writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookEntry[] entries = session.EnumerateEntries().ToArray();
            OfflineAddressBookEntry ada = entries[0];
            OfflineAddressBookEntry list = entries[2];

            Assert.Equal(OfflineAddressBookObjectType.MailUser, ada.ObjectType);
            Assert.Equal("ada@example.test", ada.SmtpAddress);
            Assert.Equal(new[] { "SMTP:ada@example.test", "smtp:alias-ada@example.test" }, ada.ProxyAddresses);
            Assert.True(ada.CanReceiveRichContent);
            Assert.True(ada.IsPropertyTruncated(0x8009101FU));
            Assert.All(ada.Properties, property => Assert.NotNull(property.RawData));
            Assert.Equal("ada@example.test", ada.ToEmailAddress().Address);
            Assert.Equal("Ada", ada.ToOutlookContact().GivenName);

            Assert.True(list.IsDistributionList);
            Assert.Equal(2, list.MemberDistinguishedNames.Count);
            Assert.Equal("all@example.test", list.ToEmailAddress().Address);
        }
    }

    [Fact]
    public void EnumeratesReferencesAndReadsRecordsOnDemand() {
        using (var stream = new MemoryStream(new OabV4Fixture().Build(), writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookEntryReference[] references = session.EnumerateEntryReferences().ToArray();
            Assert.Equal(3, references.Length);
            Assert.True(references[1].RecordOffset > references[0].RecordOffset);
            Assert.Equal("Grace Hopper", session.ReadEntry(references[1]).DisplayName);
        }
    }

    [Fact]
    public void RejectsReferencesCreatedByAnotherSessionSnapshot() {
        byte[] oab = new OabV4Fixture().Build();
        using (var firstStream = new MemoryStream(oab, writable: false))
        using (var secondStream = new MemoryStream(oab, writable: false))
        using (OfflineAddressBookSession first = OfflineAddressBookSession.Open(firstStream, "synthetic.oab"))
        using (OfflineAddressBookSession second = OfflineAddressBookSession.Open(secondStream, "synthetic.oab")) {
            OfflineAddressBookEntryReference reference = first.EnumerateEntryReferences().First();

            Assert.Throws<ArgumentException>(() => second.ReadEntry(reference));
        }
    }

    [Fact]
    public void EnforcesDeclaredEntryAndRecordLimits() {
        byte[] oab = new OabV4Fixture().Build();
        using (var stream = new MemoryStream(oab, writable: false)) {
            var options = new OfflineAddressBookReaderOptions(maxDeclaredEntries: 2);
            OfflineAddressBookLimitExceededException exception = Assert.Throws<OfflineAddressBookLimitExceededException>(
                () => OfflineAddressBookSession.Open(stream, "synthetic.oab", options));
            Assert.Equal(nameof(OfflineAddressBookReaderOptions.MaxDeclaredEntries), exception.LimitName);
        }

        using (var stream = new MemoryStream(oab, writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab",
            new OfflineAddressBookReaderOptions(maxBinaryBytes: 2))) {
            Assert.Throws<OfflineAddressBookLimitExceededException>(() => session.EnumerateEntries(
                new OfflineAddressBookEnumerationOptions(continueOnEntryError: false)).ToArray());
        }
    }

    [Fact]
    public void RejectsUnsupportedAndTruncatedComponentsExplicitly() {
        byte[] unsupported = new byte[16];
        unsupported[0] = 7;
        using (var stream = new MemoryStream(unsupported, writable: false)) {
            Assert.Throws<NotSupportedException>(() => OfflineAddressBookSession.Open(stream, "template.oab"));
        }

        byte[] truncated = new OabV4Fixture().Build().Take(20).ToArray();
        using (var stream = new MemoryStream(truncated, writable: false)) {
            Assert.Throws<InvalidDataException>(() => OfflineAddressBookSession.Open(stream, "truncated.oab"));
        }
    }

    [Fact]
    public void OpensAndEnumeratesLogicallyHugeStreamsWithoutReadingThemEagerly() {
        byte[] oab = new OabV4Fixture().Build();
        const long reportedLength = 32L * 1024 * 1024 * 1024;
        using (var stream = new ReportedLengthStream(oab, reportedLength))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "huge.oab")) {
            Assert.Equal(reportedLength, session.Files.Single().Length);
            Assert.Equal("Ada Lovelace", session.EnumerateEntries(
                new OfflineAddressBookEnumerationOptions(maxEntries: 1)).Single().DisplayName);
            Assert.True(stream.BytesRead < oab.Length * 2L);
            Assert.True(stream.MaximumReadRequest < 1024 * 1024);
        }

        using (var stream = new ReportedLengthStream(oab, reportedLength)) {
            var options = new OfflineAddressBookReaderOptions(maxInputBytes: reportedLength - 1);
            OfflineAddressBookLimitExceededException exception =
                Assert.Throws<OfflineAddressBookLimitExceededException>(() =>
                    OfflineAddressBookSession.Open(stream, "huge.oab", options));
            Assert.Equal(nameof(OfflineAddressBookReaderOptions.MaxInputBytes), exception.LimitName);
        }
    }

    internal sealed class ReportedLengthStream : Stream {
        private readonly byte[] _data;
        private readonly long _length;
        private long _position;

        internal ReportedLengthStream(byte[] data, long length) {
            _data = data;
            _length = length;
        }

        internal long BytesRead { get; private set; }
        internal int MaximumReadRequest { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position {
            get => _position;
            set => _position = value >= 0 && value <= _length
                ? value
                : throw new ArgumentOutOfRangeException(nameof(value));
        }

        public override int Read(byte[] buffer, int offset, int count) {
            MaximumReadRequest = Math.Max(MaximumReadRequest, count);
            if (_position >= _data.Length) return 0;
            int available = checked((int)Math.Min(count, _data.Length - _position));
            Buffer.BlockCopy(_data, checked((int)_position), buffer, offset, available);
            _position += available;
            BytesRead += available;
            return available;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            long target = origin == SeekOrigin.Begin ? offset :
                origin == SeekOrigin.Current ? checked(_position + offset) : checked(_length + offset);
            Position = target;
            return target;
        }

        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
