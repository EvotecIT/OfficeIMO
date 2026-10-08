using OfficeIMO.Access;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessFoundationTests {
        private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        private static AccessTable AddTable(AccessDocument document) {
            AccessTable table = document.Tables.Add("Contacts"); table.Columns.Add("Id", AccessDataType.AutoNumber); table.Columns.Add("Name", AccessDataType.ShortText, 120); return table;
        }

        [Theory]
        [InlineData("jet4.mdb", AccessFormatProfile.Jet4, AccessFileFormat.Mdb)]
        [InlineData("ace12.accdb", AccessFormatProfile.Ace12, AccessFileFormat.Accdb)]
        [InlineData("Application/objects-jet4.mdb", AccessFormatProfile.Jet4, AccessFileFormat.Mdb)]
        [InlineData("Application/objects-ace12.accdb", AccessFormatProfile.Ace12, AccessFileFormat.Accdb)]
        [InlineData("Profiles/password-jet4.mdb", AccessFormatProfile.Jet4, AccessFileFormat.Mdb)]
        [InlineData("Profiles/password-ace.accdb", AccessFormatProfile.Ace14, AccessFileFormat.Accdb)]
        [InlineData("Profiles/large-number.accdb", AccessFormatProfile.Ace16, AccessFileFormat.Accdb)]
        [InlineData("Profiles/extended-date.accdb", AccessFormatProfile.Ace17, AccessFileFormat.Accdb)]
        [InlineData("Native/catalog-scaffold.mdb", AccessFormatProfile.Jet4, AccessFileFormat.Mdb)]
        [InlineData("Native/catalog-scaffold.accdb", AccessFormatProfile.Ace12, AccessFileFormat.Accdb)]
        public async Task IndependentNativeFixturesHaveBoundedInertHeaderEvidence(string name, AccessFormatProfile profile, AccessFileFormat format) {
            byte[] bytes = File.ReadAllBytes(Fixture(name));
            using MemoryStream stream = new MemoryStream(bytes); stream.Position = 17;
            using AccessDocument document = await AccessDocument.LoadAsync(stream, new AccessLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, DecodeCatalog = false });
            Assert.Equal(17, stream.Position); Assert.True(stream.CanRead);
            Assert.Equal(profile, document.Profile); Assert.Equal(format, document.Format);
            Assert.Equal(AccessCatalogStatus.NotDecoded, document.CatalogStatus);
            Assert.Equal(AccessCatalogStatus.NotDecoded, document.VbaProject.CatalogStatus);
            Assert.False(document.Capabilities.Single(x => x.Operation == "model.edit").IsSupported);
            Assert.Contains(document.Inspection!.Diagnostics, d => d.Code == "access.protection.not-assessed");
            using SHA256 hash = SHA256.Create();
            Assert.Equal(BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(), document.Inspection.Sha256);
            Assert.Throws<NotSupportedException>(() => document.Tables["Contacts"]);
            Assert.Throws<InvalidOperationException>(() => document.Tables.Add("MustNotChange"));
            Assert.Throws<InvalidOperationException>(() => document.Save(new MemoryStream()));
            document.Dispose(); Assert.True(stream.CanRead);
        }

        [Fact]
        public void RollbackRetainsExistingIdentityAndDetachesNewObjects() {
            using AccessDocument document = AccessDocument.Create(); AccessTable table = AddTable(document);
            Guid id = table.Id; long revision = document.Revision;
            AccessTable removed;
            using (document.BeginUpdate()) { table.AppendRow(new AccessRowValues { ["Name"] = "Ada" }); removed = document.Tables.Add("Transient"); }
            Assert.Same(table, document.Tables["contacts"]); Assert.Equal(id, table.Id); Assert.Equal(revision, document.Revision); Assert.Equal(0, table.RowCount);
            Assert.Throws<InvalidOperationException>(() => removed.Columns.Add("Invalid", AccessDataType.Int32));
            using (AccessUpdateScope update = document.BeginUpdate()) { table.AppendRow(new AccessRowValues { ["Name"] = "Grace" }); update.Commit(); }
            Assert.Equal(1, table.RowCount); Assert.True(document.Revision > revision);
        }

        [Fact]
        public void ReaderLeasePreservesOmissionNullAndBinaryOwnership() {
            using AccessDocument document = AccessDocument.Create(); AccessTable table = AddTable(document); table.Columns.Add("Payload", AccessDataType.Binary);
            byte[] bytes = { 1, 2 }; table.AppendRow(new AccessRowValues { ["Name"] = null, ["Payload"] = bytes }); bytes[0] = 9;
            using (AccessDataReader reader = table.OpenDataReader()) {
                Assert.True(reader.Read()); Assert.False(reader.IsSpecified(0)); Assert.True(reader.IsSpecified(1)); Assert.True(reader.IsDBNull(0)); Assert.True(reader.IsDBNull(1));
                byte[] returned = (byte[])reader.GetValue(2); Assert.Equal(1, returned[0]); returned[0] = 8; Assert.Equal(1, ((byte[])reader.GetValue(2))[0]);
                Assert.Throws<InvalidOperationException>(() => table.AppendRow(new AccessRowValues()));
                Assert.Throws<InvalidOperationException>(() => document.BeginUpdate());
                Assert.False(reader.Read()); Assert.Throws<InvalidOperationException>(() => reader.GetValue(0));
            }
            table.AppendRow(new AccessRowValues()); Assert.Equal(2, table.RowCount);
        }

        [Fact]
        public void AssessmentIsBoundToDocumentAndCommittedRevision() {
            using AccessDocument document = AccessDocument.Create(); AccessOperationReport report = document.AssessSave("out.accdb");
            Assert.Equal(AccessOperationStatus.Supported, report.Status); report.RequireCurrent(document);
            using AccessDocument other = AccessDocument.Create(); Assert.Throws<InvalidOperationException>(() => report.RequireCurrent(other));
            AddTable(document); Assert.Throws<InvalidOperationException>(() => report.RequireCurrent(document));
            using (document.BeginUpdate()) Assert.Throws<InvalidOperationException>(() => document.AssessSave());
            document.AssessSave().RequireNoLoss();
        }

        [Fact]
        public async Task UnsupportedWritesNeverTouchPathOrStreamEvenWhenLossIsAllowed() {
            using AccessDocument document = AccessDocument.Create(); AddTable(document).AppendRow(new AccessRowValues { ["Name"] = "Ada" });
            document.Queries.Add("Names", "SELECT Name FROM Contacts;");
            string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Access-" + Guid.NewGuid().ToString("N"));
            string path = Path.Combine(root, "out.mdb");
            AccessSaveOptions options = new AccessSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow, FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace };
            AccessOperationNotSupportedException failure = Assert.Throws<AccessOperationNotSupportedException>(() => document.Save(path, options));
            Assert.Contains(failure.Report.Diagnostics, d => d.Code == "access.conversion.unsupported"); Assert.False(Directory.Exists(root));
            using MemoryStream stream = new MemoryStream(); stream.WriteByte(42); stream.Position = 0;
            Assert.Throws<AccessOperationNotSupportedException>(() => document.Save(stream, options));
            Assert.Equal(new byte[] { 42 }, stream.ToArray()); Assert.Equal(0, stream.Position);
            await Assert.ThrowsAsync<AccessOperationNotSupportedException>(() => document.SaveAsync(path, options));
            Assert.Throws<ArgumentException>(() => document.AssessSave(path, new AccessSaveOptions { Format = AccessFileFormat.Accdb }));
            Assert.False(Directory.Exists(root));
        }

        [Fact]
        public async Task CancellationAndReaderDisposalReleaseOwnership() {
            byte[] bytes = File.ReadAllBytes(Fixture("jet4.mdb")); using MemoryStream stream = new MemoryStream(bytes); stream.Position = 8;
            using CancellationTokenSource cancellation = new CancellationTokenSource(); cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => AccessDocument.LoadAsync(stream, cancellationToken: cancellation.Token));
            Assert.Equal(8, stream.Position); Assert.True(stream.CanRead);
            using AccessDocument document = AccessDocument.Create(); AccessTable table = AddTable(document); table.AppendRow(new AccessRowValues());
            using CancellationTokenSource activeCancellation = new CancellationTokenSource(); AccessDataReader reader = table.OpenDataReader(activeCancellation.Token); activeCancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => reader.Read()); reader.Dispose();
            table.AppendRow(new AccessRowValues()); Assert.Equal(2, table.RowCount);
            AccessDataReader disposedReader = table.OpenDataReader(); document.Dispose(); Assert.Throws<ObjectDisposedException>(() => disposedReader.Read()); disposedReader.Dispose();
        }

        [Fact]
        public void InputBudgetsAndMalformedHeadersFailWithoutChangingCallerStream() {
            byte[] fixture = File.ReadAllBytes(Fixture("ace12.accdb"));
            using MemoryStream stream = new MemoryStream(fixture); stream.Position = 7;
            Assert.Throws<InvalidDataException>(() => AccessDocument.Inspect(stream, new AccessLoadOptions { MaxInputBytes = fixture.Length - 1 })); Assert.Equal(7, stream.Position);
            Assert.Throws<InvalidDataException>(() => AccessDocument.Inspect(stream, new AccessLoadOptions { MaxPages = 1 })); Assert.Equal(7, stream.Position);
            foreach (int length in new[] { 0, 20, 21, 4095, fixture.Length - 1 }) {
                using MemoryStream truncated = new MemoryStream(fixture.Take(length).ToArray()); Assert.Throws<InvalidDataException>(() => AccessDocument.Inspect(truncated)); Assert.True(truncated.CanRead);
            }
            fixture[20] = 99; using MemoryStream unknown = new MemoryStream(fixture); Assert.Throws<NotSupportedException>(() => AccessDocument.Inspect(unknown));
            fixture[20] = 2; fixture[4] = 0; using MemoryStream invalid = new MemoryStream(fixture); Assert.Throws<InvalidDataException>(() => AccessDocument.Inspect(invalid));
        }

        [Fact]
        public void NonSeekableInputIsBoundedAndCallerOwned() {
            using ForwardStream stream = new ForwardStream(File.ReadAllBytes(Fixture("jet4.mdb")));
            using AccessDocument document = AccessDocument.Load(stream); Assert.Equal(AccessFormatProfile.Jet4, document.Profile); Assert.True(stream.CanRead);
            using ForwardStream limited = new ForwardStream(File.ReadAllBytes(Fixture("jet4.mdb")));
            Assert.Throws<InvalidDataException>(() => AccessDocument.Load(limited, new AccessLoadOptions { MaxInputBytes = 100 })); Assert.True(limited.CanRead);
        }

        [Fact]
        public void LoadedPathSourceIdentityDetectsExternalChangeAndReleasesTheFile() {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Access-" + Guid.NewGuid().ToString("N") + ".mdb");
            try {
                File.Copy(Fixture("jet4.mdb"), path);
                using AccessDocument document = AccessDocument.Load(path); document.ValidateSourceIdentity();
                byte[] changed = File.ReadAllBytes(path); changed[4096 + 50] ^= 1; File.WriteAllBytes(path, changed);
                Assert.Throws<IOException>(() => document.ValidateSourceIdentity());
                File.Delete(path); Assert.Throws<FileNotFoundException>(() => document.ValidateSourceIdentity());
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public void ModelValidationAndRelationshipsRetainTypedReferences() {
            using AccessDocument document = AccessDocument.Create(); AccessTable groups = document.Tables.Add("Groups"); AccessColumn parent = groups.Columns.Add("Id", AccessDataType.Int32); groups.Indexes.AddPrimaryKey("PK_Groups", "Id");
            AccessTable contacts = AddTable(document); AccessColumn child = contacts.Columns.Add("GroupId", AccessDataType.Int32);
            AccessRelationship relation = document.Relationships.Add("FK_Contacts_Groups", parent, child); Assert.Same(parent, relation.Parent); Assert.Same(child, relation.Child);
            Assert.Throws<ArgumentException>(() => contacts.AppendRow(new AccessRowValues { ["Id"] = "wrong" })); Assert.Equal(0, contacts.RowCount);
            Assert.Throws<ArgumentException>(() => document.Tables.Add("CONTACTS"));
            using AccessDocument other = AccessDocument.Create(); AccessColumn foreign = AddTable(other).Columns[0]; Assert.Throws<ArgumentException>(() => document.Relationships.Add("Invalid", foreign, child));
            Assert.Throws<NotSupportedException>(() => AccessDocument.Create(new AccessCreateOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose }));
            Assert.Throws<NotSupportedException>(() => AccessDocument.Create(new AccessCreateOptions { Format = AccessFileFormat.Mdb, Profile = AccessFormatProfile.Ace12 }));
            Assert.Throws<ArgumentException>(() => document.AssessSave("out.accdb", new AccessSaveOptions { Profile = AccessFormatProfile.Jet4 }));
        }

        [Theory]
        [InlineData(null, false)]
        [InlineData(null, true)]
        [InlineData(AccessFormatProfile.Jet4, false)]
        [InlineData(AccessFormatProfile.Jet4, true)]
        [InlineData(AccessFormatProfile.Ace12, false)]
        [InlineData(AccessFormatProfile.Ace12, true)]
        public void BinaryAccessRejectsNullAndNonBinaryFieldsConsistently(AccessFormatProfile? profile, bool omit) {
            using AccessDocument modeled = AccessDocument.Create(new AccessCreateOptions {
                Format = profile == AccessFormatProfile.Jet4 ? AccessFileFormat.Mdb : AccessFileFormat.Accdb
            });
            AccessTable table = modeled.Tables.Add("Values");
            table.Columns.Add("Payload", AccessDataType.Binary); table.Columns.Add("Label", AccessDataType.ShortText, 10);
            AccessRowValues values = new AccessRowValues { ["Label"] = "text" };
            if (!omit) values["Payload"] = null;
            table.AppendRow(values);
            using MemoryStream output = new MemoryStream();
            if (profile.HasValue) modeled.Save(output);
            using AccessDocument? native = profile.HasValue ? AccessDocument.Load(output) : null;
            using AccessDataReader reader = (native ?? modeled).Tables["Values"].OpenDataReader();
            Assert.True(reader.Read()); Assert.True(reader.IsDBNull(0));
            Assert.Contains("null", Assert.Throws<InvalidCastException>(() => reader.GetStream(0)).Message);
            Assert.Throws<InvalidCastException>(() => reader.GetBytes(0, 0, null, 0, 0));
            Assert.Contains("not binary", Assert.Throws<InvalidCastException>(() => reader.GetStream(1)).Message);
            Assert.Throws<InvalidCastException>(() => reader.GetBytes(1, 0, null, 0, 0));
        }

        private sealed class ForwardStream : Stream {
            private readonly MemoryStream _source;
            internal ForwardStream(byte[] bytes) { _source = new MemoryStream(bytes); }
            public override bool CanRead => _source.CanRead; public override bool CanSeek => false; public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException(); public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override int Read(byte[] buffer, int offset, int count) => _source.Read(buffer, offset, count);
            public override void Flush() { } public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException(); public override void SetLength(long value) => throw new NotSupportedException(); public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            protected override void Dispose(bool disposing) { if (disposing) _source.Dispose(); base.Dispose(disposing); }
        }
    }
}
