using OfficeIMO.Access;
using System.Collections.Generic;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessNativeReaderTests {
        [Theory]
        [InlineData("jet4.mdb")]
        [InlineData("ace12.accdb")]
        [InlineData("Native/catalog-scaffold.mdb")]
        [InlineData("Native/catalog-scaffold.accdb")]
        public void NativeSchemaRowsIndexesAndRelationshipsAgreeWithDao(string file) {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
            Assert.Equal(AccessCatalogStatus.Decoded, document.CatalogStatus);
            Assert.Equal(AccessCatalogStatus.Decoded, document.Forms.CatalogStatus);
            Assert.Equal(2, document.Tables.Count);
            AccessTable contacts = document.Tables["Contacts"];
            Assert.Equal(2, contacts.RowCount);
            Assert.Equal(!file.StartsWith("Native/"), contacts.Columns["Id"].IsAutoNumber);
            Assert.True(contacts.Indexes["PK_Contacts"].IsPrimaryKey);
            Assert.True(contacts.Indexes["PK_Contacts"].IsUnique);
            string relationshipName = file.StartsWith("Native/") ? "ContactGroups" : "FK_Contacts_Groups";
            Assert.True(contacts.Indexes[relationshipName].IsForeignKey);
            AccessRelationship relationship = document.Relationships[relationshipName];
            Assert.Same(document.Tables["Groups"].Columns["Id"], relationship.Parent);
            Assert.Same(contacts.Columns["GroupId"], relationship.Child);
            using AccessDataReader rows = contacts.OpenDataReader();
            Assert.True(rows.Read()); Assert.Equal(1, rows.GetInt32(0));
            Assert.Equal("Ada", rows.GetString(rows.GetOrdinal("DisplayName")));
            if (!file.StartsWith("Native/")) {
                Assert.Equal(12.3456m, rows.GetDecimal(rows.GetOrdinal("Amount")));
                Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), rows.GetDateTime(rows.GetOrdinal("CreatedAt")));
                Assert.True(rows.GetBoolean(rows.GetOrdinal("Active")));
                Assert.Equal("Synthetic text", rows.GetString(rows.GetOrdinal("Notes")));
            }
            Assert.True(rows.Read()); Assert.Equal(2, rows.GetInt32(0));
            if (!file.StartsWith("Native/")) {
                Assert.Equal("", rows.GetString(rows.GetOrdinal("DisplayName")));
                Assert.Equal(-1.25m, rows.GetDecimal(rows.GetOrdinal("Amount")));
                Assert.True(rows.IsDBNull(rows.GetOrdinal("CreatedAt")));
                Assert.False(rows.GetBoolean(rows.GetOrdinal("Active")));
                Assert.True(rows.IsDBNull(rows.GetOrdinal("Notes")));
            }
            Assert.False(rows.Read());
            Assert.Throws<NotSupportedException>(() => contacts.Columns.Add("Blocked", AccessDataType.Int32));
        }
        [Theory]
        [InlineData("Readers/values-jet4.mdb")]
        [InlineData("Readers/values-ace.accdb")]
        public void ExtendedScalarAndFragmentedRowsAgreeWithDao(string file) {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
            AccessTable table = document.Tables["Scalars"];
            Assert.Equal("Unicode label", table.Columns["Label"].Properties["Caption"]);
            Assert.Equal((short)111, table.Columns["Label"].Properties["DisplayControl"]);
            Assert.Equal(28, table.Columns["Precise"].Precision); Assert.Equal(9, table.Columns["Precise"].Scale);
            Assert.Equal(file.EndsWith("accdb"), table.Columns["Notes"].IsRichText);
            using AccessDataReader rows = table.OpenDataReader(); Assert.True(rows.Read());
            Assert.Equal("Zażółć gęślą jaźń 漢字 العربية 🙂", rows.GetString(rows.GetOrdinal("Label")));
            Assert.Equal((byte)255, rows.GetByte(rows.GetOrdinal("Tiny")));
            Assert.Equal(short.MinValue, rows.GetInt16(rows.GetOrdinal("Small")));
            Assert.Equal(int.MaxValue, rows.GetInt32(rows.GetOrdinal("Whole")));
            Assert.Equal(1.25f, rows.GetFloat(rows.GetOrdinal("Real")));
            Assert.Equal(-2.5d, rows.GetDouble(rows.GetOrdinal("Wide")));
            Assert.Equal(-922337203685477.5808m, rows.GetDecimal(rows.GetOrdinal("Money")));
            Assert.Equal(1234567890123456789.123456789m, rows.GetDecimal(rows.GetOrdinal("Precise")));
            Assert.Equal(new Guid("00112233-4455-6677-8899-aabbccddeeff"), rows.GetGuid(rows.GetOrdinal("Token")));
            Assert.Equal(string.Concat(Enumerable.Repeat("<div>Ł🙂 native long text</div>", 500)), rows.GetString(rows.GetOrdinal("Notes")));
            Assert.Equal(Enumerable.Range(0, 20000).Select(x => (byte)(x * 17 % 251)).ToArray(), rows.GetValue(rows.GetOrdinal("Payload")));
            Assert.True(rows.Read()); Assert.Equal("", rows.GetString(rows.GetOrdinal("Label"))); Assert.Equal(-0.000000001m, rows.GetDecimal(rows.GetOrdinal("Precise"))); Assert.False(rows.Read());
            AccessRelationship relation = document.Relationships["CompositeKeys"]; Assert.Equal(2, relation.Fields.Count); Assert.True(relation.CascadeDeletes); Assert.True(relation.CascadeUpdates);
            Assert.Equal(new[] { true, false }, document.Tables["ChildKeys"].Indexes["DescendingPair"].Descending);
            using AccessDataReader fragmented = document.Tables["Fragmented"].OpenDataReader(); List<int> ids = new List<int>();
            while (fragmented.Read()) { int id = fragmented.GetInt32(0); ids.Add(id); Assert.Equal(id % 2 == 1 ? string.Concat(Enumerable.Repeat("Growing row creates native overflow storage ", 5)) : "short", fragmented.GetString(1)); }
            Assert.Equal(Enumerable.Range(0, 180).Where(x => x % 3 != 0).ToArray(), ids.OrderBy(x => x).ToArray());
        }
        [Fact]
        public void StructuredValuesRetainTypedMetadataAndIndependentAttachmentBytes() {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Readers/values-ace.accdb"));
            AccessTable table = document.Tables["Structured"];
            Assert.Equal(AccessComplexKind.MultiValue, table.Columns["Tags"].ComplexDefinition!.Kind);
            Assert.Equal(AccessComplexKind.Attachment, table.Columns["Files"].ComplexDefinition!.Kind);
            using AccessDataReader rows = table.OpenDataReader(); Assert.True(rows.Read());
            AccessComplexValue tags = Assert.IsType<AccessComplexValue>(rows["Tags"]);
            Assert.Equal(new[] { "Alpha", "Ł🙂" }, tags.EnumerateValues().Cast<string>().ToArray());
            AccessComplexValue files = Assert.IsType<AccessComplexValue>(rows["Files"]);
            AccessAttachment attachment = Assert.Single(files.EnumerateAttachments());
            Assert.Equal("synthetic.bin", attachment.FileName); Assert.Equal("bin", attachment.FileType);
            Assert.Equal(Enumerable.Range(0, 256).Select(x => (byte)x).ToArray(), attachment.GetBytes());
            byte[] encoded = attachment.GetEncodedBytes(); encoded[0] = 99;
            Assert.Equal(1, attachment.GetEncodedBytes()[0]);
            using Stream stream = attachment.OpenRead(); Assert.False(stream.CanWrite); Assert.Equal(256, stream.Length);
            using CancellationTokenSource canceled = new CancellationTokenSource(); canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => attachment.GetBytes(canceled.Token));
            document.Dispose(); Assert.Throws<ObjectDisposedException>(() => attachment.GetBytes());
        }
        [Fact]
        public async Task NativeBudgetsCancellationAndLazyBinaryStreamHaveDeterministicBehavior() {
            string file = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Readers/values-ace.accdb");
            using (AccessDocument limited = AccessDocument.Load(file, new AccessLoadOptions { MaxValueBytes = 128, MaxRows = 1, TableNames = new[] { "Scalars" } })) {
                Assert.Single(limited.Tables); using AccessDataReader rows = limited.Tables[0].OpenDataReader();
                Assert.True(rows.Read()); Assert.Equal(1, rows.GetInt32(0));
                Assert.False(rows.IsDBNull(rows.GetOrdinal("Payload")));
                Assert.Throws<InvalidDataException>(() => rows.GetStream(rows.GetOrdinal("Payload")));
                Assert.Throws<InvalidDataException>(() => rows.Read());
            }
            using AccessDocument source = AccessDocument.Load(file); using AccessDataReader reader = source.Tables["Scalars"].OpenDataReader(); Assert.True(reader.Read());
            int ordinal = reader.GetOrdinal("Payload"); Assert.Equal(20000, reader.GetBytes(ordinal, 0, null, 0, 0));
            byte[] buffer = new byte[257]; Assert.Equal(257, reader.GetBytes(ordinal, 17000, buffer, 0, buffer.Length));
            Assert.Equal(Enumerable.Range(17000, 257).Select(x => (byte)(x * 17 % 251)).ToArray(), buffer);
            using Stream stream = reader.GetStream(ordinal); Assert.False(stream.CanSeek); Assert.False(stream.CanWrite);
            using MemoryStream bytes = new MemoryStream(); await stream.CopyToAsync(bytes);
            Assert.Equal(20000, bytes.Length);
            using CancellationTokenSource cancellation = new CancellationTokenSource(); cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadAsync(cancellation.Token));
            source.Dispose(); Assert.Throws<ObjectDisposedException>(() => stream.Read(buffer, 0, 1));
        }
        [Theory]
        [InlineData("access2000.mdb", "08.50")]
        [InlineData("access2002-2003.mdb", "09.50")]
        public void Jet4ApplicationFormatsKeepPropertiesQueriesSecurityAndLinksInert(string file, string version) {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Generations", file));
            Assert.Equal(version, document.Properties["AccessVersion"]);
            Assert.Contains(document.Catalog, x => x.Owner != null && x.Owner.Length != 0);
            Assert.NotEmpty(document.SystemTables["MSysACEs"].Columns);
            AccessTable table = document.Tables["Legacy"]; Assert.True(table.Columns["Link"].IsHyperlink);
            using AccessDataReader rows = table.OpenDataReader(); Assert.True(rows.Read());
            Assert.Equal("Zażółć 漢字", rows["Label"]);
            Assert.Equal("Legacy Memo", rows["Notes"]);
            Assert.Equal("Caption#https://example.invalid/path#section#Tooltip", rows["Link"]);
            Assert.Equal(new DateTime(2000, 2, 29, 12, 34, 56), rows["Occurred"]);
            Assert.True(document.Queries["LegacyNames"].HasSql); Assert.Equal("SELECT Id, Label\nFROM [Legacy]\nORDER BY Id;", document.Queries["LegacyNames"].Sql);
            Assert.True(document.Queries["LegacyRawUnion"].HasSql); Assert.Contains("UNION ALL", document.Queries["LegacyRawUnion"].Sql);
            Assert.Equal("SELECT Id+1 AS [NextId], *\nFROM [Legacy];", document.Queries["LegacyStarExpression"].Sql);
            AccessQueryDefinition grouped = document.Queries["LegacyGrouped"]; Assert.False(grouped.HasSql);
            Assert.NotEmpty(grouped.NativeRecords); Assert.Throws<NotSupportedException>(() => grouped.Sql);
            AccessQueryDefinition sized = document.Queries["LegacySizedParameter"]; Assert.False(sized.HasSql);
            Assert.Contains(sized.NativeRecords, x => x.Attribute == 2 && x.Extra == 8);
            Assert.Throws<NotSupportedException>(() => sized.Sql);
            AccessQueryDefinition qualified = document.Queries["LegacyQualifiedSource"]; Assert.False(qualified.HasSql);
            Assert.Contains(qualified.NativeRecords, x => x.Attribute == 4 && x.Name1 != null || x.Attribute == 5 && x.Expression != null);
            AccessTable link = document.Tables["LocalLink"]; Assert.True(link.IsLinked); Assert.Equal("Legacy", link.LinkedTable!.ForeignTableName);
            Assert.Equal(AccessCatalogStatus.NotDecoded, link.Columns.CatalogStatus);
            Assert.Throws<NotSupportedException>(() => link.OpenDataReader()); Assert.Throws<NotSupportedException>(() => link.RowCount);
            AccessLinkedTableInfo credential = document.Tables["CredentialLink"].LinkedTable!;
            Assert.DoesNotContain("Fixture123", credential.Connection!); Assert.Contains("[redacted]", credential.Connection!);
            using AccessDataReader catalog = document.SystemTables["MSysObjects"].OpenDataReader();
            while (catalog.Read()) if (Equals(catalog["Name"], "CredentialLink")) Assert.DoesNotContain("Fixture123", (string)catalog["Connect"]);
        }
        [Theory]
        [InlineData("extended-values.accdb", 0L)]
        [InlineData("fractional-values.accdb", 1234567L)]
        public void ModernProfileValuesAgreeWithIndependentDaoPrecisionObservation(string file, long fractionalTicks) {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Generations", file));
            Assert.Equal(AccessFormatProfile.Ace17, document.Profile);
            using AccessDataReader rows = document.Tables["Extended"].OpenDataReader(); Assert.True(rows.Read());
            Assert.Equal(new DateTime(2, 1, 2, 3, 4, 5).AddTicks(fractionalTicks), rows["Occurred"]);
            Assert.Equal(long.MaxValue, rows["Large"]);
            AccessOperationReport report = document.AssessSave("target.mdb");
            Assert.Contains(report.Diagnostics, x => x.Code == "access.conversion.loss.extended-date");
            Assert.Contains(report.Diagnostics, x => x.Code == "access.conversion.loss.large-number");
        }
        [Theory]
        [InlineData("jet4.mdb")]
        [InlineData("ace12.accdb")]
        public void NativeParameterQueryHasNormalizedSqlAndExactInertRecords(string file) {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
            AccessQueryDefinition query = document.Queries["ContactsByGroup"];
            Assert.Equal("PARAMETERS [selectedGroup] Long;\nSELECT Id, DisplayName\nFROM [Contacts]\nWHERE GroupId = selectedGroup\nORDER BY Id;", query.Sql);
            Assert.Equal("selectedGroup", Assert.Single(query.Parameters).Name); Assert.Equal(AccessDataType.Int32, query.Parameters[0].DataType);
            Assert.NotEmpty(query.NativeRecords); Assert.All(query.NativeRecords, x => Assert.NotEmpty(x.NativeBytes.GetBytes()));
        }
        [Fact]
        public void UnqualifiedCalculatedValuesRetainBytesAndDiagnosticsWithoutEvaluation() {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Generations/calculated-ace14.accdb"));
            Assert.Equal(AccessFormatProfile.Ace14, document.Profile); AccessTable table = document.Tables["Calculated"];
            Assert.True(table.Columns["Value"].IsCalculated);
            Assert.Contains(table.Columns["Value"].Diagnostics, x => x.Code == "access.value.opaque.calculated");
            using AccessDataReader rows = table.OpenDataReader(); Assert.True(rows.Read()); Assert.Equal(41, rows["Id"]);
            AccessOpaqueValue opaque = Assert.IsType<AccessOpaqueValue>(rows["Value"]); Assert.NotEmpty(opaque.GetBytes());
            byte[] bytes = opaque.GetBytes(); bytes[0] ^= 255; Assert.NotEqual(bytes[0], opaque.GetBytes()[0]);
            Assert.True(table.Columns["TextValue"].IsCalculated);
            Assert.IsType<AccessOpaqueValue>(rows["TextValue"]);
            Assert.Contains(table.Columns["TextValue"].Diagnostics, x => x.Code == "access.value.opaque.calculated");
        }
        [Fact]
        public void UserPayloadsAndRepeatedReadersUseValueBudgetWithoutConsumingMetadataBudget() {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Readers/values-ace.accdb"),
                new AccessLoadOptions { MaxMetadataBytes = 20000, MaxValueBytes = 100000 });
            for (int iteration = 0; iteration < 3; iteration++) {
                using AccessDataReader rows = document.Tables["Scalars"].OpenDataReader(); Assert.True(rows.Read());
                Assert.Equal(20000, Assert.IsType<byte[]>(rows["Payload"]).Length);
                Assert.Equal(string.Concat(Enumerable.Repeat("<div>Ł🙂 native long text</div>", 500)), rows["Notes"]);
                using AccessDataReader structured = document.Tables["Structured"].OpenDataReader(); Assert.True(structured.Read());
                AccessAttachment attachment = Assert.Single(((AccessComplexValue)structured["Files"]).EnumerateAttachments());
                Assert.Equal(256, attachment.GetBytes().Length);
            }
        }
        [Fact]
        public void AllDocumentStreamsObserveDisposalAndOriginatingCancellation() {
            using CancellationTokenSource cancellation = new CancellationTokenSource();
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Readers/values-ace.accdb"));
            using AccessDataReader rows = document.Tables["Scalars"].OpenDataReader(cancellation.Token); Assert.True(rows.Read());
            using Stream small = rows.GetStream(rows.GetOrdinal("FixedBinary")); using Stream large = rows.GetStream(rows.GetOrdinal("Payload"));
            using AccessDataReader children = document.Tables["Structured"].OpenDataReader(cancellation.Token); Assert.True(children.Read());
            AccessAttachment attachment = Assert.Single(((AccessComplexValue)children["Files"]).EnumerateAttachments(cancellation.Token)); using Stream file = attachment.OpenRead();
            Assert.True(small.ReadByte() >= 0); Assert.True(large.ReadByte() >= 0); Assert.True(file.ReadByte() >= 0);
            cancellation.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() => small.ReadByte()); Assert.ThrowsAny<OperationCanceledException>(() => large.ReadByte()); Assert.ThrowsAny<OperationCanceledException>(() => file.ReadByte());
            document.Dispose();
            Assert.Throws<ObjectDisposedException>(() => small.ReadByte()); Assert.Throws<ObjectDisposedException>(() => large.ReadByte()); Assert.Throws<ObjectDisposedException>(() => file.ReadByte());

            using AccessDocument model = AccessDocument.Create(); AccessTable table = model.Tables.Add("Binary"); table.Columns.Add("Data", AccessDataType.Binary);
            table.AppendRow(new AccessRowValues { ["Data"] = new byte[] { 1, 2 } }); using AccessDataReader modeledRows = table.OpenDataReader(); Assert.True(modeledRows.Read());
            using Stream modeledStream = modeledRows.GetStream(0); model.Dispose(); Assert.Throws<ObjectDisposedException>(() => modeledStream.ReadByte());
        }
        [Fact]
        public void ExtendedDateTimeModelValuesKeepClrPrecisionAndTargetLoss() {
            using AccessDocument model = AccessDocument.Create(); AccessTable table = model.Tables.Add("Modern"); table.Columns.Add("Occurred", AccessDataType.ExtendedDateTime);
            DateTime value = new DateTime(2, 1, 2, 3, 4, 5).AddTicks(1234567);
            table.AppendRow(new AccessRowValues { ["Occurred"] = value });
            using (AccessDataReader reader = table.OpenDataReader()) { Assert.Equal(typeof(DateTime), reader.GetFieldType(0)); Assert.True(reader.Read()); Assert.Equal(value, reader.GetDateTime(0)); }
            Assert.Contains(model.AssessSave("target.accdb").Diagnostics, x => x.Code == "access.conversion.loss.extended-date");
        }
        [Theory]
        [InlineData("jet4.mdb")]
        [InlineData("ace12.accdb")]
        public void NativeReaderWorksWithProviderNeutralDataTableConsumer(string file) {
            using AccessDocument document = AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
            using AccessDataReader reader = document.Tables["Contacts"].OpenDataReader(); using System.Data.DataTable table = new System.Data.DataTable();
            table.Load(reader); Assert.Equal(2, table.Rows.Count); Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
            Assert.Equal(12.3456m, table.Rows[0]["Amount"]); Assert.Equal(DBNull.Value, table.Rows[1]["Notes"]);
        }
    }
}
