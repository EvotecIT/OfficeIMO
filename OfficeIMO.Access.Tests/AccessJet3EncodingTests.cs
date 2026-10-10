using OfficeIMO.Access;
using System.Data;
using System.Text;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessJet3EncodingTests {
        [Theory]
        [InlineData("testV1997.mdb", "Table1", "A", AccessDataType.ShortText)]
        [InlineData("test2V1997.mdb", "MSP_PROJECTS", "PROJ_PROP_AUTHOR", AccessDataType.LongText)]
        public void UnqualifiedTextReaderMetadataKeepsExactValuesThroughDataTableLoad(string file, string tableName, string columnName, AccessDataType type) {
            byte[] bytes = Fixture(file);
            byte[] expected;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes))) {
                using AccessDataReader rows = original.Tables[tableName].OpenDataReader(); Assert.True(rows.Read());
                expected = Encoding.ASCII.GetBytes(rows.GetString(rows.GetOrdinal(columnName)));
            }
            int column = ColumnOffset(bytes, tableName, columnName);
            SetCodePage(bytes, column, 932);
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            Assert.Equal(type, document.Tables[tableName].Columns[columnName].DataType);
            using (AccessDataReader rows = document.Tables[tableName].OpenDataReader()) {
                int ordinal = rows.GetOrdinal(columnName); Assert.Equal(typeof(AccessOpaqueValue), rows.GetFieldType(ordinal));
                Assert.True(rows.Read()); Assert.Equal(expected, Assert.IsType<AccessOpaqueValue>(rows.GetValue(ordinal)).GetBytes());
            }
            using (AccessDataReader rows = document.Tables[tableName].OpenDataReader()) {
                var table = new DataTable(); table.Load(rows);
                Assert.Equal(typeof(AccessOpaqueValue), table.Columns[columnName]!.DataType);
                Assert.Equal(expected, Assert.IsType<AccessOpaqueValue>(table.Rows[0][columnName]).GetBytes());
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void UnqualifiedConnectionEncodingWithholdsCredentialsFromOrdinaryReaders(bool shortText) {
            byte[] bytes = ConnectionFixture(shortText);
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            AccessTable catalog = document.SystemTables["MSysObjects"];
            Assert.Contains(catalog.Columns["Connect"].Diagnostics, d => d.Code == "access.value.redacted.encoding");
            using AccessDataReader rows = catalog.OpenDataReader();
            int connection = rows.GetOrdinal("Connect"); Assert.Equal(typeof(string), rows.GetFieldType(connection));
            bool found = false;
            while (rows.Read()) {
                if (!Equals(rows["Name"], shortText ? "PWD=XX" : "Table1")) continue;
                found = true;
                string value = rows.GetString(connection);
                Assert.Equal("[redacted: unsupported connection encoding]", value);
                Assert.DoesNotContain(shortText ? "XX" : "FixtureOnly", value);
            }
            Assert.True(found);
            using var saved = new MemoryStream(); document.Save(saved); Assert.Equal(bytes, saved.ToArray());
            // Preservation remains an explicit raw-byte operation, including caller-owned sensitive input.
            Assert.Contains(shortText ? "PWD=XX" : "PWD=FixtureOnly", Encoding.ASCII.GetString(saved.ToArray()));
        }

        [Theory]
        [InlineData("Name1")]
        [InlineData("Name2")]
        [InlineData("Expression")]
        public void UnqualifiedQueryTextIsDiagnosedAndKeepsNativeRecords(string columnName) {
            byte[] bytes = Fixture("queryTestV1997.mdb");
            string queryName = columnName == "Name1" ? "UpdateQuery" : "UnionQuery";
            byte[][] expectedRecords;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes))) {
                Assert.True(original.Queries["UnionQuery"].HasSql);
                Assert.Equal("User Name", Assert.Single(original.Queries["UpdateQuery"].Parameters).Name);
                expectedRecords = original.Queries[queryName].NativeRecords.Select(r => r.NativeBytes.GetBytes()).ToArray();
            }
            SetCodePage(bytes, ColumnOffset(bytes, "MSysQueries", columnName), 932);
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            AccessQueryDefinition query = document.Queries[queryName];
            Assert.False(query.HasSql); Assert.Throws<NotSupportedException>(() => query.Sql);
            Assert.Contains(query.Diagnostics, d => d.Code == "access.query.text-encoding-opaque");
            Assert.Equal(expectedRecords.Length, query.NativeRecords.Count);
            for (int i = 0; i < expectedRecords.Length; i++) Assert.Equal(expectedRecords[i], query.NativeRecords[i].NativeBytes.GetBytes());
            using AccessDataReader rows = document.SystemTables["MSysQueries"].OpenDataReader();
            Assert.Equal(typeof(AccessOpaqueValue), rows.GetFieldType(rows.GetOrdinal(columnName)));
            bool opaque = false; while (rows.Read()) if (rows[columnName] is AccessOpaqueValue value) { Assert.NotEmpty(value.GetBytes()); opaque = true; }
            Assert.True(opaque);
            if (columnName == "Name1") {
                Assert.Empty(document.Queries["UpdateQuery"].Parameters);
                Assert.Contains(document.Queries["UpdateQuery"].Diagnostics, d => d.Code == "access.query.text-encoding-opaque");
            }
        }

        [Fact]
        public void UnqualifiedLinkedTableNameIsDiagnosedAndNeverResolved() {
            byte[] bytes = Fixture("testV1997.mdb");
            byte[] row;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes)))
                row = original.Catalog.Single(e => e.Name == "Table1" && e.NativeType == 1).NativeRecord!.GetBytes();
            int rowStart = CatalogRowOffset(bytes, row);
            int foreign = ColumnOffset(bytes, "MSysObjects", "ForeignName"), name = ColumnOffset(bytes, "MSysObjects", "Name"), type = ColumnOffset(bytes, "MSysObjects", "Type");
            CopyVariableIndex(bytes, name, foreign); SetCodePage(bytes, foreign, 932);
            SetNonNull(bytes, rowStart, row.Length, foreign);
            // Turn the controlled catalog entry into an inert linked-table definition.
            int typeOffset = rowStart + 1 + BitConverter.ToUInt16(bytes, type + 14);
            bytes[typeOffset] = 6; bytes[typeOffset + 1] = 0;
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            AccessTable table = document.Tables["Table1"];
            Assert.True(table.IsLinked); Assert.Null(table.LinkedTable!.ForeignTableName);
            Assert.Contains(table.Diagnostics, d => d.Code == "access.linked-table.encoding-opaque");
            Assert.Throws<NotSupportedException>(() => table.OpenDataReader());
            using AccessDataReader rows = document.SystemTables["MSysObjects"].OpenDataReader();
            while (rows.Read()) if (Equals(rows["Name"], "Table1")) Assert.Equal(Encoding.ASCII.GetBytes("Table1"), Assert.IsType<AccessOpaqueValue>(rows["ForeignName"]).GetBytes());
        }

        private static byte[] ConnectionFixture(bool shortText) {
            byte[] bytes = Fixture("testV1997.mdb");
            byte[] row, properties;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes))) {
                row = original.Catalog.Single(e => e.Name == "Table1" && e.NativeType == 1).NativeRecord!.GetBytes();
                properties = original.Tables["Table1"].NativeProperties!.GetBytes();
            }
            int rowStart = CatalogRowOffset(bytes, row);
            int connection = ColumnOffset(bytes, "MSysObjects", "Connect");
            int value = ColumnOffset(bytes, "MSysObjects", shortText ? "Name" : "LvProp");
            CopyVariableIndex(bytes, value, connection); SetCodePage(bytes, connection, 932);
            SetNonNull(bytes, rowStart, row.Length, connection);
            if (shortText) {
                bytes[connection] = 10; bytes[connection + 16] = 255; bytes[connection + 17] = 0;
                int name = Find(bytes, Encoding.ASCII.GetBytes("Table1"), rowStart, rowStart + row.Length);
                Buffer.BlockCopy(Encoding.ASCII.GetBytes("PWD=XX"), 0, bytes, name, 6);
            } else {
                // Use this independent fixture's actual Memo-compatible long-value carrier as controlled Connect input.
                int property = Find(bytes, properties, 0, bytes.Length);
                for (int i = 0; i < properties.Length; i++) bytes[property + i] = (byte)' ';
                byte[] text = Encoding.ASCII.GetBytes("PWD=FixtureOnly"); Buffer.BlockCopy(text, 0, bytes, property, text.Length);
            }
            return bytes;
        }
        private static int ColumnOffset(byte[] bytes, string tableName, string columnName) {
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            AccessTable table = document.Tables.Concat(document.SystemTables).Single(t => t.Name == tableName);
            int ordinal = table.Columns.Select(c => c.Name).ToList().IndexOf(columnName); Assert.True(ordinal >= 0);
            int page = document.Catalog.Single(e => e.Name == tableName && e.NativeType == 1).NativeId & 0x00ffffff;
            int start = page * 2048;
            int column = start + 43 + BitConverter.ToInt32(bytes, start + 31) * 8 + ordinal * 18;
            Assert.True(column + 18 <= start + 2048); return column;
        }
        private static int CatalogRowOffset(byte[] bytes, byte[] row) {
            int found = -1;
            for (int page = 1; page < bytes.Length / 2048; page++) {
                int p = page * 2048;
                if (bytes[p] != 1 || BitConverter.ToInt32(bytes, p + 4) != 2) continue;
                int count = BitConverter.ToUInt16(bytes, p + 8);
                for (int slot = 0; slot < count; slot++) {
                    int flags = BitConverter.ToUInt16(bytes, p + 10 + slot * 2), start = flags & 0x1fff;
                    int end = slot == 0 ? 2048 : BitConverter.ToUInt16(bytes, p + 8 + slot * 2) & 0x1fff;
                    if ((flags & 0xc000) != 0 || end - start != row.Length || !row.Select((b, i) => bytes[p + start + i] == b).All(equal => equal)) continue;
                    Assert.Equal(-1, found); found = p + start;
                }
            }
            Assert.True(found >= 0); return found;
        }
        private static int Find(byte[] bytes, byte[] pattern, int start, int end) => Enumerable.Range(start, end - start - pattern.Length + 1)
            .First(at => pattern.Select((b, i) => bytes[at + i] == b).All(equal => equal));
        private static void SetNonNull(byte[] bytes, int rowStart, int length, int column) {
            int number = BitConverter.ToUInt16(bytes, column + 1), nullBytes = (bytes[rowStart] + 7) / 8;
            bytes[rowStart + length - nullBytes + number / 8] |= (byte)(1 << (number % 8));
        }
        private static void CopyVariableIndex(byte[] bytes, int source, int target) => Buffer.BlockCopy(bytes, source + 3, bytes, target + 3, 2);
        private static void SetCodePage(byte[] bytes, int column, int codePage) { bytes[column + 11] = (byte)codePage; bytes[column + 12] = (byte)(codePage >> 8); }
        private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Jet3", name));
    }
}
