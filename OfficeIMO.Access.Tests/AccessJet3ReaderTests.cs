using OfficeIMO.Access;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessJet3ReaderTests {
        [Theory]
        [InlineData("testV1997.mdb")]
        [InlineData("test2V1997.mdb")]
        [InlineData("overflowTestV1997.mdb")]
        [InlineData("compIndexTestV1997.mdb")]
        [InlineData("queryTestV1997.mdb")]
        [InlineData("delColTestV1997.mdb")]
        public void IndependentJet3SchemaIndexesAndTypedRowsAgreeWithJackcess(string file) {
            XElement expected = Oracle(file);
            using AccessDocument document = AccessDocument.Load(Fixture(file));
            Assert.Equal(AccessFormatProfile.Jet3, document.Profile);
            Assert.Equal(AccessCatalogStatus.Decoded, document.CatalogStatus);
            Assert.Equal(1252, document.CodePage);
            Assert.Equal(1033, document.SortOrder);
            Assert.Equal(expected.Elements("Table").Count(), document.Tables.Count);
            foreach (XElement wanted in expected.Elements("Table")) {
                AccessTable table = document.Tables[Attribute(wanted, "name")];
                Assert.Equal(long.Parse(Attribute(wanted, "rowCount"), CultureInfo.InvariantCulture), table.RowCount);
                XElement[] columns = wanted.Elements("Column").ToArray();
                Assert.Equal(columns.Select(c => Attribute(c, "name")), table.Columns.Select(c => c.Name));
                foreach (XElement column in columns) {
                    AccessColumn actual = table.Columns[Attribute(column, "name")];
                    bool autoNumber = bool.Parse(Attribute(column, "autoNumber"));
                    Assert.Equal(NativeType(Attribute(column, "type"), autoNumber), actual.DataType);
                    Assert.Equal(autoNumber, actual.IsAutoNumber);
                    if (actual.DataType == AccessDataType.ShortText)
                        Assert.Equal(int.Parse(Attribute(column, "length"), CultureInfo.InvariantCulture), actual.MaxLength);
                }
                Assert.Equal(wanted.Elements("Index").Count(), table.Indexes.Count);
                foreach (XElement index in wanted.Elements("Index")) {
                    AccessIndex actual = table.Indexes[Attribute(index, "name")];
                    Assert.Equal(bool.Parse(Attribute(index, "primary")), actual.IsPrimaryKey);
                    Assert.Equal(bool.Parse(Attribute(index, "unique")), actual.IsUnique);
                    Assert.Equal(bool.Parse(Attribute(index, "foreign")), actual.IsForeignKey);
                    Assert.Equal(index.Elements("Column").Select(c => Attribute(c, "name")), actual.Columns.Select(c => c.Name));
                    Assert.Equal(index.Elements("Column").Select(c => bool.Parse(Attribute(c, "descending"))), actual.Descending);
                }
                using AccessDataReader reader = table.OpenDataReader();
                foreach (XElement row in wanted.Elements("Row")) {
                    Assert.True(reader.Read());
                    foreach (XElement field in row.Elements("Field")) {
                        int ordinal = reader.GetOrdinal(Attribute(field, "name"));
                        AssertField(field, reader.IsDBNull(ordinal) ? null : reader.GetValue(ordinal));
                    }
                }
                Assert.False(reader.Read());
            }
            Assert.Equal(AccessCatalogStatus.NotDecoded, document.VbaProject.CatalogStatus);
        }

        [Fact]
        public void Jet3QueriesKeepIndependentIdentitiesAndInertRecords() {
            using AccessDocument document = AccessDocument.Load(Fixture("queryTestV1997.mdb"));
            XElement[] expected = Oracle("queryTestV1997.mdb").Elements("Query").ToArray();
            Assert.Equal(expected.Length, document.Queries.Count);
            foreach (XElement wanted in expected) {
                AccessQueryDefinition query = document.Queries[Attribute(wanted, "name")];
                Assert.Equal(int.Parse(Attribute(wanted, "flags"), CultureInfo.InvariantCulture), query.NativeFlags);
                Assert.Equal(int.Parse(Attribute(wanted, "records"), CultureInfo.InvariantCulture), query.NativeRecords.Count);
                Assert.All(query.NativeRecords, record => Assert.NotEmpty(record.NativeBytes.GetBytes()));
                if (query.Name == "UnionQuery") {
                    Assert.True(query.HasSql);
                    Assert.Equal(Whitespace(wanted.Element("Sql")!.Value), Whitespace(query.Sql));
                } else {
                    Assert.False(query.HasSql);
                    Assert.Throws<NotSupportedException>(() => query.Sql);
                }
            }
            Assert.Equal("User Name", Assert.Single(document.Queries["UpdateQuery"].Parameters).Name);
        }

        [Fact]
        public async Task Jet3StreamsSelectionLazyLimitsAndCancellationKeepTheExistingApiContract() {
            byte[] bytes = File.ReadAllBytes(Fixture("testV1997.mdb"));
            using var input = new MemoryStream(bytes); input.Position = 17;
            using AccessDocument document = await AccessDocument.LoadAsync(input, new AccessLoadOptions { TableNames = new[] { "Table1" }, MaxRows = 1 });
            Assert.Equal(17, input.Position); Assert.True(input.CanRead); Assert.Single(document.Tables);
            using AccessDataReader rows = document.Tables[0].OpenDataReader();
            Assert.True(rows.Read()); Assert.Equal("a", rows["A"]);
            Assert.Throws<InvalidDataException>(() => rows.Read());
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => AccessDocument.LoadAsync(input, cancellationToken: canceled.Token));
            Assert.Throws<InvalidDataException>(() => AccessDocument.Load(new MemoryStream(bytes), new AccessLoadOptions { MaxPages = bytes.Length / 2048 - 1 }));
            Assert.Throws<InvalidDataException>(() => AccessDocument.Load(new MemoryStream(bytes), new AccessLoadOptions { MaxMetadataBytes = 63 }));
            Assert.Throws<InvalidDataException>(() => AccessDocument.Load(new MemoryStream(bytes), new AccessLoadOptions { MaxCatalogObjects = 1 }));
        }

        [Fact]
        public void Jet3LongBinaryStreamsEnforceValueLimitsLifetimeAndCancellation() {
            string file = Fixture("test2V1997.mdb");
            using (AccessDocument limited = AccessDocument.Load(file, new AccessLoadOptions { MaxValueBytes = 128 })) {
                using AccessDataReader rows = limited.Tables["MSP_PROJECTS"].OpenDataReader(); Assert.True(rows.Read());
                Assert.Throws<InvalidDataException>(() => rows.GetStream(rows.GetOrdinal("RESERVED_BINARY_DATA")));
            }
            using var cancellation = new CancellationTokenSource();
            using AccessDocument document = AccessDocument.Load(file);
            using AccessDataReader reader = document.Tables["MSP_PROJECTS"].OpenDataReader(cancellation.Token); Assert.True(reader.Read());
            using Stream stream = reader.GetStream(reader.GetOrdinal("RESERVED_BINARY_DATA"));
            Assert.False(stream.CanWrite); Assert.True(stream.ReadByte() >= 0);
            cancellation.Cancel(); Assert.ThrowsAny<OperationCanceledException>(() => stream.ReadByte());
            document.Dispose(); Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }

        [Theory]
        [InlineData("testV1997.mdb")]
        [InlineData("queryTestV1997.mdb")]
        public void Jet3UnchangedPreservationIsExactAndDoesNotEnableEditingOrConversion(string file) {
            byte[] bytes = File.ReadAllBytes(Fixture(file));
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            Assert.Throws<NotSupportedException>(() => document.Tables.Add("NewTable"));
            Assert.False(document.Capabilities.Single(capability => capability.Operation == "vba.edit").IsSupported);
            Assert.False(document.Capabilities.Single(capability => capability.Operation == "application.events.write").IsSupported);
            Assert.Throws<NotSupportedException>(() => document.GetVbaProject());
            OfficeVbaProject project = OfficeVbaProject.Create("Jet3Unavailable");
            Assert.Throws<NotSupportedException>(() => document.SetVbaProject(project));
            Assert.Equal(0, document.Revision);
            using var output = new MemoryStream(); document.Save(output); Assert.Equal(bytes, output.ToArray());
            output.Position = 0;
            using AccessDocument reopened = AccessDocument.Load(output); Assert.Equal(document.Tables.Select(t => t.Name), reopened.Tables.Select(t => t.Name));
            using var conversion = new MemoryStream();
            Assert.Throws<AccessOperationNotSupportedException>(() => document.Save(conversion, new AccessSaveOptions { Format = AccessFileFormat.Accdb }));
            Assert.Equal(0, conversion.Length);
        }

        [Theory]
        [InlineData(66, "password")]
        [InlineData(62, "encryption")]
        [InlineData(60, "code page")]
        public void ProtectedAndUnsupportedEncodingHeadersRemainPreserveOnly(int offset, string reason) {
            byte[] bytes = File.ReadAllBytes(Fixture("testV1997.mdb"));
            if (offset == 60) { ushort unsupported = 65001; bytes[60] ^= (byte)(1252 ^ unsupported); bytes[61] ^= (byte)((1252 ^ unsupported) >> 8); }
            else bytes[offset] ^= 1; // RC4 header masking preserves this change in the decoded byte.
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            Assert.Equal(AccessCatalogStatus.NotDecoded, document.CatalogStatus);
            Assert.Contains(document.Diagnostics, diagnostic => diagnostic.Message.Contains(reason));
            Assert.Empty(document.Tables);
            using var output = new MemoryStream(); document.Save(output); Assert.Equal(bytes, output.ToArray());
        }

        [Fact]
        public void UnsupportedJet3ColumnEncodingPreservesOpaqueBytesWithoutLosingTheCatalog() {
            byte[] bytes = File.ReadAllBytes(Fixture("testV1997.mdb"));
            SetFirstColumnCodePage(bytes, 65001);
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            Assert.Equal(AccessCatalogStatus.Decoded, document.CatalogStatus);
            Assert.Contains(document.Tables["Table1"].Columns["A"].Diagnostics, diagnostic => diagnostic.Code == "access.value.opaque.encoding");
            using AccessDataReader rows = document.Tables["Table1"].OpenDataReader(); Assert.True(rows.Read());
            Assert.Equal(new byte[] { (byte)'a' }, Assert.IsType<AccessOpaqueValue>(rows["A"]).GetBytes());
            Assert.Equal("b", rows["B"]);
        }

        [Theory]
        [InlineData(1250, "Łbcdefg")]
        [InlineData(1252, "£bcdefg")]
        [InlineData(437, "úbcdefg")]
        [InlineData(850, "úbcdefg")]
        [InlineData(10000, "£bcdefg")]
        public void Jet3ColumnCodePagesUseTheQualifiedSingleByteMaps(int codePage, string expected) {
            byte[] bytes = File.ReadAllBytes(Fixture("testV1997.mdb"));
            SetFirstColumnCodePage(bytes, codePage);
            byte[] label = System.Text.Encoding.ASCII.GetBytes("abcdefg");
            int[] positions = Enumerable.Range(0, bytes.Length - label.Length + 1).Where(at => label.Select((value, offset) => bytes[at + offset] == value).All(match => match)).ToArray();
            bytes[Assert.Single(positions)] = 0xa3;
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            using AccessDataReader rows = document.Tables["Table1"].OpenDataReader(); Assert.True(rows.Read()); Assert.Equal("a", rows["A"]);
            Assert.True(rows.Read()); Assert.Equal(expected, rows["A"]);
        }

        [Fact]
        public void MalformedJet3JumpIndicesAreRejectedWhenTheUserRowIsReached() {
            byte[] bytes = File.ReadAllBytes(Fixture("test2V1997.mdb"));
            int definition;
            using (AccessDocument source = AccessDocument.Load(new MemoryStream(bytes)))
                definition = source.Catalog.Single(entry => entry.Name == "MSP_PROJECTS" && entry.NativeType == 1).NativeId & 0x00ffffff;
            bool changed = false;
            for (int page = 1; page < bytes.Length / 2048 && !changed; page++) {
                int p = page * 2048;
                if (bytes[p] != 1 || BitConverter.ToInt32(bytes, p + 4) != definition) continue;
                int count = BitConverter.ToUInt16(bytes, p + 8);
                for (int slot = 0; slot < count; slot++) {
                    int flags = BitConverter.ToUInt16(bytes, p + 10 + slot * 2);
                    if ((flags & 0xc000) != 0) continue;
                    int start = flags & 0x1fff, end = slot == 0 ? 2048 : BitConverter.ToUInt16(bytes, p + 8 + slot * 2) & 0x1fff;
                    if (end - start <= 256) continue;
                    int nullBytes = (bytes[p + start] + 7) / 8;
                    int variableCount = p + end - nullBytes - 1;
                    bytes[variableCount - 1] = (byte)(bytes[variableCount] + 1);
                    changed = true; break;
                }
            }
            Assert.True(changed);
            using AccessDocument document = AccessDocument.Load(new MemoryStream(bytes));
            using AccessDataReader rows = document.Tables["MSP_PROJECTS"].OpenDataReader();
            Assert.Throws<InvalidDataException>(() => rows.Read());
        }

        private static string Fixture(string file) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Jet3", file);
        private static void SetFirstColumnCodePage(byte[] bytes, int codePage) {
            int definition;
            using (AccessDocument source = AccessDocument.Load(new MemoryStream(bytes)))
                definition = source.Catalog.Single(entry => entry.Name == "Table1" && entry.NativeType == 1).NativeId & 0x00ffffff;
            int header = definition * 2048;
            int column = header + 43 + BitConverter.ToInt32(bytes, header + 31) * 8;
            Assert.Equal(10, bytes[column]);
            bytes[column + 11] = (byte)codePage; bytes[column + 12] = (byte)(codePage >> 8);
        }
        private static XElement Oracle(string file) => XDocument.Load(Fixture("expected.xml")).Root!.Elements("File").Single(element => Attribute(element, "name") == file);
        private static string Attribute(XElement element, string name) => element.Attribute(name)!.Value;
        private static string Whitespace(string value) => Regex.Replace(value, @"\s+", " ").Trim();
        private static AccessDataType NativeType(string type, bool autoNumber) => type switch {
            "BOOLEAN" => AccessDataType.Boolean, "BYTE" => AccessDataType.Byte, "INT" => AccessDataType.Int16,
            "LONG" => autoNumber ? AccessDataType.AutoNumber : AccessDataType.Int32, "MONEY" => AccessDataType.Currency,
            "DOUBLE" => AccessDataType.Double, "SHORT_DATE_TIME" => AccessDataType.DateTime, "TEXT" => AccessDataType.ShortText,
            "MEMO" => AccessDataType.LongText, "OLE" => AccessDataType.Binary, "GUID" => AccessDataType.Guid,
            _ => throw new InvalidDataException("Unexpected oracle type: " + type)
        };
        private static void AssertField(XElement expected, object? actual) {
            string value = expected.Value;
            switch (Attribute(expected, "kind")) {
                case "null": Assert.Null(actual); break;
                case "String": Assert.Equal(value, Assert.IsType<string>(actual)); break;
                case "Byte": Assert.Equal(byte.Parse(value, CultureInfo.InvariantCulture), Assert.IsType<byte>(actual)); break;
                case "Short": Assert.Equal(short.Parse(value, CultureInfo.InvariantCulture), Assert.IsType<short>(actual)); break;
                case "Integer": Assert.Equal(int.Parse(value, CultureInfo.InvariantCulture), Assert.IsType<int>(actual)); break;
                case "Double": Assert.Equal(double.Parse(value, CultureInfo.InvariantCulture), Assert.IsType<double>(actual)); break;
                case "BigDecimal": Assert.Equal(decimal.Parse(value, CultureInfo.InvariantCulture), Assert.IsType<decimal>(actual)); break;
                case "Boolean": Assert.Equal(bool.Parse(value), Assert.IsType<bool>(actual)); break;
                case "date": Assert.Equal(DateTimeOffset.Parse(value, CultureInfo.InvariantCulture).DateTime, Assert.IsType<DateTime>(actual)); break;
                case "bytes":
                    byte[] bytes = Assert.IsType<byte[]>(actual); Assert.Equal(int.Parse(Attribute(expected, "length"), CultureInfo.InvariantCulture), bytes.Length);
                    using (SHA256 hash = SHA256.Create()) Assert.Equal(Attribute(expected, "sha256"), BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant());
                    break;
                default: throw new InvalidDataException("Unexpected oracle value kind.");
            }
        }
    }
}
