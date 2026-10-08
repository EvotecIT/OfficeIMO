using OfficeIMO.Access;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessNativeSafetyTests {
        private static byte[] Source(string name = "Readers/values-ace.accdb") => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
        private static int I32(byte[] bytes, int offset) => BitConverter.ToInt32(bytes, offset);
        private static int U16(byte[] bytes, int offset) => BitConverter.ToUInt16(bytes, offset);
        private static void Write32(byte[] bytes, int offset, int value) => Array.Copy(BitConverter.GetBytes(value), 0, bytes, offset, 4);
        private static int Definition(byte[] bytes, string field) {
            byte[] name = Encoding.Unicode.GetBytes(field);
            for (int page = 1; page < bytes.Length / 4096; page++) {
                int start = page * 4096; if (bytes[start] != 2) continue;
                int count = U16(bytes, start + 45), physicalIndexes = I32(bytes, start + 51), position = start + 63 + physicalIndexes * 12 + count * 25;
                for (int i = 0; i < count; i++) { int length = U16(bytes, position); position += 2; if (length == name.Length && bytes.Skip(position).Take(length).SequenceEqual(name)) return page; position += length; }
            }
            throw new InvalidOperationException("The independently produced fixture field was not found.");
        }
        [Theory]
        [InlineData("LvProp", "Type", "TypX")]
        [InlineData("Attribute", "ObjectId", "ObjectIX")]
        [InlineData("szRelationship", "grbit", "grbix")]
        [InlineData("ComplexID", "ComplexID", "ComplexIX")]
        public void RequiredSystemCatalogFieldsCannotBecomeEmptyDecodedInventories(string marker, string field, string replacement) {
            byte[] bytes = Source(); int start = Definition(bytes, marker) * 4096;
            int count = U16(bytes, start + 45), indexes = I32(bytes, start + 51), position = start + 63 + indexes * 12 + count * 25;
            bool changed = false;
            for (int i = 0; i < count; i++) {
                int length = U16(bytes, position); position += 2;
                if (Encoding.Unicode.GetString(bytes, position, length).Equals(field, StringComparison.OrdinalIgnoreCase)) {
                    byte[] name = Encoding.Unicode.GetBytes(replacement); Assert.Equal(length, name.Length); Array.Copy(name, 0, bytes, position, length); changed = true;
                }
                position += length;
            }
            Assert.True(changed); using MemoryStream input = new MemoryStream(bytes);
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => AccessDocument.Load(input)); Assert.Contains("required field", error.Message);
        }
        [Fact]
        public void UnknownPropertiesRetainBytesAndProduceNamedProfileMappingLoss() {
            byte[] bytes = Source(); int changed = 0;
            for (int i = 0; i < bytes.Length - 3; i++) if (bytes[i] == 'M' && bytes[i + 1] == 'R' && bytes[i + 2] == '2' && bytes[i + 3] == 0) { bytes[i] = (byte)'X'; changed++; }
            Assert.True(changed > 0); using MemoryStream input = new MemoryStream(bytes); using AccessDocument document = AccessDocument.Load(input);
            AccessTable table = document.Tables["Scalars"]; Assert.NotNull(table.NativeProperties); Assert.Equal((byte)'X', table.NativeProperties!.GetBytes()[0]);
            Assert.Contains(table.Diagnostics, x => x.Code == "access.properties.opaque");
            Assert.Contains(document.AssessSave("target.mdb").Diagnostics, x => x.Code == "access.conversion.loss.opaque-properties" && x.ObjectId == table.Id);
            using AccessDataReader rows = table.OpenDataReader(); Assert.True(rows.Read()); Assert.Equal(int.MaxValue, rows["Whole"]);
        }
        [Theory]
        [InlineData("szRelationship")]
        [InlineData("szReferencedObject")]
        [InlineData("szReferencedColumn")]
        [InlineData("szObject")]
        [InlineData("szColumn")]
        public void NullRelationshipNamesRejectEvenWhenTablesAreExcluded(string field) {
            byte[] bytes = Source(); (int Start, int End, int Number, int VariableIndex) row = RelationshipRow(bytes, field);
            int nullBytes = (U16(bytes, row.Start) + 7) / 8;
            bytes[row.End - nullBytes + row.Number / 8] &= (byte)~(1 << (row.Number % 8));
            using MemoryStream input = new MemoryStream(bytes);
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => AccessDocument.Load(input, new AccessLoadOptions { TableNames = new[] { "Scalars" } }));
            Assert.Contains("missing or null", error.Message);
        }
        [Fact]
        public void CompositeRelationshipMembersCannotNameDifferentTables() {
            byte[] bytes = Source(); (int Start, int End, int Number, int VariableIndex) row = RelationshipRow(bytes, "szReferencedObject");
            int nullBytes = (U16(bytes, row.Start) + 7) / 8, offset = row.End - nullBytes - 4 - row.VariableIndex * 2;
            int position = row.Start + U16(bytes, offset);
            // Change one member's table identity, accepting either native Unicode representation in the fixture.
            if (bytes[position] == 255 && bytes[position + 1] == 254) position += 2;
            Assert.Equal((byte)'P', bytes[position]); bytes[position] = (byte)'X';
            using MemoryStream input = new MemoryStream(bytes);
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => AccessDocument.Load(input)); Assert.Contains("inconsistent", error.Message);
        }
        [Theory]
        [InlineData("database")]
        [InlineData("table")]
        [InlineData("column")]
        public void OpaqueValuesInEveryPropertyOwnerProduceNamedMappingLoss(string owner) {
            byte[] bytes = Source("Generations/access2000.mdb"); string property = owner == "database" ? "AccessVersion" : "AllowZeroLength";
            Assert.True(MakePropertyOpaque(bytes, property, owner == "table") > 0);
            using MemoryStream input = new MemoryStream(bytes); using AccessDocument document = AccessDocument.Load(input); AccessTable table = document.Tables["Legacy"];
            IReadOnlyDictionary<string, object?> properties = owner == "database" ? document.Properties : owner == "table" ? table.Properties : table.Columns["Link"].Properties;
            Assert.Equal(99u, Assert.IsType<AccessOpaqueValue>(properties[property]).NativeType);
            Assert.Empty(document.Diagnostics); Assert.Empty(table.Diagnostics); Assert.NotNull(document.NativeProperties); Assert.NotNull(table.NativeProperties);
            Guid identity = owner == "database" ? document.Id : table.Id;
            Assert.Contains(document.AssessSave("target.accdb").Diagnostics, x => x.Code == "access.conversion.loss.opaque-properties" && x.ObjectId == identity);
            using AccessDataReader rows = table.OpenDataReader(); Assert.True(rows.Read()); Assert.Equal("Zażółć 漢字", rows["Label"]);
        }
        private static (int Start, int End, int Number, int VariableIndex) RelationshipRow(byte[] bytes, string field) {
            int table = Definition(bytes, "szRelationship"), start = table * 4096, count = U16(bytes, start + 45), indexes = I32(bytes, start + 51);
            int columns = start + 63 + indexes * 12, position = columns + count * 25, number = -1, variable = -1;
            for (int i = 0; i < count; i++) {
                int length = U16(bytes, position); position += 2;
                if (Encoding.Unicode.GetString(bytes, position, length) == field) { number = U16(bytes, columns + i * 25 + 5); variable = U16(bytes, columns + i * 25 + 7); }
                position += length;
            }
            Assert.True(number >= 0);
            for (int page = 1; page < bytes.Length / 4096; page++) {
                int offset = page * 4096; if (bytes[offset] != 1 || I32(bytes, offset + 4) != table) continue;
                for (int slot = 0; slot < U16(bytes, offset + 12); slot++) {
                    int record = U16(bytes, offset + 14 + slot * 2); if ((record & 0xe000) != 0) continue;
                    return (offset + record, slot == 0 ? offset + 4096 : offset + (U16(bytes, offset + 12 + slot * 2) & 8191), number, variable);
                }
            }
            throw new InvalidOperationException("The independently produced relationship row was not found.");
        }
        private static int MakePropertyOpaque(byte[] bytes, string property, bool tableDefault) {
            int changed = 0;
            for (int signature = 0; signature < bytes.Length - 4; signature++) {
                if (bytes[signature] != 'M' || bytes[signature + 1] != 'R' || bytes[signature + 2] != '2' || bytes[signature + 3] != 0) continue;
                List<string> names = new System.Collections.Generic.List<string>(); int position = signature + 4;
                while (position <= bytes.Length - 6) {
                    int length = I32(bytes, position), type = U16(bytes, position + 4);
                    if (length < 6 || length > bytes.Length - position || type != 128 && type != 0 && type != 1 && type != 2) break;
                    int block = position + 6, end = position + length;
                    if (type == 128) {
                        names.Clear(); for (int item = block; item < end;) { int size = U16(bytes, item); item += 2; names.Add(Encoding.Unicode.GetString(bytes, item, size)); item += size; }
                    } else if (length > 6) {
                        int nameLength = I32(bytes, block), first = block + nameLength, last = first; bool found = false;
                        for (int value = first; value < end; value += U16(bytes, value)) {
                            last = value;
                            if (names[U16(bytes, value + 4)] == property) { bytes[value + 3] = 99; changed++; found = true; }
                        }
                        if (found && tableDefault) {
                            // Promote this valid column map to a table default map without changing its native LVAL length.
                            int gap = nameLength - 6; Assert.True(gap > 0); bytes[position + 4] = 0; Write32(bytes, block, 6); bytes[block + 4] = bytes[block + 5] = 0;
                            Array.Copy(bytes, first, bytes, block + 6, end - first); int finalValue = last - gap;
                            Array.Copy(BitConverter.GetBytes((ushort)(U16(bytes, finalValue) + gap)), 0, bytes, finalValue, 2);
                        }
                    }
                    position = end;
                }
            }
            return changed;
        }
        [Fact]
        public void CyclicAndTruncatedDefinitionChainsRejectBeforeUserDataRead() {
            byte[] cyclic = Source(); int definition = Definition(cyclic, "Precise") * 4096;
            Write32(cyclic, definition + 8, 4097); Write32(cyclic, definition + 4, definition / 4096);
            using MemoryStream cycle = new MemoryStream(cyclic); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(cycle)); Assert.True(cycle.CanRead);
            byte[] truncated = Source(); definition = Definition(truncated, "Precise") * 4096;
            Write32(truncated, definition + 8, 4097); Write32(truncated, definition + 4, 0);
            using MemoryStream shortChain = new MemoryStream(truncated); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(shortChain));
        }
        [Fact]
        public void AmbiguousNativeFieldNamesAndMetadataBudgetsRejectDeterministically() {
            byte[] bytes = Source(); int start = Definition(bytes, "Precise") * 4096;
            int count = U16(bytes, start + 45), indexes = I32(bytes, start + 51), position = start + 63 + indexes * 12 + count * 25;
            for (int i = 0; i < count; i++) { int length = U16(bytes, position); position += 2; if (Encoding.Unicode.GetString(bytes, position, length) == "Tiny") Array.Copy(Encoding.Unicode.GetBytes("Wide"), 0, bytes, position, length); position += length; }
            using MemoryStream ambiguous = new MemoryStream(bytes); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(ambiguous));
            using MemoryStream bounded = new MemoryStream(Source()); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(bounded, new AccessLoadOptions { MaxMetadataBytes = 4096 }));
            using MemoryStream catalog = new MemoryStream(Source()); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(catalog, new AccessLoadOptions { MaxCatalogObjects = 1 }));
        }
        [Fact]
        public void LateMalformedRowsRemainLazyAndDoNotPreventFirstRowAccess() {
            byte[] bytes = Source(); int table = Definition(bytes, "Precise"); int changed = -1;
            for (int page = 1; page < bytes.Length / 4096; page++) {
                int offset = page * 4096; if (bytes[offset] != 1 || I32(bytes, offset + 4) != table || U16(bytes, offset + 12) < 2) continue;
                // Corrupt the second live slot, leaving the first record and all metadata intact.
                bytes[offset + 16] = 0; bytes[offset + 17] = 0; changed = page; break;
            }
            Assert.True(changed > 0); using MemoryStream source = new MemoryStream(bytes); using AccessDocument document = AccessDocument.Load(source);
            using AccessDataReader rows = document.Tables["Scalars"].OpenDataReader(); Assert.True(rows.Read()); Assert.Equal(1, rows.GetInt32(0));
            Assert.Throws<InvalidDataException>(() => rows.Read());
        }
        [Fact]
        public void LongValueCyclesRejectOnRequestedPayloadAndHonorCancellation() {
            byte[] bytes = Source(); int changed = -1;
            for (int page = 1; page < bytes.Length / 4096; page++) {
                int offset = page * 4096;
                if (bytes[offset] != 1 || Encoding.ASCII.GetString(bytes, offset + 4, 4) != "LVAL") continue;
                int rows = U16(bytes, offset + 12);
                for (int row = 0; row < rows; row++) {
                    int start = U16(bytes, offset + 14 + row * 2) & 8191; int pointer = I32(bytes, offset + start);
                    if (pointer == 0 || pointer < 0 || (pointer >> 8) >= bytes.Length / 4096) continue;
                    Write32(bytes, offset + start, page * 256 + row); changed = page; break;
                }
                if (changed > 0) break;
            }
            Assert.True(changed > 0); using MemoryStream source = new MemoryStream(bytes); using AccessDocument document = AccessDocument.Load(source);
            using AccessDataReader reader = document.Tables["Scalars"].OpenDataReader(); Assert.True(reader.Read()); Assert.Equal(1, reader["Id"]);
            Assert.Throws<InvalidDataException>(() => reader.GetString(reader.GetOrdinal("Notes")));
        }
        [Fact]
        public async Task MislabeledFileFamiliesRejectInSynchronousAndAsynchronousRoutes() {
            string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Access-Family-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try {
                string mislabeled = Path.Combine(root, "wrong.mdb"); File.WriteAllBytes(mislabeled, Source("ace12.accdb"));
                Assert.Throws<InvalidDataException>(() => AccessDocument.Inspect(mislabeled));
                Assert.Throws<InvalidDataException>(() => AccessDocument.Load(mislabeled));
                await Assert.ThrowsAsync<InvalidDataException>(() => AccessDocument.LoadAsync(mislabeled));
                string compiled = Path.Combine(root, "wrong.accde"); File.WriteAllBytes(compiled, Source("ace12.accdb"));
                Assert.Throws<NotSupportedException>(() => AccessDocument.Load(compiled));
            } finally { Directory.Delete(root, recursive: true); }
        }
    }
}
