using System.Text;
using OfficeIMO.Access;

namespace OfficeIMO.Access.Tests;

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
        Assert.True(changed); using var input = new MemoryStream(bytes);
        var error = Assert.Throws<InvalidDataException>(() => AccessDocument.Load(input)); Assert.Contains("required field", error.Message);
    }
    [Fact]
    public void UnknownPropertiesRetainBytesAndProduceNamedProfileMappingLoss() {
        byte[] bytes = Source(); int changed = 0;
        for (int i = 0; i < bytes.Length - 3; i++) if (bytes[i] == 'M' && bytes[i + 1] == 'R' && bytes[i + 2] == '2' && bytes[i + 3] == 0) { bytes[i] = (byte)'X'; changed++; }
        Assert.True(changed > 0); using var input = new MemoryStream(bytes); using var document = AccessDocument.Load(input);
        var table = document.Tables["Scalars"]; Assert.NotNull(table.NativeProperties); Assert.Equal((byte)'X', table.NativeProperties!.GetBytes()[0]);
        Assert.Contains(table.Diagnostics, x => x.Code == "access.properties.opaque");
        Assert.Contains(document.AssessSave("target.mdb").Diagnostics, x => x.Code == "access.conversion.loss.opaque-properties" && x.ObjectId == table.Id);
        using var rows = table.OpenDataReader(); Assert.True(rows.Read()); Assert.Equal(int.MaxValue, rows["Whole"]);
    }
    [Fact]
    public void CyclicAndTruncatedDefinitionChainsRejectBeforeUserDataRead() {
        byte[] cyclic = Source(); int definition = Definition(cyclic, "Precise") * 4096;
        Write32(cyclic, definition + 8, 4097); Write32(cyclic, definition + 4, definition / 4096);
        using var cycle = new MemoryStream(cyclic); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(cycle)); Assert.True(cycle.CanRead);
        byte[] truncated = Source(); definition = Definition(truncated, "Precise") * 4096;
        Write32(truncated, definition + 8, 4097); Write32(truncated, definition + 4, 0);
        using var shortChain = new MemoryStream(truncated); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(shortChain));
    }
    [Fact]
    public void AmbiguousNativeFieldNamesAndMetadataBudgetsRejectDeterministically() {
        byte[] bytes = Source(); int start = Definition(bytes, "Precise") * 4096;
        int count = U16(bytes, start + 45), indexes = I32(bytes, start + 51), position = start + 63 + indexes * 12 + count * 25;
        for (int i = 0; i < count; i++) { int length = U16(bytes, position); position += 2; if (Encoding.Unicode.GetString(bytes, position, length) == "Tiny") Array.Copy(Encoding.Unicode.GetBytes("Wide"), 0, bytes, position, length); position += length; }
        using var ambiguous = new MemoryStream(bytes); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(ambiguous));
        using var bounded = new MemoryStream(Source()); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(bounded, new AccessLoadOptions { MaxMetadataBytes = 4096 }));
        using var catalog = new MemoryStream(Source()); Assert.Throws<InvalidDataException>(() => AccessDocument.Load(catalog, new AccessLoadOptions { MaxCatalogObjects = 1 }));
    }
    [Fact]
    public void LateMalformedRowsRemainLazyAndDoNotPreventFirstRowAccess() {
        byte[] bytes = Source(); int table = Definition(bytes, "Precise"); int changed = -1;
        for (int page = 1; page < bytes.Length / 4096; page++) {
            int offset = page * 4096; if (bytes[offset] != 1 || I32(bytes, offset + 4) != table || U16(bytes, offset + 12) < 2) continue;
            // Corrupt the second live slot, leaving the first record and all metadata intact.
            bytes[offset + 16] = 0; bytes[offset + 17] = 0; changed = page; break;
        }
        Assert.True(changed > 0); using var source = new MemoryStream(bytes); using var document = AccessDocument.Load(source);
        using var rows = document.Tables["Scalars"].OpenDataReader(); Assert.True(rows.Read()); Assert.Equal(1, rows.GetInt32(0));
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
        Assert.True(changed > 0); using var source = new MemoryStream(bytes); using var document = AccessDocument.Load(source);
        using var reader = document.Tables["Scalars"].OpenDataReader(); Assert.True(reader.Read()); Assert.Equal(1, reader["Id"]);
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
