using OfficeIMO.Access;

namespace OfficeIMO.Access.Tests;

public sealed class AccessCreationTests {
    private static byte[] Save(AccessDocument document, AccessSaveOptions? options = null) { using var output = new MemoryStream(); document.Save(output, options); Assert.True(output.CanWrite); return output.ToArray(); }

    [Theory]
    [InlineData(AccessFileFormat.Mdb)]
    [InlineData(AccessFileFormat.Accdb)]
    public void NativeCreationPreservesSchemaValuesAndModelOmission(AccessFileFormat format) {
        using var document = AccessDocument.Create(new AccessCreateOptions { Format = format, DatabaseTitle = "Synthetic Ł🙂 database" });
        var parent = document.Tables.Add("Parent Table"); parent.Columns.Add("Id", AccessDataType.Int32); parent.Indexes.AddPrimaryKey("PK_Parent", "Id"); parent.AppendRow(new AccessRowValues { ["Id"] = 3 });
        var table = document.Tables.Add("Items_01"); table.Columns.AddAutoNumber("Id", 101); table.Columns.Add("ParentId", AccessDataType.Int32); table.Columns.Add("Unicode", AccessDataType.ShortText); table.Columns.Add("Memo", AccessDataType.LongText); table.Columns.Add("Blob", AccessDataType.Binary); table.Columns.AddDecimal("Number", 28, 9); table.Columns.Add("Flag", AccessDataType.Boolean); table.Columns.Add("Token", AccessDataType.Guid); table.Indexes.AddPrimaryKey("PK_Items", "Id");
        document.Relationships.Add("ItemParents", parent.Columns["Id"], table.Columns["ParentId"]);
        byte[] binary = Enumerable.Range(0, 20000).Select(i => (byte)(i % 251)).ToArray(); string memo = string.Concat(Enumerable.Repeat("Ł🙂 漢字\r\n", 2000)); Guid guid = Guid.NewGuid();
        table.AppendRow(new AccessRowValues { ["ParentId"] = 3, ["Unicode"] = "Zażółć Ł🙂", ["Memo"] = memo, ["Blob"] = binary, ["Number"] = -1234567890123456789.123456789m, ["Token"] = guid });
        byte[] bytes = Save(document);
        using (var model = table.OpenDataReader()) { Assert.True(model.Read()); Assert.False(model.IsSpecified(0)); Assert.False(model.IsSpecified(6)); }
        using var source = AccessDocument.Load(new MemoryStream(bytes));
        Assert.Equal(format, source.Format); Assert.Equal(2, source.Tables.Count); Assert.Single(source.Relationships);
        Assert.Equal("Synthetic Ł🙂 database", source.Properties["AppTitle"]); Assert.Equal(1252, source.CodePage); Assert.Equal(1033, source.SortOrder);
        var loaded = source.Tables["Items_01"]; Assert.True(loaded.Columns["Id"].IsAutoNumber); Assert.True(loaded.Indexes["PK_Items"].IsPrimaryKey); Assert.Equal(9, loaded.Columns["Number"].Scale);
        using var rows = loaded.OpenDataReader(); Assert.True(rows.Read()); Assert.Equal(101, rows["Id"]); Assert.Equal("Zażółć Ł🙂", rows["Unicode"]); Assert.Equal(memo, rows["Memo"]); Assert.Equal(binary, (byte[])rows["Blob"]); Assert.Equal(-1234567890123456789.123456789m, rows["Number"]); Assert.Equal(false, rows["Flag"]); Assert.Equal(guid, rows["Token"]); Assert.False(rows.Read());
        Assert.Equal(bytes, Save(document));
    }

    [Theory]
    [InlineData(AccessFileFormat.Mdb)]
    [InlineData(AccessFileFormat.Accdb)]
    public void AllocationDefinitionChainsAndLongValueBoundariesRemainReadable(AccessFileFormat format) {
        using var document = AccessDocument.Create(new AccessCreateOptions { Format = format });
        var wide = document.Tables.Add("Wide"); var wideRow = new AccessRowValues();
        for (int i = 0; i < 255; i++) { string name = "Field" + i.ToString("D3"); wide.Columns.Add(name, AccessDataType.Byte); wideRow[name] = (byte)i; }
        wide.AppendRow(wideRow);
        var values = document.Tables.Add("Lengths"); values.Columns.Add("Id", AccessDataType.Int32); values.Columns.Add("Payload", AccessDataType.Binary); values.Indexes.AddPrimaryKey("PK_Lengths", "Id");
        foreach (int length in new[] { 0, 1, 64, 65, 4072, 4076, 4077, 8144, 8145 }) values.AppendRow(new AccessRowValues { ["Id"] = length, ["Payload"] = Enumerable.Range(0, length).Select(i => (byte)(i % 251)).ToArray() });
        var allocated = document.Tables.Add("Allocated"); allocated.Columns.Add("Id", AccessDataType.AutoNumber); allocated.Columns.Add("Padding", AccessDataType.ShortText, 255); allocated.Indexes.AddPrimaryKey("PK_Allocated", "Id");
        for (int i = 0; i < 4200; i++) allocated.AppendRow(new AccessRowValues { ["Padding"] = new string('a', 255) });
        byte[] bytes = Save(document); Assert.True(bytes.Length > 512 * 4096);
        using var source = AccessDocument.Load(new MemoryStream(bytes)); Assert.Equal(255, source.Tables["Wide"].Columns.Count); Assert.Equal(4200, source.Tables["Allocated"].RowCount);
        using (var row = source.Tables["Wide"].OpenDataReader()) { Assert.True(row.Read()); Assert.Equal((byte)254, row["Field254"]); }
        using (var rows = source.Tables["Lengths"].OpenDataReader()) while (rows.Read()) { int length = (int)rows["Id"]; Assert.Equal(Enumerable.Range(0, length).Select(i => (byte)(i % 251)), (byte[])rows["Payload"]); }
        using var all = source.Tables["Allocated"].OpenDataReader(); int count = 0; while (all.Read()) { count++; Assert.Equal(count, all["Id"]); Assert.Equal(255, ((string)all["Padding"]).Length); } Assert.Equal(4200, count);
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("foreign")]
    [InlineData("unicode-key")]
    [InlineData("null-boolean")]
    [InlineData("decimal-scale")]
    [InlineData("date-precision")]
    [InlineData("row-limit")]
    [InlineData("empty-index-type")]
    [InlineData("unicode-name")]
    [InlineData("unicode-value")]
    [InlineData("text-collation-duplicate")]
    public void UnsupportedOrLossyModelsFailBeforeDestinationChanges(string scenario) {
        using var document = AccessDocument.Create(); var table = document.Tables.Add("Items"); table.Columns.Add("Id", AccessDataType.Int32); table.Indexes.AddPrimaryKey("PK_Items", "Id");
        switch (scenario) {
            case "duplicate": table.AppendRow(new AccessRowValues { ["Id"] = 1 }); table.AppendRow(new AccessRowValues { ["Id"] = 1 }); break;
            case "foreign":
                var parent = document.Tables.Add("Parents"); parent.Columns.Add("Id", AccessDataType.Int32); parent.Indexes.AddPrimaryKey("PK_Parents", "Id"); document.Relationships.Add("ParentItems", parent.Columns["Id"], table.Columns["Id"]); table.AppendRow(new AccessRowValues { ["Id"] = 1 }); break;
            case "unicode-key": table.Columns.Add("Name", AccessDataType.ShortText); table.Indexes.AddUnique("UniqueName", "Name"); table.AppendRow(new AccessRowValues { ["Id"] = 1, ["Name"] = "Ł🙂" }); break;
            case "null-boolean": table.Columns.Add("Flag", AccessDataType.Boolean); table.AppendRow(new AccessRowValues { ["Id"] = 1, ["Flag"] = null }); break;
            case "decimal-scale": table.Columns.AddDecimal("Number", 6, 2); table.AppendRow(new AccessRowValues { ["Id"] = 1, ["Number"] = 1.234m }); break;
            case "date-precision": table.Columns.Add("Date", AccessDataType.DateTime); table.AppendRow(new AccessRowValues { ["Id"] = 1, ["Date"] = new DateTime(2026, 1, 1).AddTicks(1) }); break;
            case "row-limit":
                var row = new AccessRowValues { ["Id"] = 1 }; for (int i = 0; i < 8; i++) { string name = "Text" + i; table.Columns.Add(name, AccessDataType.ShortText, 255); row[name] = new string('a', 255); } table.AppendRow(row); break;
            case "empty-index-type": table.Columns.Add("Wide", AccessDataType.Double); table.Indexes.Add("WideIndex", "Wide"); break;
            case "unicode-name": table.Columns.Add("Bad\ud800", AccessDataType.Byte); break;
            case "unicode-value": table.Columns.Add("Text", AccessDataType.ShortText); table.AppendRow(new AccessRowValues { ["Id"] = 1, ["Text"] = "Bad\ud800" }); break;
            case "text-collation-duplicate":
                table.Columns.Add("Text", AccessDataType.ShortText); table.Indexes.AddUnique("UniqueText", "Text");
                table.AppendRow(new AccessRowValues { ["Id"] = 1, ["Text"] = "a b " }); table.AppendRow(new AccessRowValues { ["Id"] = 2, ["Text"] = "A B" }); break;
        }
        var options = new AccessSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
        Assert.Equal(AccessOperationStatus.Unsupported, document.AssessSave(options).Status);
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true); destination.Position = 1;
        Assert.Throws<AccessOperationNotSupportedException>(() => document.Save(destination, options)); Assert.Equal(1, destination.Position); Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
        string root = Path.Combine(Path.GetTempPath(), "AccessCreationRejected-" + Guid.NewGuid().ToString("N"));
        Assert.Throws<AccessOperationNotSupportedException>(() => document.Save(Path.Combine(root, "out.accdb"), options)); Assert.False(Directory.Exists(root));
    }

    [Fact]
    public void AssessmentCacheInvalidationBudgetsAndCancellationProtectOutput() {
        using var document = AccessDocument.Create(); var table = document.Tables.Add("Items"); table.Columns.Add("Id", AccessDataType.AutoNumber); table.Indexes.AddPrimaryKey("PK_Items", "Id");
        table.AppendRow(new AccessRowValues()); document.AssessSave().RequireNoLoss();
        using (var update = document.BeginUpdate()) { table.AppendRow(new AccessRowValues()); }
        table.AppendRow(new AccessRowValues());
        using var source = AccessDocument.Load(new MemoryStream(Save(document))); Assert.Equal(2, source.Tables["Items"].RowCount);
        Assert.Equal(AccessOperationStatus.Unsupported, document.AssessSave(new AccessSaveOptions { MaxOutputBytes = 4096 }).Status);
        using var assessmentCancellation = new CancellationTokenSource(); document.AssessSave(cancellationToken: assessmentCancellation.Token).RequireNoLoss(); assessmentCancellation.Cancel();
        Assert.NotEmpty(Save(document)); // An old assessment token cannot cancel a later save.
        using var output = new MemoryStream(); output.WriteByte(42); output.Position = 0;
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel(); Assert.ThrowsAny<OperationCanceledException>(() => document.Save(output, cancellationToken: cancelled.Token)); Assert.Equal(new byte[] { 42 }, output.ToArray());
    }
}
