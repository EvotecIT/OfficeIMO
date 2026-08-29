using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkBoundaryTests {
    [Fact]
    public void Enforces_the_combined_decompressed_iwa_budget_across_entries() {
        byte[] first = ArchiveRecord(1, 1, new byte[48]);
        byte[] second = ArchiveRecord(2, 6000, new byte[48]);
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(first)),
            ("Index/Table.iwa", FrameIwa(second)));
        int perEntryLimit = Math.Max(first.Length, second.Length);
        var options = new IWorkReadOptions {
            MaximumDecompressedIwaBytes = perEntryLimit,
            MaximumSnappyChunkBytes = perEntryLimit,
            MaximumTotalDecompressedIwaBytes = first.Length + second.Length - 1
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options));

        Assert.Contains("source-wide limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Decodes_numbers_rows_that_use_wide_offsets() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Wide", 1, 1, 42d, wideOffsets: true)
        });

        IWorkNumbersTable table = Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers().Sheets).Tables);

        Assert.Equal(42d, Assert.IsType<double>(table.GetCell(1, 1)!.Value), 10);
    }

    [Fact]
    public void Preserves_cached_numbers_values_when_a_formula_marker_is_present() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula", 1, 1, 42d, hasFormula: true)
        });

        IWorkNumbersCell cell = Assert.Single(Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers().Sheets).Tables).Cells);

        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Equal("=?", cell.Formula);
        Assert.Equal(42d, Assert.IsType<double>(cell.Value), 10);
    }

    [Fact]
    public void Enforces_the_materialized_cell_budget_across_all_tables() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("First", 1, 1, 1d),
            new TableSpec("Second", 1, 1, 2d)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumMaterializedCells = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("source-wide limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Pre_bnc_numbers_storage_is_preserved_but_not_claimed_as_editable() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Legacy", 1, 1, 0d, legacyStorage: true)
        }, includePreview: true);

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection projection = source.ReadNumbers();
        IWorkImportReport report = projection.CreateImportReport(
            IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview);

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_LEGACY_CELL_STORAGE");
        Assert.Contains(report.UnsupportedRecords, record => record.MessageType == 6002);
    }

    [Fact]
    public void Unsupported_numbers_table_models_are_preserved_and_disable_editable_claims() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Unsupported", 1, 1, 0d, missingModel: true)
        }, includePreview: true);

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection projection = source.ReadNumbers();
        IWorkImportReport report = projection.CreateImportReport(
            IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview);

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_TABLE_MODEL_UNSUPPORTED");
        Assert.Contains(report.UnsupportedRecords, record => record.MessageType == 6000);
    }

    [Fact]
    public void Visual_only_numbers_import_bypasses_semantic_table_limits() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Large", 100, 1, 42d)
        }, includePreview: true);
        var options = new IWorkReadOptions {
            ImportMode = IWorkImportMode.VisualOnly,
            MaximumTableRows = 1
        };

        using var result = ExcelDocument.LoadNumbersWithReport(package, options);

        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.Projection.Sheets);
        Assert.Equal(0, result.ImportReport.ReconstructedItemCount);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_SEMANTIC_PROJECTION_SKIPPED");
    }

    [Fact]
    public void Counts_zip_directory_nodes_against_the_package_entry_limit() {
        using var package = new MemoryStream();
        using (var archive = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            archive.CreateEntry("Index/");
            archive.CreateEntry("Index/Subdirectory/");
            WriteEntry(archive, "Index/Document.iwa", FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>())));
        }
        package.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumEntryCount = 2 }));

        Assert.Contains("entry count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Counts_directory_bundle_nodes_against_the_package_entry_limit() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-entry-limit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(directory, "Index", "Subdirectory"));
        try {
            File.WriteAllBytes(Path.Combine(directory, "Index", "Document.iwa"),
                FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>())));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                IWorkSourceDocument.Open(directory, IWorkDocumentKind.Numbers,
                    new IWorkReadOptions { MaximumEntryCount = 2 }));

            Assert.Contains("entry count", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Splits_the_first_numbers_table_when_leading_text_would_overflow_xlsx_rows() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Full Height", 1_048_576, 1, 42d)
        }, textBox: "Leading text");
        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.Equal(2, result.Document.Sheets.Count);
        Assert.Equal("Leading text", result.Document.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(42d, result.Document.Sheets[1].CellAt(1, 1).GetValue<double>(), 10);
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        Assert.Equal(2, reopened.Sheets.Count);
    }

    [Fact]
    public void Ignores_resource_names_and_malformed_files_that_only_look_like_previews() {
        using FileStream input = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        using var source = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        using var package = new MemoryStream();
        using (var target = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries.Where(candidate =>
                         !string.Equals(candidate.FullName, "preview.jpg", StringComparison.OrdinalIgnoreCase)
                         && !string.Equals(candidate.FullName, "preview-web.jpg", StringComparison.OrdinalIgnoreCase))) {
                ZipArchiveEntry copy = target.CreateEntry(entry.FullName);
                using Stream sourceStream = entry.Open();
                using Stream targetStream = copy.Open();
                sourceStream.CopyTo(targetStream);
            }
            byte[] validPreview = ReadEntry(source, "preview.jpg");
            WriteEntry(target, "Data/unrelated-preview.jpg", validPreview);
            WriteEntry(target, "preview.jpg", new byte[] { 0xff, 0xd8, 0xff });
            WriteEntry(target, "preview-web.jpg", validPreview);
        }
        package.Position = 0;

        IWorkSourceDocument document = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.DoesNotContain(document.Previews, preview => preview.Path.StartsWith("Data/", StringComparison.Ordinal));
        Assert.DoesNotContain(document.Previews, preview => preview.Path == "preview.jpg");
        Assert.Equal("preview-web.jpg", document.PreferredRasterPreview!.Path);
    }

    private static MemoryStream CreateNumbersPackage(IReadOnlyList<TableSpec> tables, string? textBox = null,
        bool includePreview = false) {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        var records = new List<byte[]>();
        var sheetFields = new List<byte[]> { StringField(1, "Sheet") };
        var documentFields = new[] { ReferenceField(1, sheetId) };
        records.Add(ArchiveRecord(documentId, 1, Message(documentFields)));

        for (int index = 0; index < tables.Count; index++) {
            TableSpec table = tables[index];
            ulong tableInfoId = checked((ulong)(10 + index * 3));
            ulong modelId = tableInfoId + 1;
            ulong tileId = tableInfoId + 2;
            sheetFields.Add(ReferenceField(2, tableInfoId));
            records.Add(ArchiveRecord(tableInfoId, 6000,
                table.MissingModel ? Message() : Message(ReferenceField(2, modelId))));
            if (table.MissingModel) continue;

            byte[] rowInfo = table.LegacyStorage
                ? Message(VarintField(1, 0), BytesField(3, new byte[] { 1 }), BytesField(4, new byte[] { 1 }))
                : CreateBncRow(table.Value, table.WideOffsets, table.HasFormula);
            byte[] tilePayload = Message(BytesField(5, rowInfo));
            records.Add(ArchiveRecord(tileId, 6002, tilePayload));

            byte[] tileEntry = Message(VarintField(1, 0), ReferenceField(2, tileId));
            byte[] tileStorage = Message(BytesField(1, tileEntry));
            byte[] store = Message(BytesField(3, tileStorage));
            byte[] model = Message(
                BytesField(4, store),
                VarintField(6, checked((ulong)table.Rows)),
                VarintField(7, checked((ulong)table.Columns)),
                StringField(8, table.Name));
            records.Add(ArchiveRecord(modelId, 6001, model));
        }

        if (textBox != null) {
            const ulong shapeId = 1000;
            const ulong storageId = 1001;
            sheetFields.Add(ReferenceField(2, shapeId));
            records.Add(ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))));
            records.Add(ArchiveRecord(storageId, 2001, Message(StringField(3, textBox))));
        }

        records.Insert(1, ArchiveRecord(sheetId, 2, Message(sheetFields.ToArray())));
        byte[] iwaStream = Message(records.ToArray());
        return includePreview
            ? CreatePackage(
                ("Index/Document.iwa", FrameIwa(iwaStream)),
                ("preview.png", Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")))
            : CreatePackage(("Index/Document.iwa", FrameIwa(iwaStream)));
    }

    private static byte[] CreateBncRow(double value, bool wideOffsets, bool hasFormula) {
        int cellOffset = wideOffsets ? 4 : 0;
        var buffer = new byte[cellOffset + (hasFormula ? 24 : 20)];
        buffer[cellOffset] = 5;
        buffer[cellOffset + 1] = 2;
        WriteUInt32(buffer, cellOffset + 8, (1u << 1) | (hasFormula ? 1u << 9 : 0));
        Buffer.BlockCopy(BitConverter.GetBytes(value), 0, buffer, cellOffset + 12, 8);
        ushort encodedOffset = checked((ushort)(wideOffsets ? cellOffset / 4 : cellOffset));
        byte[] offsets = { (byte)encodedOffset, (byte)(encodedOffset >> 8) };
        var fields = new List<byte[]> {
            VarintField(1, 0),
            BytesField(6, buffer),
            BytesField(7, offsets)
        };
        if (wideOffsets) fields.Add(VarintField(8, 1));
        return Message(fields.ToArray());
    }

    private static byte[] ArchiveRecord(ulong identifier, uint type, byte[] payload) {
        byte[] messageInfo = Message(VarintField(1, type), VarintField(3, checked((ulong)payload.Length)));
        byte[] archiveInfo = Message(VarintField(1, identifier), BytesField(2, messageInfo));
        return Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo, payload);
    }

    private static byte[] FrameIwa(byte[] uncompressed) {
        var block = new List<byte>(uncompressed.Length + 10);
        block.AddRange(Varint(checked((ulong)uncompressed.Length)));
        int encodedLength = checked(uncompressed.Length - 1);
        if (uncompressed.Length <= 60) {
            block.Add(checked((byte)(encodedLength << 2)));
        } else {
            int lengthBytes = encodedLength <= byte.MaxValue ? 1
                : encodedLength <= ushort.MaxValue ? 2
                : encodedLength <= 0x00ff_ffff ? 3 : 4;
            block.Add(checked((byte)((59 + lengthBytes) << 2)));
            for (int index = 0; index < lengthBytes; index++) block.Add((byte)(encodedLength >> (index * 8)));
        }
        block.AddRange(uncompressed);
        int chunkLength = block.Count;
        var result = new List<byte>(chunkLength + 4) {
            0,
            (byte)chunkLength,
            (byte)(chunkLength >> 8),
            (byte)(chunkLength >> 16)
        };
        result.AddRange(block);
        return result.ToArray();
    }

    private static byte[] ReferenceField(int field, ulong identifier) =>
        BytesField(field, Message(VarintField(1, identifier)));

    private static byte[] StringField(int field, string value) =>
        BytesField(field, System.Text.Encoding.UTF8.GetBytes(value));

    private static byte[] VarintField(int field, ulong value) =>
        Message(Varint(checked((ulong)(field << 3))), Varint(value));

    private static byte[] BytesField(int field, byte[] value) =>
        Message(Varint(checked((ulong)((field << 3) | 2))), Varint(checked((ulong)value.Length)), value);

    private static byte[] Varint(ulong value) {
        var result = new List<byte>(10);
        do {
            byte next = (byte)(value & 0x7f);
            value >>= 7;
            if (value != 0) next |= 0x80;
            result.Add(next);
        } while (value != 0);
        return result.ToArray();
    }

    private static byte[] Message(params byte[][] parts) {
        int length = parts.Sum(part => part.Length);
        var result = new byte[length];
        int offset = 0;
        foreach (byte[] part in parts) {
            Buffer.BlockCopy(part, 0, result, offset, part.Length);
            offset += part.Length;
        }
        return result;
    }

    private static MemoryStream CreatePackage(params (string Path, byte[] Bytes)[] entries) {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string path, byte[] bytes) in entries) WriteEntry(archive, path, bytes);
        }
        stream.Position = 0;
        return stream;
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(path);
        using Stream target = entry.Open();
        target.Write(bytes, 0, bytes.Length);
    }

    private static byte[] ReadEntry(ZipArchive archive, string path) {
        ZipArchiveEntry entry = archive.GetEntry(path)!;
        using Stream source = entry.Open();
        using var bytes = new MemoryStream();
        source.CopyTo(bytes);
        return bytes.ToArray();
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static string Fixture(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            relativePath.Replace('/', Path.DirectorySeparatorChar));

    private sealed class TableSpec {
        internal TableSpec(string name, int rows, int columns, double value,
            bool wideOffsets = false, bool legacyStorage = false, bool hasFormula = false,
            bool missingModel = false) {
            Name = name;
            Rows = rows;
            Columns = columns;
            Value = value;
            WideOffsets = wideOffsets;
            LegacyStorage = legacyStorage;
            HasFormula = hasFormula;
            MissingModel = missingModel;
        }

        internal string Name { get; }
        internal int Rows { get; }
        internal int Columns { get; }
        internal double Value { get; }
        internal bool WideOffsets { get; }
        internal bool LegacyStorage { get; }
        internal bool HasFormula { get; }
        internal bool MissingModel { get; }
    }
}
