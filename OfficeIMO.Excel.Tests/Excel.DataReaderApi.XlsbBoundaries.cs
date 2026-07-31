using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Read;
using System.Data.Common;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_XlsbRejectsMissingWorkbookEndBoundary() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingEndBook.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            RemoveFinalXlsbRecord(path, "xl/workbook.bin", expectedRecordType: 132);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("BrtBeginBook/BrtEndBook", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsTruncatedSharedStringHeader() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedSstHeader.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            RemoveFirstXlsbRecordPayload(path, "xl/sharedStrings.bin", recordType: 159);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("BrtBeginSst", exception.Message, StringComparison.Ordinal);
            Assert.Contains("truncated", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsMissingWorksheetRelationship() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingWorksheetRelationship.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            DuplicateFirstXlsbBundleSheet(path, mutateRelationshipId: true);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("missing relationship", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsCaseInsensitiveDuplicateWorksheetNames() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateWorksheetName.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            DuplicateFirstXlsbBundleSheet(path, changeWorksheetNameCase: true);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("duplicate worksheet name", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsTruncatedCellFormatRecord() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedCellFormat.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("styles-dates-formulas.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayloadAfter(
                path,
                "xl/styles.bin",
                collectionBeginType: 617,
                recordType: 47,
                data => data.Take(4).ToArray());

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("BrtXf", exception.Message, StringComparison.Ordinal);
            Assert.Contains("payload length", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsDuplicateCustomNumberFormat() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateNumberFormat.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("styles-dates-formulas.xlsb"), path);
        try {
            DuplicateFirstXlsbRecord(path, "xl/styles.bin", recordType: 44);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("duplicate custom number format", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsMismatchedCustomNumberFormatCount() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidNumberFormatCount.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("styles-dates-formulas.xlsb"), path);
        try {
            IncrementXlsbStyleCollectionDeclaredCount(path, 615);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("number-format collection", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("declares", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("contains", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsExplicitlyEmptyCellFormatCollection() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.EmptyCellFormats.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            EmptyXlsbStyleCollection(
                path,
                beginRecordType: 617,
                itemRecordType: 47,
                endRecordType: 618);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("non-empty cell-format collection", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(611, 43, 612)]
    [InlineData(603, 45, 604)]
    [InlineData(613, 46, 614)]
    [InlineData(626, 47, 627)]
    public void OpenDataReader_XlsbRejectsEmptyMandatoryStyleCollection(
        int beginRecordType,
        int itemRecordType,
        int endRecordType) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.EmptyMandatoryStyleCollection.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            EmptyXlsbStyleCollection(path, beginRecordType, itemRecordType, endRecordType);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("required formatting collections", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(626, 0, 0, "parent cell-style reference")]
    [InlineData(617, 0, ushort.MaxValue, "parent cell-style reference")]
    [InlineData(617, 4, ushort.MaxValue, "font, fill, or border reference")]
    [InlineData(617, 6, ushort.MaxValue, "font, fill, or border reference")]
    [InlineData(617, 8, ushort.MaxValue, "font, fill, or border reference")]
    public void OpenDataReader_XlsbRejectsInvalidCellFormatReference(
        int collectionBeginType,
        int fieldOffset,
        int value,
        string expectedMessage) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidCellFormatReference.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayloadAfter(
                path,
                "xl/styles.bin",
                collectionBeginType,
                recordType: 47,
                data => {
                    byte[] replacement = data.ToArray();
                    replacement[fieldOffset] = (byte)value;
                    replacement[fieldOffset + 1] = (byte)(value >> 8);
                    return replacement;
                });

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains(expectedMessage, exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsCellOutsideRowHeaderSpans() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.CellOutsideRowSpan.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayload(
                path,
                "xl/worksheets/sheet1.bin",
                recordType: 0,
                data => {
                    var replacement = new byte[25];
                    Buffer.BlockCopy(data, 0, replacement, 0, 13);
                    WriteUInt32LittleEndian(replacement, 13, 1U);
                    WriteUInt32LittleEndian(replacement, 17, 100U);
                    WriteUInt32LittleEndian(replacement, 21, 100U);
                    return replacement;
                });

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("not covered", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsTruncatedRowHeaderSpans() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedRowSpan.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayload(
                path,
                "xl/worksheets/sheet1.bin",
                recordType: 0,
                data => {
                    var replacement = new byte[17];
                    Buffer.BlockCopy(data, 0, replacement, 0, 13);
                    WriteUInt32LittleEndian(replacement, 13, 1U);
                    return replacement;
                });

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("column-span payload", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsDuplicateWorksheetDimension() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateDimension.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            DuplicateFirstXlsbRecord(
                path,
                "xl/worksheets/sheet1.bin",
                recordType: 148);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("more than one BrtWsDim", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsMissingWorksheetDimension() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingDimension.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            RemoveFirstXlsbRecord(
                path,
                "xl/worksheets/sheet1.bin",
                recordType: 148);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("required BrtWsDim", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsRowHeaderOutsideSheetData() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.RowOutsideSheetData.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            InsertFirstXlsbRecordBefore(
                path,
                "xl/worksheets/sheet1.bin",
                sourceRecordType: 0,
                targetRecordType: 145);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("row or cell record", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("outside", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(43, "BrtFont")]
    [InlineData(45, "BrtFill")]
    [InlineData(46, "BrtBorder")]
    public void OpenDataReader_XlsbRejectsTruncatedMandatoryStylePayload(
        int recordType,
        string expectedRecordName) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedMandatoryStyle.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayload(
                path,
                "xl/styles.bin",
                recordType,
                data => data.Take(1).ToArray());

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains(expectedRecordName, exception.Message, StringComparison.Ordinal);
            Assert.Contains("truncated", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(603, 604, "fill")]
    [InlineData(611, 612, "font")]
    [InlineData(613, 614, "border")]
    [InlineData(615, 616, "number-format")]
    [InlineData(617, 618, "cell-format")]
    [InlineData(626, 627, "cell-style-format")]
    public void OpenDataReader_XlsbRejectsDuplicateStyleCollection(
        int beginRecordType,
        int endRecordType,
        string expectedCollectionName) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateStyleCollection.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("styles-dates-formulas.xlsb"), path);
        try {
            DuplicateFirstXlsbCollection(
                path,
                "xl/styles.bin",
                beginRecordType,
                endRecordType);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("more than one", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(expectedCollectionName, exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsExternalWorksheetRelationship() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.ExternalWorksheetRelationship.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            MakeFirstWorksheetRelationshipExternal(path);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("external relationship", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbNextResultRejectsTruncatedCellPayloadDuringDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedSkippedSheetCell.{Guid.NewGuid():N}.xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Ready");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Never delivered");
                File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
            }
            TruncateFirstXlsbCellRecordPayload(
                path,
                "xl/worksheets/sheet2.bin");

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => reader.NextResult());

            Assert.Contains("cell record", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("truncated", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbNextResultRejectsMissingCellStyleDuringDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidSkippedSheetStyle.{Guid.NewGuid():N}.xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Ready");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Never delivered");
                File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
            }
            ReplaceFirstXlsbDataCellStyleIndex(
                path,
                0x00FFFFFEU,
                "xl/worksheets/sheet2.bin");

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => reader.NextResult());

            Assert.Contains("missing cell format", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbNextResultRejectsFormulaTokensDuringDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.UnsupportedSkippedSheetFormula.{Guid.NewGuid():N}.xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Ready");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellFormula(2, 1, "1+1");
                File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { UseCachedFormulaResult = false });
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));

            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => reader.NextResult());

            Assert.Contains("formula-token", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbNextResultRejectsDuplicateRowHeaderDuringDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateSkippedSheetRow.{Guid.NewGuid():N}.xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Ready");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Never delivered");
                File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
            }
            DuplicateFirstXlsbRecord(
                path,
                "xl/worksheets/sheet2.bin",
                recordType: 0);

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => reader.NextResult());

            Assert.Contains("non-increasing row index", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    private static void RemoveFinalXlsbRecord(
        string path,
        string entryName,
        int expectedRecordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        XlsbRecordSlice final = default;
        bool found = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            final = record;
            found = true;
        }

        Assert.True(found);
        Assert.Equal(expectedRecordType, final.Type);
        Assert.Equal(bytes.Length, reader.Position);
        byte[] truncated = new byte[final.RecordOffset];
        Buffer.BlockCopy(bytes, 0, truncated, 0, truncated.Length);
        ReplaceZipEntry(path, entryName, truncated);
    }

    private static void RemoveFirstXlsbRecordPayload(
        string path,
        string entryName,
        int recordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        XlsbRecordSlice target = default;
        bool found = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            if (record.Type == recordType) {
                target = record;
                found = true;
                break;
            }
        }

        Assert.True(found);
        Assert.InRange(target.Size, 1, 127);
        Assert.Equal(target.Size, bytes[target.PayloadOffset - 1]);
        byte[] mutated = new byte[bytes.Length - target.Size];
        Buffer.BlockCopy(bytes, 0, mutated, 0, target.PayloadOffset);
        mutated[target.PayloadOffset - 1] = 0;
        Buffer.BlockCopy(
            bytes,
            target.PayloadOffset + target.Size,
            mutated,
            target.PayloadOffset,
            bytes.Length - target.PayloadOffset - target.Size);
        ReplaceZipEntry(path, entryName, mutated);
    }

    private static void RemoveFirstXlsbRecord(
        string path,
        string entryName,
        int recordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        using var output = new MemoryStream();
        bool removed = false;
        foreach (XlsbRecord record in records) {
            if (!removed && record.Type == recordType) {
                removed = true;
                continue;
            }

            XlsbRecordWriter.Write(output, record.Type, record.Data);
        }

        Assert.True(removed);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void EmptyXlsbStyleCollection(
        string path,
        int beginRecordType,
        int itemRecordType,
        int endRecordType) {
        const string entryName = "xl/styles.bin";
        byte[] bytes = ReadZipEntry(path, entryName);
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        using var output = new MemoryStream(bytes.Length);
        bool inCollection = false;
        bool foundBegin = false;
        bool foundItem = false;
        bool foundEnd = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            int headerLength = record.PayloadOffset - record.RecordOffset;
            if (record.Type == beginRecordType) {
                Assert.False(inCollection);
                Assert.Equal(sizeof(uint), record.Size);
                output.Write(bytes, record.RecordOffset, headerLength);
                output.Write(new byte[sizeof(uint)], 0, sizeof(uint));
                inCollection = true;
                foundBegin = true;
                continue;
            }

            if (inCollection && record.Type == itemRecordType) {
                foundItem = true;
                continue;
            }

            int recordLength = checked(headerLength + record.Size);
            output.Write(bytes, record.RecordOffset, recordLength);
            if (record.Type == endRecordType) {
                Assert.True(inCollection);
                inCollection = false;
                foundEnd = true;
            }
        }

        Assert.True(foundBegin);
        Assert.True(foundItem);
        Assert.True(foundEnd);
        Assert.False(inCollection);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void DuplicateFirstXlsbBundleSheet(
        string path,
        bool mutateRelationshipId = false,
        bool changeWorksheetNameCase = false) {
        byte[] bytes = ReadZipEntry(path, "xl/workbook.bin");
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        XlsbRecordSlice bundle = default;
        XlsbRecordSlice endBook = default;
        bool foundBundle = false;
        bool foundEndBook = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            if (!foundBundle && record.Type == 156) {
                bundle = record;
                foundBundle = true;
            }
            if (record.Type == 132) {
                endBook = record;
                foundEndBook = true;
            }
        }

        Assert.True(foundBundle);
        Assert.True(foundEndBook);
        int bundleLength = bundle.PayloadOffset - bundle.RecordOffset + bundle.Size;
        var duplicate = new byte[bundleLength];
        Buffer.BlockCopy(
            bytes,
            bundle.RecordOffset,
            duplicate,
            0,
            bundleLength);
        int payloadOffset = bundle.PayloadOffset - bundle.RecordOffset;
        int relationshipLength = checked((int)ReadUInt32LittleEndian(
            duplicate,
            payloadOffset + 8));
        int relationshipOffset = payloadOffset + 12;
        int nameLengthOffset = checked(
            relationshipOffset + relationshipLength * 2);
        int nameLength = checked((int)ReadUInt32LittleEndian(
            duplicate,
            nameLengthOffset));
        int nameOffset = nameLengthOffset + 4;
        if (mutateRelationshipId) {
            Assert.True(relationshipLength > 0);
            duplicate[relationshipOffset] =
                duplicate[relationshipOffset] == (byte)'z'
                    ? (byte)'y'
                    : (byte)'z';
        }
        if (changeWorksheetNameCase) {
            bool changed = false;
            for (int index = 0; index < nameLength; index++) {
                int characterOffset = nameOffset + index * 2;
                char value = (char)(
                    duplicate[characterOffset]
                    | duplicate[characterOffset + 1] << 8);
                if (!char.IsLetter(value)) {
                    continue;
                }

                char replacement = char.IsUpper(value)
                    ? char.ToLowerInvariant(value)
                    : char.ToUpperInvariant(value);
                duplicate[characterOffset] = (byte)replacement;
                duplicate[characterOffset + 1] = (byte)(replacement >> 8);
                changed = true;
                break;
            }
            Assert.True(changed);
        }

        var expanded = new byte[bytes.Length + duplicate.Length];
        Buffer.BlockCopy(
            bytes,
            0,
            expanded,
            0,
            endBook.RecordOffset);
        Buffer.BlockCopy(
            duplicate,
            0,
            expanded,
            endBook.RecordOffset,
            duplicate.Length);
        Buffer.BlockCopy(
            bytes,
            endBook.RecordOffset,
            expanded,
            endBook.RecordOffset + duplicate.Length,
            bytes.Length - endBook.RecordOffset);
        ReplaceZipEntry(path, "xl/workbook.bin", expanded);
    }

    private static void ReplaceFirstXlsbRecordPayload(
        string path,
        string entryName,
        int recordType,
        Func<byte[], byte[]> replace) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        using var output = new MemoryStream();
        bool replaced = false;
        foreach (XlsbRecord record in records) {
            byte[] data = !replaced && record.Type == recordType
                ? replace(record.Data)
                : record.Data;
            if (!replaced && record.Type == recordType) {
                replaced = true;
            }
            XlsbRecordWriter.Write(output, record.Type, data);
        }

        Assert.True(replaced);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void ReplaceFirstXlsbRecordPayloadAfter(
        string path,
        string entryName,
        int collectionBeginType,
        int recordType,
        Func<byte[], byte[]> replace) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        using var output = new MemoryStream();
        bool inCollection = false;
        bool replaced = false;
        foreach (XlsbRecord record in records) {
            if (record.Type == collectionBeginType) {
                inCollection = true;
            }
            byte[] data = !replaced && inCollection && record.Type == recordType
                ? replace(record.Data)
                : record.Data;
            if (!replaced && inCollection && record.Type == recordType) {
                replaced = true;
            }
            XlsbRecordWriter.Write(output, record.Type, data);
        }

        Assert.True(replaced);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void DuplicateFirstXlsbRecord(
        string path,
        string entryName,
        int recordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        using var output = new MemoryStream();
        bool duplicated = false;
        foreach (XlsbRecord record in records) {
            XlsbRecordWriter.Write(output, record.Type, record.Data);
            if (!duplicated && record.Type == recordType) {
                XlsbRecordWriter.Write(output, record.Type, record.Data);
                duplicated = true;
            }
        }

        Assert.True(duplicated);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void InsertFirstXlsbRecordBefore(
        string path,
        string entryName,
        int sourceRecordType,
        int targetRecordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        XlsbRecord source = Assert.Single(
            records.Where(record => record.Type == sourceRecordType).Take(1));
        using var output = new MemoryStream();
        bool inserted = false;
        foreach (XlsbRecord record in records) {
            if (!inserted && record.Type == targetRecordType) {
                XlsbRecordWriter.Write(output, source.Type, source.Data);
                inserted = true;
            }

            XlsbRecordWriter.Write(output, record.Type, record.Data);
        }

        Assert.True(inserted);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void DuplicateFirstXlsbCollection(
        string path,
        string entryName,
        int beginRecordType,
        int endRecordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        int beginIndex = records.ToList().FindIndex(record => record.Type == beginRecordType);
        Assert.True(beginIndex >= 0);
        int endIndex = records.ToList().FindIndex(
            beginIndex,
            record => record.Type == endRecordType);
        Assert.True(endIndex >= beginIndex);

        using var output = new MemoryStream();
        for (int index = 0; index < records.Count; index++) {
            XlsbRecord record = records[index];
            XlsbRecordWriter.Write(output, record.Type, record.Data);
            if (index != endIndex) {
                continue;
            }

            for (int duplicateIndex = beginIndex; duplicateIndex <= endIndex; duplicateIndex++) {
                XlsbRecord duplicate = records[duplicateIndex];
                XlsbRecordWriter.Write(output, duplicate.Type, duplicate.Data);
            }
        }

        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void MakeFirstWorksheetRelationshipExternal(string path) {
        const string entryName = "xl/_rels/workbook.bin.rels";
        byte[] bytes = ReadZipEntry(path, entryName);
        XDocument document;
        using (var input = new MemoryStream(bytes, writable: false)) {
            document = XDocument.Load(input);
        }

        XNamespace relationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement relationship = document
            .Descendants(relationshipsNamespace + "Relationship")
            .First(element => ((string?)element.Attribute("Type"))
                ?.EndsWith("/worksheet", StringComparison.Ordinal) == true);
        relationship.SetAttributeValue("Target", "https://example.invalid/sheet1.bin");
        relationship.SetAttributeValue("TargetMode", "External");

        using var output = new MemoryStream();
        document.Save(output);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void TruncateFirstXlsbCellRecordPayload(
        string path,
        string entryName) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        using var output = new MemoryStream();
        bool truncated = false;
        foreach (XlsbRecord record in records) {
            bool isCell = record.Type is >= 1 and <= 11 or 62;
            byte[] data = !truncated && isCell
                ? record.Data.Take(sizeof(int)).ToArray()
                : record.Data;
            XlsbRecordWriter.Write(output, record.Type, data);
            truncated |= isCell;
        }

        Assert.True(truncated);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static byte[] ReadZipEntry(string path, string entryName) {
        using var archive = ZipFile.OpenRead(path);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"The XLSB fixture has no '{entryName}' part.");
        using Stream input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static void ReplaceZipEntry(
        string path,
        string entryName,
        byte[] bytes) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry original = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"The XLSB fixture has no '{entryName}' part.");
        original.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            entryName,
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, bytes.Length);
    }

    private static uint ReadUInt32LittleEndian(byte[] bytes, int offset) =>
        (uint)(
            bytes[offset]
            | bytes[offset + 1] << 8
            | bytes[offset + 2] << 16
            | bytes[offset + 3] << 24);

}
