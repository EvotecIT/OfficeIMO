using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Utilities;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_OpenFileBacked_BoundsCompressedContentTypesBeforeNormalization() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                using (var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false))
                using (Stream contentTypes = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal).Open()) {
                    byte[] prefix = Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                    contentTypes.Write(prefix, 0, prefix.Length);
                    var padding = new byte[8192];
                    for (int index = 0; index < padding.Length; index++) padding[index] = (byte)' ';
                    long remaining = ExcelPackageUtilities.MaximumContentTypesEntryBytes + 1L;
                    while (remaining > 0) {
                        int count = (int)Math.Min(padding.Length, remaining);
                        contentTypes.Write(padding, 0, count);
                        remaining -= count;
                    }
                }

                Assert.True(new FileInfo(path).Length < 1_000_000L);
                InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.OpenFileBacked(path));
                Assert.Contains("content-types part", exception.Message, StringComparison.OrdinalIgnoreCase);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task Test_RemovePivotInteraction_RechecksSharedCacheInsideTransaction() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            data.CellValue(1, 1, "Region");
            data.CellValue(1, 2, "Sales");
            data.CellValue(2, 1, "East");
            data.CellValue(2, 2, 10d);
            data.AddPivotTable(
                "A1:B2",
                "D2",
                "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
            ExcelPivotInteractionInfo first = document.AddPivotSlicer(
                "SalesPivot",
                "Region",
                data.Name,
                new ExcelSlicerViewOptions { Name = "FirstFilter", Row = 6, Column = 1 });

            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            workbookLock.EnterReadLock();
            Task<bool>? removal = null;
            using var removalStarted = new ManualResetEventSlim();
            try {
                removal = Task.Factory.StartNew(
                    () => {
                        removalStarted.Set();
                        return document.RemovePivotInteraction("FirstFilter");
                    },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                Assert.True(removalStarted.Wait(TimeSpan.FromSeconds(10)));
                Assert.True(
                    SpinWait.SpinUntil(
                        () => workbookLock.WaitingWriteCount > 0 || removal.IsCompleted,
                        TimeSpan.FromSeconds(10)),
                    $"Removal task did not reach the transaction boundary (status: {removal.Status}).");
                Assert.False(removal.IsCompleted, "Removal completed before it acquired the held workbook transaction lock.");
                using (data.BeginNoLock()) {
                    document.AddPivotSlicer(
                        "SalesPivot",
                        "Region",
                        data.Name,
                        new ExcelSlicerViewOptions {
                            Name = "SecondFilter",
                            CacheName = first.CacheName,
                            Row = 6,
                            Column = 5
                        });
                }
            } finally {
                workbookLock.ExitReadLock();
            }

            Assert.True(await removal!);
            ExcelPivotInteractionInfo remaining = Assert.Single(document.GetPivotInteractions());
            Assert.Equal("SecondFilter", remaining.Name);
            Assert.Equal(first.CacheName, remaining.CacheName);
            Assert.Single(document.WorkbookPartRoot.SlicerCacheParts);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ExtendedConnectionParameters_RemapAcrossStructuralAndRangeEdits() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo query = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "ExtendedConnection",
                WorksheetName = sheet.Name,
                TableName = "ExtendedResults",
                ColumnNames = new[] { "Value" }
            });
            ConnectionsPart nativePart = document.WorkbookPartRoot.ConnectionsPart!;
            document.WorkbookPartRoot.DeletePart(nativePart);
            string xml =
                "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\">" +
                "<connection id=\"" + query.ConnectionId + "\" name=\"ExtendedConnection\" type=\"5\">" +
                "<parameters count=\"1\"><parameter name=\"Input\" cell=\"A5\"/></parameters>" +
                "</connection></connections>";
            ExcelPackagePartInfo partInfo = document.AddWorkbookConnectionMetadata(xml);
            ExtendedPart part = Assert.IsType<ExtendedPart>(document.WorkbookPartRoot.GetPartById(partInfo.RelationshipId));

            sheet.InsertRows(3, 2);
            Assert.Contains("cell=\"A7\"", ReadConnectionPartText(part), StringComparison.Ordinal);

            sheet.MoveRange("A7", "C9");
            Assert.Contains("cell=\"C9\"", ReadConnectionPartText(part), StringComparison.Ordinal);

            sheet.InsertColumns(2, 1);
            Assert.Contains("cell=\"D9\"", ReadConnectionPartText(part), StringComparison.Ordinal);

            sheet.InsertCells("D5", ExcelCellShiftDirection.Down);
            Assert.Contains("cell=\"D10\"", ReadConnectionPartText(part), StringComparison.Ordinal);

            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                WriteConnectionPartText(part, ReadConnectionPartText(part).Replace("D10", "A1"));
                throw new InvalidOperationException("Rollback probe");
            }, new ExcelMutationPlanOptions(), CancellationToken.None));
            Assert.Contains("cell=\"D10\"", ReadConnectionPartText(part), StringComparison.Ordinal);
        }

        [Fact]
        public void Test_PackageWorksheetCopy_ReusesSharedInCellImagePayload() {
            using var sourceDocument = ExcelDocument.Create(new MemoryStream());
            ExcelSheet source = sourceDocument.AddWorksheet("Source");
            source.SetInCellImage(1, 1, TinyPng, altText: "Shared A");
            source.CellValue(1, 2, "placeholder");
            Cell first = source.WorksheetPart.Worksheet!.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A1");
            Cell second = source.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            second.CellValue = (CellValue?)first.CellValue?.CloneNode(true);
            second.DataType = first.DataType?.Value;
            second.ValueMetaIndex = first.ValueMetaIndex?.Value;
            second.InlineString = null;
            source.WorksheetPart.Worksheet.Save();
            Assert.Equal(2, source.GetInCellImages().Count);

            using var targetDocument = ExcelDocument.Create(new MemoryStream());
            targetDocument.AddWorksheet("Existing");
            ExcelSheet copied = targetDocument.CopyWorksheetFrom(
                sourceDocument,
                source.Name,
                "Copied",
                ExcelSheetNameValidationMode.Sanitize,
                new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });

            Assert.Equal(2, copied.GetInCellImages().Count);
            ExtendedPart relationshipPart = targetDocument.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<ExtendedPart>()
                .Single(part => part.Parts.Any(child =>
                    child.OpenXmlPart.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)));
            Assert.Single(relationshipPart.Parts, child =>
                child.OpenXmlPart.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase));
            Assert.Empty(targetDocument.ValidateOpenXml());
        }

        private static string ReadConnectionPartText(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static void WriteConnectionPartText(OpenXmlPart part, string xml) {
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml), writable: false);
            part.FeedData(stream);
        }
    }
}
