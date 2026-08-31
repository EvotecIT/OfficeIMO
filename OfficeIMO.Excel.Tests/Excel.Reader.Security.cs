using System.IO.Compression;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_SharedStringMetadata_DoesNotAllocateDeclaredCapacity() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderSharedStringMetadataLimit.xlsx");

            try {
                CreateSharedStringWorkbook(filePath, "<si><t>safe</t></si>", "2147483647", "2147483647");

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions {
                    MaxSharedStringItems = 4
                });
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:A1");

                Assert.Equal("safe", values[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_SharedStringCache_RejectsTooManyItems() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderSharedStringTooManyItems.xlsx");

            try {
                CreateSharedStringWorkbook(filePath, "<si><t>one</t></si><si><t>two</t></si><si><t>three</t></si>", "3", "3");

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions {
                    MaxSharedStringItems = 2
                });

                Assert.Throws<InvalidDataException>(() => reader.GetSheet("Data").ReadRange("A1:A1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_SharedStringCache_RejectsOversizedItemText() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderSharedStringOversizedItem.xlsx");

            try {
                CreateSharedStringWorkbook(filePath, "<si><r><t>ab</t></r><r><t>cd</t></r></si>", "1", "1");

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions {
                    MaxSharedStringItemCharacters = 3
                });

                Assert.Throws<InvalidDataException>(() => reader.GetSheet("Data").ReadRange("A1:A1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_SharedStringCache_RejectsAggregateTextBudgetOverflow() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderSharedStringAggregateLimit.xlsx");

            try {
                CreateSharedStringWorkbook(filePath, "<si><t>one</t></si><si><t>two</t></si>", "2", "2");

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions {
                    MaxSharedStringCharacters = 5
                });

                Assert.Throws<InvalidDataException>(() => reader.GetSheet("Data").ReadRange("A1:A1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_SharedStringUtf8Cache_BoundsEntriesAndLargeValues() {
            var xml = new StringBuilder(
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            for (int index = 0; index <= SharedStringCache.Utf8CacheSlotCount; index++) {
                xml.Append("<si><t>value-").Append(index).Append("</t></si>");
            }
            string largeValue = new string('x', SharedStringCache.MaximumCachedUtf8ItemBytes + 1);
            xml.Append("<si><t>").Append(largeValue).Append("</t></si></sst>");
            byte[] payload = Encoding.UTF8.GetBytes(xml.ToString());
            var cache = SharedStringCache.Build(
                () => new MemoryStream(payload, writable: false),
                new ExcelReadOptions {
                    MaxSharedStringItems = SharedStringCache.Utf8CacheSlotCount + 2,
                    MaxSharedStringItemCharacters = largeValue.Length
                });

            Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> first));
            Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> repeated));
            Assert.Same(first.Array, repeated.Array);

            Assert.True(cache.TryGetUtf8(
                SharedStringCache.Utf8CacheSlotCount,
                out ArraySegment<byte> collision));
            Assert.NotSame(first.Array, collision.Array);
            Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> afterEviction));
            Assert.NotSame(first.Array, afterEviction.Array);

            int largeIndex = SharedStringCache.Utf8CacheSlotCount + 1;
            Assert.True(cache.TryGetUtf8(largeIndex, out ArraySegment<byte> largeFirst));
            Assert.True(cache.TryGetUtf8(largeIndex, out ArraySegment<byte> largeSecond));
            Assert.NotSame(largeFirst.Array, largeSecond.Array);
            Assert.Equal(largeValue, Encoding.UTF8.GetString(largeSecond.Array!, largeSecond.Offset, largeSecond.Count));
        }

        private static void CreateSharedStringWorkbook(string filePath, string sharedStringItemsXml, string count, string uniqueCount) {
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "placeholder");
                document.Save();
            }

            ReplacePackageEntry(filePath, "xl/sharedStrings.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + count + "\" uniqueCount=\"" + uniqueCount + "\">" +
                sharedStringItemsXml +
                "</sst>");
        }

        private static void ReplacePackageEntry(string filePath, string entryName, string content) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry? existing = archive.GetEntry(entryName);
            existing?.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryName);
            using var writer = new StreamWriter(replacement.Open(), Encoding.UTF8);
            writer.Write(content);
        }
    }
}
