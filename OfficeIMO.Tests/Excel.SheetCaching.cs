using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_SheetWrappersAreCachedBetweenReads() {
            string filePath = Path.Combine(_directoryWithFiles, "SheetCacheReuse.xlsx");
            using var document = ExcelDocument.Create(filePath);
            document.AddWorkSheet("One");
            document.AddWorkSheet("Two");

            var firstRead = document.Sheets;
            var secondRead = document.Sheets;

            Assert.Equal(firstRead.Count, secondRead.Count);
            for (int i = 0; i < firstRead.Count; i++) {
                Assert.Same(firstRead[i], secondRead[i]);
            }

            document.Save();
        }

        [Fact]
        public void Test_SheetCacheInvalidatesOnMutations() {
            string filePath = Path.Combine(_directoryWithFiles, "SheetCacheInvalidations.xlsx");
            using var document = ExcelDocument.Create(filePath);
            document.AddWorkSheet("Alpha");
            document.AddWorkSheet("Beta");

            var baseline = document.Sheets;
            Assert.True(baseline.Count >= 2);

            document.AddWorkSheet("Gamma");
            var afterAdd = document.Sheets;
            Assert.Contains(afterAdd, sheet => string.Equals(sheet.Name, "Gamma", StringComparison.Ordinal));

            document.RemoveWorkSheet("Alpha");
            var afterRemove = document.Sheets;
            Assert.DoesNotContain(afterRemove, sheet => string.Equals(sheet.Name, "Alpha", StringComparison.Ordinal));

            document.AddTableOfContents(placeFirst: true, withHyperlinks: false, includeNamedRanges: false, styled: false);
            var afterMove = document.Sheets;
            Assert.Equal("TOC", afterMove.First().Name);

            document.Save();
        }

        [Fact]
        public void Benchmark_CachedVersusUncachedSheetAccess() {
            string filePath = Path.Combine(_directoryWithFiles, "SheetCacheBenchmark.xlsx");
            using var document = ExcelDocument.Create(filePath);
            document.AddWorkSheet("SheetA");
            document.AddWorkSheet("SheetB");
            document.AddWorkSheet("SheetC");

            document.InvalidateSheetCache();
            ExcelSheet.ResetInstanceCountForTests();

            var initial = document.Sheets;
            int sheetCount = initial.Count;

            for (int i = 0; i < 500; i++) {
                var sheets = document.Sheets;
                Assert.Equal(sheetCount, sheets.Count);
            }

            int cachedInstances = ExcelSheet.InstancesCreatedForTests;

            document.InvalidateSheetCache();
            ExcelSheet.ResetInstanceCountForTests();

            document.SheetCachingEnabled = false;

            for (int i = 0; i < 100; i++) {
                var sheets = document.Sheets;
                Assert.Equal(sheetCount, sheets.Count);
            }

            int uncachedInstances = ExcelSheet.InstancesCreatedForTests;

            Assert.True(cachedInstances < uncachedInstances);
            Assert.True(uncachedInstances >= sheetCount * 50);

            document.Save();
        }
    }
}
