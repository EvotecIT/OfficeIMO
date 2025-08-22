using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_LoadMissingFile_ThrowsWithPath() {
            string filePath = Path.Combine(_directoryWithFiles, "missing.xlsx");
            var ex = Assert.Throws<FileNotFoundException>(() => ExcelDocument.Load(filePath));
            Assert.Equal($"File '{filePath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public void Test_LoadNullPath_ThrowsArgumentNullException() {
            var ex = Assert.Throws<ArgumentNullException>(() => ExcelDocument.Load(null!));
            Assert.Equal("filePath", ex.ParamName);
        }

        [Fact]
        public async Task Test_LoadAsyncMissingFile_ThrowsWithPath() {
            string filePath = Path.Combine(_directoryWithFiles, "missingAsync.xlsx");
            var ex = await Assert.ThrowsAsync<FileNotFoundException>(() => ExcelDocument.LoadAsync(filePath));
            Assert.Equal($"File '{filePath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public async Task Test_LoadAsyncNullPath_ThrowsArgumentNullException() {
            var ex = await Assert.ThrowsAsync<ArgumentNullException>(() => ExcelDocument.LoadAsync(null!));
            Assert.Equal("path", ex.ParamName);
        }
    }
}
