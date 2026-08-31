using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void GetSheetNames_FallsBackToSdkForUnexpectedWorkbookContentType() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "[Content_Types].xml";
            string contentTypes = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string changed = contentTypes.Replace(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "text/plain");
            Assert.NotEqual(contentTypes, changed);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(changed));

            Assert.Throws<XlsxTabularFastPathNotSupportedException>(() =>
                XlsxTabularWorkbook.ReadSheetNames(path, new ExcelReadOptions()));
            Assert.Equal(new[] { "Data" }, ExcelDocument.GetSheetNames(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_SdkPathEnforcesWorksheetDefinitionLimit() {
        string xlsxPath = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MetadataLimits.{Guid.NewGuid():N}.xlsx");
        string xlsmPath = Path.ChangeExtension(xlsxPath, ".xlsm");
        try {
            using (var document = ExcelDocument.Create(xlsxPath)) {
                document.AddWorksheet("First");
                document.AddWorksheet("Second");
                document.Save();
            }
            File.Move(xlsxPath, xlsmPath);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    xlsmPath,
                    new ExcelReadOptions { MaxWorksheets = 1 }));
            Assert.Contains("1 worksheet definitions", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(xlsxPath);
            File.Delete(xlsmPath);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_SdkFallbackEnforcesMetadataPartLimit(bool useBytes) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "[Content_Types].xml";
            XDocument contentTypes = XDocument.Parse(
                Encoding.UTF8.GetString(ReadZipEntry(path, entryName)));
            XNamespace contentTypeNamespace =
                "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement workbookOverride = contentTypes.Root!
                .Elements(contentTypeNamespace + "Override")
                .Single(element => string.Equals(
                    (string?)element.Attribute("PartName"),
                    "/xl/workbook.xml",
                    StringComparison.Ordinal));
            workbookOverride.SetAttributeValue("ContentType", "text/plain");
            ReplaceZipEntry(
                path,
                entryName,
                Encoding.UTF8.GetBytes(contentTypes.ToString(SaveOptions.DisableFormatting)));

            var options = new ExcelReadOptions { MaxMetadataPartBytes = 64 };
            InvalidDataException exception = useBytes
                ? Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.OpenDataReader(File.ReadAllBytes(path), options))
                : Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.OpenDataReader(path, options));
            Assert.Contains("[Content_Types].xml", exception.Message, StringComparison.Ordinal);
            Assert.Contains("64 bytes", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void MetadataPartLimitCountsBytesReadInsteadOfTrustingStreamMetadata() {
        using var stream = new NonSeekableReadStream(new byte[65]);
        var options = new ExcelReadOptions { MaxMetadataPartBytes = 64 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            ExcelDocumentReader.DrainMetadataPartStream(
                stream,
                "xl/workbook.xml",
                options));

        Assert.Contains("xl/workbook.xml", exception.Message, StringComparison.Ordinal);
        Assert.Contains("64 bytes", exception.Message, StringComparison.Ordinal);
    }

    private sealed class NonSeekableReadStream : Stream {
        private readonly MemoryStream _inner;

        internal NonSeekableReadStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
