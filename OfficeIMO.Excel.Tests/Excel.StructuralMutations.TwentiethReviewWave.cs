using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InCellImage_ReplacementDetachesPackageSharedImagePart() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Original");
            ExtendedPart relationshipPart = document.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<ExtendedPart>()
                .Single(part => part.RelationshipType.EndsWith("/richValueRel", StringComparison.Ordinal));
            OpenXmlPart sharedImagePart = Assert.Single(relationshipPart.Parts).OpenXmlPart;
            const string sharedRelationshipId = "rIdSharedInCellImage";
            sheet.WorksheetPart.AddPart(sharedImagePart, sharedRelationshipId);
            byte[] replacement = TinyPng.Concat(new byte[] { 0 }).ToArray();

            sheet.SetInCellImage(1, 1, replacement, altText: "Replacement");

            OpenXmlPart retainedSharedPart = sheet.WorksheetPart.GetPartById(sharedRelationshipId);
            using Stream retainedStream = retainedSharedPart.GetStream(FileMode.Open, FileAccess.Read);
            using var retainedBuffer = new MemoryStream();
            retainedStream.CopyTo(retainedBuffer);
            Assert.Equal(TinyPng, retainedBuffer.ToArray());
            Assert.Equal(replacement, Assert.Single(sheet.GetInCellImages()).Bytes);
        }

        [Fact]
        public void Test_InCellImage_NonSeekablePayloadStopsAtAggregateBudget() {
            using var source = new CountingNonSeekableReadStream(length: 1_000_000);

            Assert.Throws<InvalidOperationException>(() => ExcelSheet.ReadInCellImagePayload(source, 32, 32));

            Assert.InRange(source.BytesRead, 33, 100_000);
        }

        [Fact]
        public void Test_AutoFilterTopBottom_UsesPercentageSpecificMaximum() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");

            sheet.AutoFilterTopBottom("A1:A10", 0, 100, percent: true);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                sheet.AutoFilterTopBottom("A1:A10", 0, 101, percent: true));
            sheet.AutoFilterTopBottom("A1:A10", 0, 500, percent: false);
        }

        private sealed class CountingNonSeekableReadStream : Stream {
            private readonly long _length;
            private long _position;

            internal CountingNonSeekableReadStream(long length) {
                _length = length;
            }

            internal long BytesRead => _position;
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override int Read(byte[] buffer, int offset, int count) {
                int available = (int)Math.Min(count, _length - _position);
                if (available <= 0) return 0;
                Array.Clear(buffer, offset, available);
                _position += available;
                return available;
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
