using System.Buffers.Binary;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingFontMetricsValidationTests {
        [Theory]
        [InlineData("truncated")]
        [InlineData("zero")]
        [InlineData("excess")]
        public void FontLoadingRejectsInvalidHorizontalMetricsBeforeMeasurement(string corruption) {
            byte[] data = ManagedTextShapingTestAssets.CreateFont('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
            OfficeTrueTypeFont valid = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(data));
            Assert.True(valid.HasGlyphs("0123456789"));
            Assert.True(valid.Measure("0123456789", 11D) > 0D);
            int count = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(4));
            for (int index = 0; index < count; index++) {
                int record = 12 + index * 16;
                string tag = Encoding.ASCII.GetString(data, record, 4);
                int offset = (int)BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(record + 8));
                if (tag == "hmtx" && corruption == "truncated")
                    BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(record + 12), 2U);
                if (tag == "hhea" && corruption != "truncated")
                    BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(offset + 34), corruption == "zero" ? (ushort)0 : ushort.MaxValue);
            }

            // Installed-family discovery uses this parser. A rejected face follows
            // the existing missing-font fallback instead of failing during sizing.
            Assert.Null(OfficeTrueTypeFont.TryLoad(data));
        }
    }
}
