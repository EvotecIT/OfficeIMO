using System;
using System.Buffers.Binary;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingFontAverageAdvanceMetricsTests {
        [Fact]
        public void AverageAdvanceRetainsSourceDesignUnitsAndEmWithoutReplacingGlyphWidths() {
            byte[] font = ManagedTextShapingTestAssets.CreateAverageAdvanceMetricsFont(513);
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(TableOffset(font, "head") + 18), 2048);
            OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(font));
            int glyph = reader.MapGlyph('A');
            Assert.Equal(500, reader.AdvanceWidth(glyph));

            OfficeOpenTypeAverageAdvanceMetrics fromBytes = Assert.IsType<OfficeOpenTypeAverageAdvanceMetrics>(
                OfficeOpenTypeAverageAdvanceMetrics.TryRead(font));
            OfficeOpenTypeAverageAdvanceMetrics fromReader = Assert.IsType<OfficeOpenTypeAverageAdvanceMetrics>(
                OfficeOpenTypeAverageAdvanceMetrics.TryRead(reader));

            Assert.Equal(513, fromBytes.AverageAdvance);
            Assert.Equal(2048, fromBytes.UnitsPerEm);
            Assert.Equal(fromBytes.AverageAdvance, fromReader.AverageAdvance);
            Assert.Equal(fromBytes.UnitsPerEm, fromReader.UnitsPerEm);
            Assert.Equal(500, reader.AdvanceWidth(glyph));
        }

        [Fact]
        public void AverageAdvanceRespectsDeclaredOs2LengthWhenBackingBytesExist() {
            byte[] font = ManagedTextShapingTestAssets.CreateAverageAdvanceMetricsFont(513);
            BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(TableRecord(font, "OS/2") + 12), 3);
            OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(font));
            Assert.True(reader.TryGetTable("OS/2", out int offset, out int length));
            Assert.Equal(3, length);
            Assert.Equal(513, reader.ReadInt16(offset + 2));

            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(reader));
            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(font));
            Assert.Equal(500, reader.AdvanceWidth(reader.MapGlyph('A')));
        }

        [Fact]
        public void AverageAdvanceKeepsMissingOs2UnavailableWithoutRejectingTheFont() {
            byte[] font = ManagedTextShapingTestAssets.CreateAverageAdvanceMetricsFont(null);
            OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(font));
            Assert.False(reader.TryGetTable("OS/2", out _, out _));

            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(reader));
            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(font));
            Assert.Equal(500, reader.AdvanceWidth(reader.MapGlyph('A')));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void AverageAdvanceKeepsSignedNonpositiveValuesUnavailable(short averageAdvance) {
            byte[] font = ManagedTextShapingTestAssets.CreateAverageAdvanceMetricsFont(averageAdvance);
            OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(font));

            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(reader));
            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(font));
            Assert.Equal(500, reader.AdvanceWidth(reader.MapGlyph('A')));
        }

        [Fact]
        public void AverageAdvanceRequiresValidSharedFontMetrics() {
            byte[] font = ManagedTextShapingTestAssets.CreateAverageAdvanceMetricsFont(513);
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(TableOffset(font, "head") + 18), 0);
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(font);

            Assert.Null(reader);
            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(reader));
            Assert.Null(OfficeOpenTypeAverageAdvanceMetrics.TryRead(font));
        }

        private static int TableOffset(byte[] font, string tag) =>
            checked((int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(TableRecord(font, tag) + 8)));

        private static int TableRecord(byte[] font, string tag) {
            for (int index = 0; index < BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4)); index++) {
                int record = 12 + index * 16;
                if (System.Text.Encoding.ASCII.GetString(font, record, 4) == tag) return record;
            }
            throw new InvalidOperationException("The synthetic font is missing its requested table.");
        }
    }
}

namespace OfficeIMO.TestAssets {
    internal static partial class ManagedTextShapingTestAssets {
        internal static byte[] CreateAverageAdvanceMetricsFont(short? averageAdvance) {
            byte[]? os2 = null;
            if (averageAdvance.HasValue) {
                os2 = new byte[4];
                BinaryPrimitives.WriteInt16BigEndian(os2.AsSpan(2), averageAdvance.Value);
            }
            return CreateFontFromCmap(CreateFormat12Cmap(new[] { (int)'A' }), includeTrailingMetric: true, os2: os2);
        }
    }
}
