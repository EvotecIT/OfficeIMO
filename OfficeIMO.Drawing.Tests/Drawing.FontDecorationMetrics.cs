using System;
using System.Buffers.Binary;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingFontDecorationMetricsTests {
        [Theory]
        [InlineData(1000, 250, -125, 0.25D, -0.125D, 0.05D, 0.025D)]
        [InlineData(2048, 250, -125, 0.1220703125D, -0.06103515625D, 0.0244140625D, 0.01220703125D)]
        [InlineData(1000, -250, 125, -0.25D, 0.125D, 0.05D, 0.025D)]
        public void DecorationMetricsNormalizeSignedFieldsByTheDeclaredEm(int unitsPerEm, short strikePosition,
            short underlinePosition, double expectedStrikePosition, double expectedUnderlinePosition,
            double expectedStrikeThickness, double expectedUnderlineThickness) {
            byte[] font = Font(strikePosition, 50, underlinePosition, 25, unitsPerEm);

            OfficeOpenTypeDecorationMetrics metrics = Read(font);

            OfficeOpenTypeDecorationMetrics.Metric strike = AssertMetric(metrics.Strikethrough);
            OfficeOpenTypeDecorationMetrics.Metric underline = AssertMetric(metrics.Underline);
            Assert.Equal(expectedStrikePosition, strike.PositionRatio, 12);
            Assert.Equal(expectedStrikeThickness, strike.ThicknessRatio, 12);
            Assert.Equal(expectedUnderlinePosition, underline.PositionRatio, 12);
            Assert.Equal(expectedUnderlineThickness, underline.ThicknessRatio, 12);
        }

        [Fact]
        public void DecorationMetricsDoNotReadStrikePositionFromTheFollowingTable() {
            byte[] os2 = new byte[28];
            WriteInt16(os2, 26, 50);
            byte[] font = ManagedTextShapingTestAssets.CreateDecorationMetricsFont(os2, Post(-125, 25));
            OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(font));
            Assert.True(reader.TryGetTable("OS/2", out int os2Offset, out int length));
            Assert.Equal(28, length);
            Assert.True(reader.TryGetTable("post", out int postOffset, out _));
            Assert.Equal(os2Offset + length, postOffset);
            Assert.Equal(3, reader.ReadInt16(postOffset));

            OfficeOpenTypeDecorationMetrics metrics = Read(font);

            Assert.Null(metrics.Strikethrough);
            Assert.Equal(-0.125D, AssertMetric(metrics.Underline).PositionRatio);
        }

        [Fact]
        public void DecorationMetricsDoNotReadUnderlineFromTheFollowingTable() {
            byte[] font = ManagedTextShapingTestAssets.CreateDecorationMetricsFont(Os2(250, 50), new byte[8]);
            OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(font));
            Assert.True(reader.TryGetTable("post", out int postOffset, out int length));
            Assert.Equal(8, length);
            Assert.True(reader.TryGetTable("bait", out int baitOffset, out _));
            Assert.Equal(postOffset + length, baitOffset);
            Assert.Equal(-125, reader.ReadInt16(baitOffset));
            Assert.Equal(25, reader.ReadInt16(baitOffset + 2));

            OfficeOpenTypeDecorationMetrics metrics = Read(font);

            Assert.Null(metrics.Underline);
            Assert.Equal(0.25D, AssertMetric(metrics.Strikethrough).PositionRatio);
        }

        [Theory]
        [InlineData("OS/2", 29)]
        [InlineData("post", 11)]
        public void DecorationMetricsRespectDeclaredLengthsEvenWhenTrailingFieldBytesExist(string tag, uint length) {
            byte[] font = Font(250, 50, -125, 25);
            int record = TableRecord(font, tag);
            BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(record + 12), length);
            Assert.NotNull(OfficeOpenTypeReader.TryCreate(font));

            OfficeOpenTypeDecorationMetrics metrics = Read(font);

            Assert.Equal(tag != "OS/2", metrics.Strikethrough.HasValue);
            Assert.Equal(tag != "post", metrics.Underline.HasValue);
        }

        [Theory]
        [InlineData("OS/2", 0)]
        [InlineData("OS/2", -1)]
        [InlineData("post", 0)]
        [InlineData("post", -1)]
        public void DecorationMetricsKeepNonpositiveThicknessUnavailableIndependently(string tag, short thickness) {
            byte[] font = Font(250, tag == "OS/2" ? thickness : (short)50,
                -125, tag == "post" ? thickness : (short)25);

            OfficeOpenTypeDecorationMetrics metrics = Read(font);

            Assert.Equal(tag != "OS/2", metrics.Strikethrough.HasValue);
            Assert.Equal(tag != "post", metrics.Underline.HasValue);
        }

        [Theory]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(false, false)]
        public void DecorationMetricsKeepMissingTablesUnavailable(bool strikePresent, bool underlinePresent) {
            byte[] font = ManagedTextShapingTestAssets.CreateDecorationMetricsFont(
                strikePresent ? Os2(250, 50) : null, underlinePresent ? Post(-125, 25) : null);

            OfficeOpenTypeDecorationMetrics? metrics = OfficeOpenTypeDecorationMetrics.TryRead(font);

            if (!strikePresent && !underlinePresent) {
                Assert.Null(metrics);
            } else {
                Assert.True(metrics.HasValue);
                Assert.Equal(strikePresent, metrics.Value.Strikethrough.HasValue);
                Assert.Equal(underlinePresent, metrics.Value.Underline.HasValue);
            }
        }

        [Fact]
        public void DecorationMetricsRequireAValidatedFontAndDoNotInventMissingFontMetrics() {
            byte[] font = Font(250, 50, -125, 25);
            int head = TableOffset(font, "head");
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(head + 18), 0);

            Assert.Null(OfficeOpenTypeDecorationMetrics.TryRead(font));
            Assert.Null(OfficeOpenTypeDecorationMetrics.TryRead(Font(250, 0, -125, -1)));
            Assert.Null(OfficeOpenTypeDecorationMetrics.TryRead(Array.Empty<byte>()));
            Assert.Null(OfficeOpenTypeDecorationMetrics.TryRead(null));
        }

        [Theory]
        [InlineData(2.5D, 1D, 500, 500, 2D, 2D)]
        [InlineData(2.5D, 1D, -500, 500, -2D, 2D)]
        [InlineData(0.625D, 0.25D, 500, 500, 0.5D, 0.5D)]
        [InlineData(0.125D, 0.5D, 500, 1, 0D, 0.5D)]
        public void DecorationDeviceMetricsRoundFontAndSignedMetricTiesBeforeReturningPoints(double fontSize,
            double pointsPerPixel, short position, short thickness, double expectedPosition, double expectedThickness) {
            OfficeOpenTypeDecorationMetrics.Metric metric = AssertMetric(Read(Font(position, thickness, -125, 25)).Strikethrough);

            metric.GetDeviceMetrics(fontSize, pointsPerPixel, out double positionPoints, out double thicknessPoints);

            Assert.Equal(expectedPosition, positionPoints, 12);
            Assert.Equal(expectedThickness, thicknessPoints, 12);
        }

        private static OfficeOpenTypeDecorationMetrics Read(byte[] font) =>
            Assert.IsType<OfficeOpenTypeDecorationMetrics>(OfficeOpenTypeDecorationMetrics.TryRead(font));

        private static OfficeOpenTypeDecorationMetrics.Metric AssertMetric(OfficeOpenTypeDecorationMetrics.Metric? metric) =>
            Assert.IsType<OfficeOpenTypeDecorationMetrics.Metric>(metric);

        private static byte[] Font(short strikePosition, short strikeThickness, short underlinePosition,
            short underlineThickness, int unitsPerEm = 1000) {
            byte[] font = ManagedTextShapingTestAssets.CreateDecorationMetricsFont(
                Os2(strikePosition, strikeThickness), Post(underlinePosition, underlineThickness));
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(TableOffset(font, "head") + 18), checked((ushort)unitsPerEm));
            return font;
        }

        private static byte[] Os2(short position, short thickness) {
            var data = new byte[30];
            WriteInt16(data, 26, thickness);
            WriteInt16(data, 28, position);
            return data;
        }

        private static byte[] Post(short position, short thickness) {
            var data = new byte[12];
            BinaryPrimitives.WriteUInt32BigEndian(data, 0x00030000U);
            WriteInt16(data, 8, position);
            WriteInt16(data, 10, thickness);
            return data;
        }

        private static void WriteInt16(byte[] data, int offset, short value) =>
            BinaryPrimitives.WriteInt16BigEndian(data.AsSpan(offset), value);

        private static int TableOffset(byte[] data, string tag) =>
            checked((int)BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(TableRecord(data, tag) + 8)));

        private static int TableRecord(byte[] data, string tag) {
            for (int index = 0; index < BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(4)); index++) {
                int record = 12 + index * 16;
                if (System.Text.Encoding.ASCII.GetString(data, record, 4) == tag) return record;
            }
            throw new InvalidOperationException("The synthetic font is missing its requested table.");
        }
    }
}

namespace OfficeIMO.TestAssets {
    internal static partial class ManagedTextShapingTestAssets {
        internal static byte[] CreateDecorationMetricsFont(byte[]? os2, byte[]? post) {
            var additional = new System.Collections.Generic.List<(string Tag, byte[] Data)>();
            if (post != null) additional.Add(("post", post));
            // A plausible signed position/positive thickness immediately after
            // a short post table exposes readers that use the whole font bound.
            additional.Add(("bait", new byte[] { 0xFF, 0x83, 0x00, 0x19 }));
            // Two glyphs with one long horizontal metric also require the
            // second glyph's trailing two-byte side bearing in canonical sfnt.
            return CreateFontFromCmap(CreateFormat12Cmap(new[] { (int)'A' }), includeTrailingMetric: true,
                os2: os2, additionalTables: additional);
        }
    }
}
