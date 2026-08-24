using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class DrawingRasterTests {
        [Fact]
        public void OfficeRasterImageDecoder_DecodesGifFirstFrameThroughSharedRasterPath() {
            byte[] gif = CreateIndexedGif(
                2,
                2,
                new[] { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White },
                new byte[] { 0, 1, 2, 3 });

            Assert.True(OfficeRasterImageDecoder.TryDecode(gif, out OfficeRasterImage? image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(2, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.White, image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeRasterContainerInspectorClassifiesUntimedSingleFrameGifAsStatic() {
            byte[] gif = CreateSinglePixelGif();

            Assert.True(OfficeRasterContainerInspector.TryInspect(gif, out OfficeRasterContainerInfo? container));
            Assert.False(container!.IsAnimated);
            Assert.Equal(OfficeRasterFrameKind.Image, container.Frames[0].Kind);
        }

        [Fact]
        public void OfficeRasterContainerInspectorRetainsSingleFrameGifTimingSemantics() {
            byte[] gif = CreateSinglePixelGif();
            int descriptor = Array.IndexOf(gif, (byte)0x2C);
            byte[] timed = gif.Take(descriptor)
                .Concat(new byte[] { 0x21, 0xF9, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00 })
                .Concat(gif.Skip(descriptor))
                .ToArray();

            Assert.True(OfficeRasterContainerInspector.TryInspect(timed, out OfficeRasterContainerInfo? container));
            Assert.True(container!.IsAnimated);
            Assert.Equal(TimeSpan.FromMilliseconds(10), container.Frames[0].Duration);
            Assert.True(OfficeRasterImageDecoder.TryDecode(
                timed,
                new OfficeRasterDecodeOptions { AnimationPolicy = OfficeRasterAnimationPolicy.UseSelectedFrame },
                out _,
                out OfficeRasterDecodeInfo info));
            Assert.True(info.AnimationDiscarded);
            Assert.False(info.FramesOrPagesDiscarded);
            Assert.NotNull(info.Diagnostic);
            Assert.False(OfficeRasterImageDecoder.TryDecode(
                timed,
                new OfficeRasterDecodeOptions { AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated },
                out _,
                out _));
        }

        [Fact]
        public void OfficeRasterContainerInspectorRetainsSingleFrameGifLoopSemantics() {
            byte[] gif = CreateSinglePixelGif();
            int descriptor = Array.IndexOf(gif, (byte)0x2C);
            byte[] loopExtension = {
                0x21, 0xFF, 0x0B,
                (byte)'N', (byte)'E', (byte)'T', (byte)'S', (byte)'C', (byte)'A',
                (byte)'P', (byte)'E', (byte)'2', (byte)'.', (byte)'0',
                0x03, 0x01, 0x00, 0x00, 0x00
            };
            byte[] looping = gif.Take(descriptor)
                .Concat(loopExtension)
                .Concat(gif.Skip(descriptor))
                .ToArray();

            Assert.True(OfficeRasterContainerInspector.TryInspect(looping, out OfficeRasterContainerInfo? container));
            Assert.True(container!.IsAnimated);
            Assert.Equal(0, container.LoopCount);
        }

        [Fact]
        public void OfficeRasterImageDecoder_DecodesInterlacedGifRowsThroughSharedRasterPath() {
            byte[] gif = CreateIndexedGif(
                1,
                4,
                new[] { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White },
                new byte[] { 0, 1, 2, 3 },
                interlaced: true);

            Assert.True(OfficeRasterImageDecoder.TryDecode(gif, out OfficeRasterImage? image));
            Assert.Equal(1, image!.Width);
            Assert.Equal(4, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 2));
            Assert.Equal(OfficeColor.White, image.GetPixel(0, 3));
        }

        [Fact]
        public void OfficeRasterImageDecoder_FillsLogicalGifCanvasWithBackgroundColor() {
            byte[] gif = CreateIndexedGif(
                4,
                4,
                new[] { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White },
                new byte[] { 0 },
                imageLeft: 1,
                imageTop: 1,
                imageWidth: 1,
                imageHeight: 1,
                backgroundColorIndex: 1);

            Assert.True(OfficeRasterImageDecoder.TryDecode(gif, out OfficeRasterImage? image));
            Assert.True(OfficeRasterContainerInspector.TryInspect(gif, out OfficeRasterContainerInfo? container));
            Assert.Equal(OfficeColor.Lime, image!.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Red, image.GetPixel(1, 1));
            Assert.Equal(OfficeColor.Lime, container!.Background);

            const int imageDescriptorOffset = 25;
            byte[] transparentBackground = gif.Take(imageDescriptorOffset)
                .Concat(new byte[] { 0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x01, 0x00 })
                .Concat(gif.Skip(imageDescriptorOffset))
                .ToArray();
            Assert.True(OfficeRasterContainerInspector.TryInspect(
                transparentBackground, out OfficeRasterContainerInfo? transparentContainer));
            Assert.Equal(OfficeColor.Transparent, transparentContainer!.Background);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_PaintsDecodedGifImages() {
            byte[] gif = CreateSinglePixelGif();
            OfficeDrawing drawing = new OfficeDrawing(20, 16);
            drawing.AddImage(
                gif,
                "image/gif",
                new OfficeImageProjection(new OfficeImagePlacement(4, 3, 8, 6)),
                "GIF marker");

            OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.Black);

            Assert.Equal(OfficeColor.White, rendered.GetPixel(7, 5));
        }

        [Fact]
        public void OfficeRasterImageDecoder_SelectsCompositedGifFrameAndReportsAnimationLoss() {
            byte[] gif = CreateTwoFrameGif();
            var options = new OfficeRasterDecodeOptions { FrameIndex = 1 };

            Assert.True(OfficeRasterImageDecoder.TryDecode(gif, options, out OfficeRasterImage? image, out OfficeRasterDecodeInfo info));

            Assert.Equal(OfficeColor.Lime, image!.GetPixel(0, 0));
            Assert.Equal(2, info.FrameCount);
            Assert.Equal(1, info.SelectedFrameIndex);
            Assert.True(info.Succeeded);
            Assert.True(info.IsAnimated);
            Assert.True(info.AnimationDiscarded);
            Assert.NotNull(info.Diagnostic);
        }

        [Fact]
        public void OfficeRasterImageDecoder_RejectsAnimatedGifWhenPolicyRequiresExactStaticInput() {
            byte[] gif = CreateTwoFrameGif();
            var options = new OfficeRasterDecodeOptions {
                AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated
            };

            Assert.False(OfficeRasterImageDecoder.TryDecode(gif, options, out OfficeRasterImage? image, out OfficeRasterDecodeInfo info));

            Assert.Null(image);
            Assert.False(info.Succeeded);
            Assert.Equal(2, info.FrameCount);
            Assert.False(info.AnimationDiscarded);
            Assert.Contains("rejected", info.Diagnostic, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OfficeGifReader_ClearsSelectedFrameWhenTrailingContentIsMalformed() {
            byte[] valid = CreateTwoFrameGif();
            byte[] malformed = valid.Take(valid.Length - 1).Concat(new byte[] { 0x21 }).ToArray();

            Assert.False(OfficeGifReader.TryDecodeFrame(malformed, 0, out OfficeRasterImage? image, out int frameCount));

            Assert.Null(image);
            Assert.Equal(2, frameCount);
        }

        [Fact]
        public void GifInspectionAndDecodeRequireTheTerminalTrailer() {
            byte[] valid = CreateSinglePixelGif();
            byte[] truncated = valid.Take(valid.Length - 1).ToArray();

            Assert.False(OfficeRasterContainerInspector.TryInspect(truncated, out _));
            Assert.False(OfficeGifReader.TryDecodeFrame(truncated, 0, out _, out _));
            Assert.False(OfficeRasterImageDecoder.TryDecode(truncated, out _));
        }

        [Fact]
        public void GifInspectionRejectsFramesWithoutAnActiveColorTable() {
            byte[] valid = CreateSinglePixelGif();
            int globalTableBytes = 3 << ((valid[10] & 7) + 1);
            byte[] malformed = valid.Take(13)
                .Concat(valid.Skip(13 + globalTableBytes))
                .ToArray();
            malformed[10] &= 0x7F;

            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed, out _));
            Assert.False(OfficeGifReader.TryDecodeFrame(malformed, 0, out _, out _));
            Assert.False(OfficeRasterImageDecoder.TryDecode(malformed, out _));
        }

        [Theory]
        [InlineData(1)]
        [InlineData(9)]
        public void GifInspectionRejectsInvalidLzwMinimumCodeSize(byte minimumCodeSize) {
            byte[] malformed = CreateSinglePixelGif();
            int imageDescriptorOffset = Array.IndexOf(malformed, (byte)0x2C);
            malformed[imageDescriptorOffset + 10] = minimumCodeSize;

            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed, out _));
            Assert.False(OfficeRasterImageDecoder.TryDecode(malformed, out _));
        }

        [Fact]
        public void OfficeGifReader_SkipsLzwExpansionForUnselectedTrailingFrames() {
            byte[] gif = CreateTwoFrameGif(out int secondFrameDescriptorOffset);
            gif[secondFrameDescriptorOffset + 12] = 0x07;

            Assert.True(OfficeGifReader.TryDecodeFrame(gif, 0, out OfficeRasterImage? selected, out int frameCount));
            Assert.False(OfficeGifReader.TryDecodeFrame(gif, 1, out OfficeRasterImage? malformed, out _));

            Assert.NotNull(selected);
            Assert.Equal(2, frameCount);
            Assert.Null(malformed);
        }

        [Fact]
        public void OfficeGifReaderTransfersSelectedCanvasWithoutFullCanvasSnapshots() {
            byte[] source = CreateIndexedGif(
                2048,
                2048,
                new[] { OfficeColor.Red, OfficeColor.Black },
                new byte[] { 0 },
                imageWidth: 1,
                imageHeight: 1);
            int descriptorOffset = Array.IndexOf(source, (byte)0x2C);
            byte[] previousDisposal = { 0x21, 0xF9, 0x04, 0x0C, 0, 0, 0, 0 };
            byte[] gif = source.Take(descriptorOffset)
                .Concat(previousDisposal)
                .Concat(source.Skip(descriptorOffset))
                .ToArray();
            OfficeGifReader.TryDecodeFrame(CreateSinglePixelGif(), 0, out _, out _);

#if NET8_0_OR_GREATER
            long before = GC.GetAllocatedBytesForCurrentThread();
#endif
            Assert.True(OfficeGifReader.TryDecodeFrame(
                gif, 0, out OfficeRasterImage? selected, out int frameCount));
#if NET8_0_OR_GREATER
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
#endif

            Assert.Equal(1, frameCount);
            Assert.Equal((2048, 2048), (selected!.Width, selected.Height));
            Assert.Equal(OfficeColor.Red, selected.GetPixel(0, 0));
#if NET8_0_OR_GREATER
            Assert.True(allocated < 24L * 1024L * 1024L,
                $"Selected GIF frame allocated {allocated:N0} bytes.");
#endif
        }

        [Fact]
        public void CompleteContentValidationRejectsMalformedTrailingGifFrame() {
            byte[] gif = CreateTwoFrameGif(out int secondFrameDescriptorOffset);
            gif[secondFrameDescriptorOffset + 12] = 0x07;

            Assert.True(OfficeGifReader.TryDecodeFrame(gif, 0, out OfficeRasterImage? selected, out int frameCount));
            Assert.NotNull(selected);
            Assert.Equal(2, frameCount);
            Assert.False(OfficeImageReader.TryValidateContent(gif, "animated.gif", out _));
        }

        [Fact]
        public void CompleteContentValidationRejectsBytesAfterGifTrailer() {
            byte[] gif = CreateTwoFrameGif();
            byte[] withTrailingBytes = new byte[gif.Length + 1];
            Buffer.BlockCopy(gif, 0, withTrailingBytes, 0, gif.Length);
            withTrailingBytes[withTrailingBytes.Length - 1] = 0x00;

            Assert.False(OfficeImageReader.TryValidateContent(withTrailingBytes, "trailing.gif", out _));
        }

        [Fact]
        public void CompleteContentValidationRejectsFullBytesAfterGifLzwEndCode() {
            byte[] valid = CreateIndexedGif(
                1,
                1,
                new[] { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White },
                new byte[] { 0 });
            const int imageDescriptorOffset = 25;
            int blockLengthOffset = imageDescriptorOffset + 11;
            int insertOffset = blockLengthOffset + 1 + valid[blockLengthOffset];
            var malformed = valid.ToList();
            malformed.Insert(insertOffset, 0x00);
            malformed[blockLengthOffset]++;

            Assert.False(OfficeImageReader.TryValidateContent(malformed.ToArray(), "trailing-lzw.gif", out _));
        }

        [Theory]
        [InlineData(0xFF)]
        [InlineData(0x01)]
        public void CompleteContentValidationRejectsMalformedKnownGifExtensionHeaders(byte extensionLabel) {
            byte[] valid = CreateSinglePixelGif();
            int imageDescriptorOffset = Array.IndexOf(valid, (byte)0x2C);
            var malformed = valid.ToList();
            malformed.InsertRange(imageDescriptorOffset, new byte[] { 0x21, extensionLabel, 0x01, 0x41, 0x00 });

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed.ToArray(), 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed.ToArray(), "extension.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed.ToArray(), out _));
        }

        [Theory]
        [InlineData(0x80)]
        [InlineData(0x10)]
        [InlineData(0x1C)]
        public void GifInspectionAndContentValidationRejectReservedGraphicControlValues(byte packed) {
            byte[] valid = CreateSinglePixelGif();
            int imageDescriptorOffset = Array.IndexOf(valid, (byte)0x2C);
            var malformed = valid.ToList();
            malformed.InsertRange(
                imageDescriptorOffset,
                new byte[] { 0x21, 0xF9, 0x04, packed, 0x00, 0x00, 0x00, 0x00 });

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed.ToArray(), 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed.ToArray(), "reserved-gce.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed.ToArray(), out _));
        }

        [Fact]
        public void CompleteContentValidationRejectsDanglingOrRepeatedGraphicControls() {
            byte[] valid = CreateSinglePixelGif();
            int imageDescriptorOffset = Array.IndexOf(valid, (byte)0x2C);
            int trailerOffset = Array.LastIndexOf(valid, (byte)0x3B);
            byte[] graphicControl = { 0x21, 0xF9, 0x04, 0, 0, 0, 0, 0 };
            byte[] dangling = valid.Take(trailerOffset)
                .Concat(graphicControl)
                .Concat(valid.Skip(trailerOffset))
                .ToArray();
            byte[] repeated = valid.Take(imageDescriptorOffset)
                .Concat(graphicControl)
                .Concat(graphicControl)
                .Concat(valid.Skip(imageDescriptorOffset))
                .ToArray();

            Assert.True(OfficeGifReader.TryDecodeFrame(dangling, 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(dangling, "dangling-control.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(dangling, out _));
            Assert.True(OfficeGifReader.TryDecodeFrame(repeated, 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(repeated, "repeated-control.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(repeated, out _));
        }

        [Fact]
        public void PlainTextExtensionsConsumePendingGraphicControlState() {
            byte[] valid = CreateIndexedGif(
                1,
                1,
                new[] { OfficeColor.Red, OfficeColor.Lime },
                new byte[] { 0 });
            int imageDescriptorOffset = Array.IndexOf(valid, (byte)0x2C);
            var withPlainText = valid.ToList();
            withPlainText.InsertRange(
                imageDescriptorOffset,
                new byte[] {
                    0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00,
                    0x21, 0x01, 0x0C,
                    0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x01, 0x00, 0x01,
                    0x00
                });

            Assert.True(OfficeGifReader.TryDecodeFrame(withPlainText.ToArray(), 0, out OfficeRasterImage? image, out _));
            Assert.Equal(OfficeColor.Red, Assert.IsType<OfficeRasterImage>(image).GetPixel(0, 0));
            Assert.True(OfficeImageReader.TryValidateContent(withPlainText.ToArray(), "plain-text.gif", out _));
            Assert.True(OfficeRasterContainerInspector.TryInspect(withPlainText.ToArray(), out _));
        }

        [Fact]
        public void CompleteContentValidationBoundsPlainTextTransparencyAgainstTheGlobalTable() {
            byte[] valid = CreateIndexedGif(
                1,
                1,
                new[] { OfficeColor.Red, OfficeColor.Lime },
                new byte[] { 0 });
            int imageDescriptorOffset = Array.IndexOf(valid, (byte)0x2C);
            var malformed = valid.ToList();
            malformed.InsertRange(
                imageDescriptorOffset,
                new byte[] {
                    0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x02, 0x00,
                    0x21, 0x01, 0x0C,
                    0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x01, 0x00, 0x01,
                    0x00
                });

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed.ToArray(), 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed.ToArray(), "plain-text-transparency.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed.ToArray(), out _));
        }

        [Fact]
        public void CompleteContentValidationRejectsTransparencyIndexesOutsideTheActiveColorTable() {
            byte[] valid = CreateSinglePixelGif();
            int imageDescriptorOffset = Array.IndexOf(valid, (byte)0x2C);
            var malformed = valid.ToList();
            malformed.InsertRange(
                imageDescriptorOffset,
                new byte[] { 0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x02, 0x00 });

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed.ToArray(), 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed.ToArray(), "transparency-index.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed.ToArray(), out _));
        }

        [Theory]
        [InlineData(0x08)]
        [InlineData(0x10)]
        public void CompleteContentValidationRejectsReservedImageDescriptorBits(byte reservedBit) {
            byte[] malformed = CreateSinglePixelGif();
            int imageDescriptorOffset = Array.IndexOf(malformed, (byte)0x2C);
            malformed[imageDescriptorOffset + 9] |= reservedBit;

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed, 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed, "reserved-descriptor.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed, out _));
        }

        [Fact]
        public void CompleteContentValidationRejectsBackgroundIndexesOutsideTheGlobalColorTable() {
            byte[] malformed = CreateIndexedGif(
                1,
                1,
                new[] { OfficeColor.Red, OfficeColor.Lime },
                new byte[] { 0 },
                backgroundColorIndex: 2);

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed, 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed, "background-index.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed, out _));
        }

        [Fact]
        public void CompleteContentValidationRequiresZeroBackgroundIndexWithoutAGlobalColorTable() {
            byte[] malformed = CreateGifWithOnlyALocalColorTable(backgroundColorIndex: 1);

            Assert.True(OfficeGifReader.TryDecodeFrame(malformed, 0, out _, out _));
            Assert.False(OfficeImageReader.TryValidateContent(malformed, "local-palette.gif", out _));
            Assert.False(OfficeRasterContainerInspector.TryInspect(malformed, out _));
        }

        private static byte[] CreateSinglePixelGif() =>
            Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

        private static byte[] CreateGifWithOnlyALocalColorTable(byte backgroundColorIndex) {
            byte[] source = CreateSinglePixelGif();
            const int logicalScreenLength = 13;
            const int colorTableLength = 6;
            const int imageDescriptorLength = 10;
            int sourceDescriptorOffset = logicalScreenLength + colorTableLength;
            var result = new List<byte>(source.Length);
            result.AddRange(source.Take(logicalScreenLength));
            result[10] &= 0x7F;
            result[11] = backgroundColorIndex;
            result.AddRange(source.Skip(sourceDescriptorOffset).Take(imageDescriptorLength));
            result[result.Count - 1] |= 0x80;
            result.AddRange(source.Skip(logicalScreenLength).Take(colorTableLength));
            result.AddRange(source.Skip(sourceDescriptorOffset + imageDescriptorLength));
            return result.ToArray();
        }

        private static byte[] CreateTwoFrameGif() => CreateTwoFrameGif(out _);

        private static byte[] CreateTwoFrameGif(out int secondFrameDescriptorOffset) {
            OfficeColor[] palette = { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White };
            byte[] first = CreateIndexedGif(1, 1, palette, new byte[] { 0 });
            byte[] second = CreateIndexedGif(1, 1, palette, new byte[] { 1 });
            const int imageDescriptorOffset = 25;
            var result = new List<byte>(first.Length + second.Length - imageDescriptorOffset);
            result.AddRange(first.Take(first.Length - 1));
            secondFrameDescriptorOffset = result.Count;
            result.AddRange(second.Skip(imageDescriptorOffset).Take(second.Length - imageDescriptorOffset - 1));
            result.Add(0x3B);
            return result.ToArray();
        }

        private static byte[] CreateIndexedGif(
            int width,
            int height,
            IReadOnlyList<OfficeColor> palette,
            IReadOnlyList<byte> pixels,
            bool interlaced = false,
            int imageLeft = 0,
            int imageTop = 0,
            int? imageWidth = null,
            int? imageHeight = null,
            int backgroundColorIndex = 0) {
            int frameWidth = imageWidth ?? width;
            int frameHeight = imageHeight ?? height;
            if (pixels.Count != frameWidth * frameHeight) {
                throw new ArgumentException("Pixel count must match GIF dimensions.", nameof(pixels));
            }

            int colorTableSize = 2;
            while (colorTableSize < palette.Count) {
                colorTableSize *= 2;
            }

            int minimumCodeSize = Math.Max(2, GetRequiredBits(colorTableSize - 1));
            byte[] lzw = EncodeGifLzw(ReorderGifPixels(frameWidth, frameHeight, pixels, interlaced), minimumCodeSize);
            var bytes = new List<byte>();
            bytes.AddRange(new byte[] { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a' });
            WriteUInt16LittleEndian(bytes, width);
            WriteUInt16LittleEndian(bytes, height);
            bytes.Add((byte)(0x80 | ((minimumCodeSize - 1) << 4) | (GetRequiredBits(colorTableSize - 1) - 1)));
            bytes.Add((byte)backgroundColorIndex);
            bytes.Add(0);
            for (int i = 0; i < colorTableSize; i++) {
                OfficeColor color = i < palette.Count ? palette[i] : OfficeColor.Black;
                bytes.Add(color.R);
                bytes.Add(color.G);
                bytes.Add(color.B);
            }

            bytes.Add(0x2C);
            WriteUInt16LittleEndian(bytes, imageLeft);
            WriteUInt16LittleEndian(bytes, imageTop);
            WriteUInt16LittleEndian(bytes, frameWidth);
            WriteUInt16LittleEndian(bytes, frameHeight);
            bytes.Add(interlaced ? (byte)0x40 : (byte)0x00);
            bytes.Add((byte)minimumCodeSize);
            bytes.Add((byte)lzw.Length);
            bytes.AddRange(lzw);
            bytes.Add(0);
            bytes.Add(0x3B);
            return bytes.ToArray();
        }

        private static byte[] ReorderGifPixels(int width, int height, IReadOnlyList<byte> pixels, bool interlaced) {
            if (!interlaced) {
                return pixels.ToArray();
            }

            var reordered = new List<byte>(pixels.Count);
            foreach (int y in EnumerateGifRows(height)) {
                for (int x = 0; x < width; x++) {
                    reordered.Add(pixels[(y * width) + x]);
                }
            }

            return reordered.ToArray();
        }

        private static byte[] EncodeGifLzw(IReadOnlyList<byte> indices, int minimumCodeSize) {
            int clearCode = 1 << minimumCodeSize;
            int endCode = clearCode + 1;
            int dictionaryCount = clearCode + 2;
            int codeSize = minimumCodeSize + 1;
            int previousCode = -1;
            var bits = new List<int>();

            WriteBits(bits, clearCode, codeSize);
            for (int i = 0; i < indices.Count; i++) {
                WriteBits(bits, indices[i], codeSize);
                if (previousCode >= 0 && dictionaryCount < 4096) {
                    dictionaryCount++;
                    if (dictionaryCount == (1 << codeSize) && codeSize < 12) {
                        codeSize++;
                    }
                }

                previousCode = indices[i];
            }

            WriteBits(bits, endCode, codeSize);
            var bytes = new byte[(bits.Count + 7) / 8];
            for (int i = 0; i < bits.Count; i++) {
                bytes[i / 8] |= (byte)(bits[i] << (i % 8));
            }

            return bytes;
        }

        private static void WriteBits(List<int> bits, int value, int count) {
            for (int i = 0; i < count; i++) {
                bits.Add((value >> i) & 1);
            }
        }

        private static IEnumerable<int> EnumerateGifRows(int height) {
            int[] starts = { 0, 4, 2, 1 };
            int[] steps = { 8, 8, 4, 2 };
            for (int pass = 0; pass < starts.Length; pass++) {
                for (int y = starts[pass]; y < height; y += steps[pass]) {
                    yield return y;
                }
            }
        }

        private static int GetRequiredBits(int value) {
            int bits = 0;
            do {
                bits++;
                value >>= 1;
            } while (value > 0);

            return bits;
        }

        private static void WriteUInt16LittleEndian(List<byte> bytes, int value) {
            bytes.Add((byte)(value & 0xFF));
            bytes.Add((byte)((value >> 8) & 0xFF));
        }
    }
}
