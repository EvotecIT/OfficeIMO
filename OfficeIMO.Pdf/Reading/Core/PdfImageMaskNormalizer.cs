using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfImageMaskNormalizer {
    internal static bool TryBuildPngFile(
        int width,
        int height,
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (width <= 0 ||
            height <= 0 ||
            !IsImageMask(stream, objects) ||
            !PdfImageBufferLimits.TryGetScanlineBufferSize(
                width,
                height,
                2,
                maxDecodedStreamBytes,
                out _,
                out int scanlineBytes) ||
            !TryReadDecodedStreamBytes(stream, objects, maxDecodedStreamBytes, out var maskPixels)) {
            return false;
        }

        long sourceRowLengthLong = ((long)width + 7) / 8;
        long expectedLengthLong = sourceRowLengthLong * height;
        long outputRowLengthLong = (long)width * 2;
        if (sourceRowLengthLong > int.MaxValue ||
            expectedLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int sourceRowLength = (int)sourceRowLengthLong;
        int expectedLength = (int)expectedLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (maskPixels.Length < expectedLength) {
            return false;
        }

        if (!PdfImageDecodeTransform.TryCreateIndexed(stream.Dictionary, objects, out PdfImageDecodeTransform? decodeTransform)) {
            return false;
        }
        byte[] scanlines = new byte[scanlineBytes];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int sourceRow = row * sourceRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int sample = ReadMaskSample(maskPixels, sourceRow, pixel);
                if (decodeTransform is not null) {
                    sample = decodeTransform.TransformIndexedSample(sample, 1, 1);
                }

                int outputPixel = outputRow + 1 + pixel * 2;
                scanlines[outputPixel] = 0;
                scanlines[outputPixel + 1] = sample == 0 ? (byte)0 : (byte)255;
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            8,
            4,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    internal static bool IsImageMask(PdfStream stream, Dictionary<int, PdfIndirectObject> objects) {
        return stream.Dictionary.Items.TryGetValue("ImageMask", out var imageMaskObj) &&
            PdfObjectLookup.ResolveChain(objects, imageMaskObj) is PdfBoolean imageMask &&
            imageMask.Value;
    }

    private static bool TryReadDecodedStreamBytes(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out byte[] bytes) {
        return PdfImageStreamDecoder.TryDecode(stream, objects, out bytes, maxDecodedStreamBytes) && bytes.Length > 0;
    }

    private static int ReadMaskSample(byte[] maskPixels, int rowOffset, int pixelIndex) {
        int sourceByte = maskPixels[rowOffset + pixelIndex / 8];
        int shift = 7 - (pixelIndex % 8);
        return (sourceByte >> shift) & 1;
    }
}
