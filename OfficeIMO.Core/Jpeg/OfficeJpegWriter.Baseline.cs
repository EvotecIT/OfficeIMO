using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegWriter {
    // Baseline encoding with fixed Huffman tables consumes each MCU immediately.
    // Progressive scans and optimized tables retain coefficients in the existing path.
    private static void EncodeBaselinePixels(
        Stream stream, int width, int height, byte[] rgba, int stride, int rowOffset, int rowStride,
        ComponentSpec[] components, int maxH, int maxV, int[] qY, int[] qC, HuffmanTableSet tables,
        CancellationToken cancellationToken, Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        var allComponents = new int[components.Length];
        for (int index = 0; index < allComponents.Length; index++) allComponents[index] = index;
        WriteSos(stream, components, allComponents, 0, 63, 0, 0);

        var writer = new BitWriter(stream);
        var luma = new int[64];
        var cb = new int[64];
        var cr = new int[64];
        var quantized = new int[64];
        var workspace = new double[64];
        int previousY = 0, previousCb = 0, previousCr = 0;
        int mcuWidth = maxH * 8;
        int mcuHeight = maxV * 8;
        int columns = (width + mcuWidth - 1) / mcuWidth;
        int rows = (height + mcuHeight - 1) / mcuHeight;
        ComponentSpec y = components[0];
        bool hasChroma = components.Length > 1;
        for (int row = 0; row < rows; row++) {
            // The first row checkpoint runs before any output in WriteRgbaCore.
            if (row != 0) checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.JpegCoefficientRow);
            cancellationToken.ThrowIfCancellationRequested();
            for (int column = 0; column < columns; column++) {
                for (int vertical = 0; vertical < y.V; vertical++) {
                    for (int horizontal = 0; horizontal < y.H; horizontal++) {
                        LoadBlockLuma(rgba, stride, rowOffset, rowStride, width, height,
                            column * mcuWidth + horizontal * 8, row * mcuHeight + vertical * 8, luma);
                        EncodeBlock(writer, luma, qY, tables.DcLuma.Table, tables.AcLuma.Table,
                            ref previousY, quantized, workspace);
                    }
                }
                if (!hasChroma) continue;
                LoadBlockChroma(rgba, stride, rowOffset, rowStride, width, height,
                    column * mcuWidth, row * mcuHeight, maxH / components[1].H, maxV / components[1].V, cb, cr);
                EncodeBlock(writer, cb, qC, tables.DcChroma.Table, tables.AcChroma.Table,
                    ref previousCb, quantized, workspace);
                EncodeBlock(writer, cr, qC, tables.DcChroma.Table, tables.AcChroma.Table,
                    ref previousCr, quantized, workspace);
            }
        }
        writer.Flush();
    }
}
