using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private static JpegFrame ParseFrameHeader(OfficeByteView data, byte frameMarker = 0xC0) {
        var precision = data[0];
        if ((frameMarker == 0xC3 || frameMarker == 0xCB) ? precision < 2 || precision > 16 :
            precision != 8 && !(precision == 12 && (frameMarker == 0xC1 || frameMarker == 0xC2 || frameMarker == 0xC9 || frameMarker == 0xCA))) throw new FormatException("Unsupported JPEG precision.");
        var height = ReadUInt16BE(data, 1);
        var width = ReadUInt16BE(data, 3);
        var components = data[5];
        if (width == 0 || height == 0) throw new FormatException("Invalid JPEG dimensions.");
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) {
            throw new FormatException(JpegDimensionsLimitMessage);
        }
        if (components < 1 || components > 255) {
            throw new FormatException("Unsupported JPEG component count.");
        }
        if (data.Length < 6 + components * 3) throw new FormatException("Invalid JPEG SOF segment.");

        var frame = new JpegFrame {
            Precision = precision,
            Width = width,
            Height = height,
            ComponentCount = components,
            Components = new Component[components]
        };

        var offset = 6;
        var maxH = 0;
        var maxV = 0;
        for (var i = 0; i < components; i++) {
            var id = data[offset++];
            var sampling = data[offset++];
            var h = sampling >> 4;
            var v = sampling & 0x0F;
            var qt = data[offset++];
            if (h == 0 || v == 0 || h > 4 || v > 4) throw new FormatException("Invalid JPEG sampling factors.");
            if (qt >= 4) throw new FormatException("Unsupported JPEG quantization table.");
            frame.Components[i] = new Component {
                Id = id,
                H = h,
                V = v,
                QuantId = qt
            };
            if (h > maxH) maxH = h;
            if (v > maxV) maxV = v;
        }
        frame.MaxH = maxH;
        frame.MaxV = maxV;
        return frame;
    }

    internal static bool IsSupportedRgbaFrameHeader(byte[] data, int offset, int length, byte frameMarker = 0xC0) {
        try {
            JpegFrame frame = ParseFrameHeader(new OfficeByteView(data).Slice(offset, length), frameMarker);
            return frame.ComponentCount is 1 or 3 or 4;
        } catch (Exception ex) when (ex is FormatException || ex is ArgumentException ||
                                     ex is IndexOutOfRangeException || ex is OverflowException) {
            return false;
        }
    }

    private static ScanHeader ParseScanHeader(OfficeByteView data, ref JpegFrame frame) {
        var components = data[0];
        if (components == 0 || components > 4 || components > frame.ComponentCount) throw new FormatException("Invalid JPEG scan component count.");
        if (data.Length < 1 + components * 2 + 3) throw new FormatException("Invalid JPEG scan header.");

        int samplingUnits = 0;
        var indices = new int[components];
        var seenComponents = new bool[frame.ComponentCount];
        var offset = 1;
        for (var i = 0; i < components; i++) {
            var id = data[offset++];
            var table = data[offset++];
            var dc = table >> 4;
            var ac = table & 0x0F;
            var index = FindComponentIndex(frame.Components, id);
            if (index < 0) throw new FormatException("Unknown JPEG component in scan.");
            if (seenComponents[index]) throw new FormatException("Duplicate JPEG component in scan.");
            seenComponents[index] = true;
            frame.Components[index].DcTable = (byte)dc;
            frame.Components[index].AcTable = (byte)ac;
            indices[i] = index;
            samplingUnits += frame.Components[index].H * frame.Components[index].V;
            if (components > 1 && samplingUnits > 10)
                throw new FormatException("JPEG scan sampling factors exceed supported limits.");
        }

        var ss = data[offset++];
        var se = data[offset++];
        var ahal = data[offset++];

        return new ScanHeader {
            ComponentIndices = indices,
            Ss = ss,
            Se = se,
            Ah = (byte)(ahal >> 4),
            Al = (byte)(ahal & 0x0F)
        };
    }

    private static int FindComponentIndex(Component[] components, int id) {
        for (var i = 0; i < components.Length; i++) {
            if (components[i].Id == id) return i;
        }
        return -1;
    }

}
