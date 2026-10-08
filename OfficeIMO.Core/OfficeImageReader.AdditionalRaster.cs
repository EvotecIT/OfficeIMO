using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private static bool TryReadAdditionalRaster(byte[] bytes, CancellationToken token, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        token.ThrowIfCancellationRequested();
        if (bytes.Length >= 3 && bytes[0] == 'P' && bytes[1] >= '1' && bytes[1] <= '6') {
            if (!OfficePortableMapCodec.TryIdentify(bytes, out int width, out int height)) return false;
            info = new OfficeImageInfo(OfficeImageFormat.PortableMap, width, height); return true;
        }
        if (OfficeTgaCodec.TryIdentify(bytes, out int tgaWidth, out int tgaHeight)) { info = new OfficeImageInfo(OfficeImageFormat.Tga, tgaWidth, tgaHeight); return true; }
        return false;
    }
}
