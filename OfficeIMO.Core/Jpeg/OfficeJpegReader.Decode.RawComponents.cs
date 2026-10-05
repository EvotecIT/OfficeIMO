using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Container-owned channels, such as TIFF gray/alpha and CMYK/alpha, have no
    // standalone JPEG color interpretation. Preserve their frame order and samples.
    private static void CopyRawComponents(JpegFrame frame, BaselineComponentState[] states,
        bool highQualityChroma, CancellationToken cancellationToken, byte[] components) {
        int componentCount = frame.ComponentCount, maxH = frame.MaxH, maxV = frame.MaxV;
        for (int y = 0; y < frame.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                for (int channel = 0; channel < componentCount; channel++)
                    components[(y * frame.Width + x) * componentCount + channel] =
                        (byte)SampleComponent(states, channel, x, y, maxH, maxV, 0, highQualityChroma);
            }
        }
    }
}
