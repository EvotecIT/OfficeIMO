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
                        (byte)SampleComponent(states, channel, x, y, maxH, maxV, 0, highQualityChroma, imageWidth: frame.Width, imageHeight: frame.Height);
            }
        }
    }
    // TIFF keeps native sample precision through alpha unassociation and ICC color
    // conversion. This internal path writes words in the container's byte order.
    private static byte[] CopyRawSamples16(JpegFrame frame, BaselineState state,
        bool littleEndian, bool highQualityChroma, CancellationToken token) {
        foreach (bool decoded in state.DecodedComponents)
            if (!decoded) throw new System.FormatException("Missing JPEG component scan.");
        byte[] output = new byte[checked(frame.Width * frame.Height * frame.ComponentCount * 2)];
        int at = 0;
        for (int y = 0; y < frame.Height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                if ((x & 4095) == 0) token.ThrowIfCancellationRequested();
                for (int channel = 0; channel < frame.ComponentCount; channel++) {
                    int sample = SampleComponent(state.Components, channel, x, y, frame.MaxH, frame.MaxV, 0,
                        highQualityChroma, preserveRaw16: true, imageWidth: frame.Width, imageHeight: frame.Height);
                    output[at++] = (byte)(littleEndian ? sample : sample >> 8);
                    output[at++] = (byte)(littleEndian ? sample >> 8 : sample);
                }
            }
        }
        return output;
    }

}
