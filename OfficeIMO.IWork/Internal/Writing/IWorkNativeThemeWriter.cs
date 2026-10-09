using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Emits the owned 30-slot color palette addressed by native iWork style controls.</summary>
internal static class IWorkNativeThemeWriter {
    internal static void AddColorPresets(IWorkProtoWriter theme, int maximumBytes, CancellationToken cancellationToken) {
        IWorkProtoWriter Color(float r, float g, float b, float alpha) => new IWorkProtoWriter(maximumBytes, cancellationToken)
            .UInt(1, 1).Float(3, r).Float(4, g).Float(5, b).Float(6, alpha).UInt(12, 1);
        foreach (float grey in new[] { 1f, .8f, .6f, .3f, 0f }) theme.Message(10, Color(grey, grey, grey, 1));
        theme.Message(10, Color(0, 0, 0, 0));
        foreach (var rgb in new[] { (0f, .5f, 1f), (0f, .8f, .5f), (.5f, .8f, 0f),
                     (1f, .8f, 0f), (1f, .2f, .2f), (.6f, .2f, 1f) }) {
            foreach (float scale in new[] { 1f, .8f, .6f, .4f })
                theme.Message(10, Color(rgb.Item1 * scale, rgb.Item2 * scale, rgb.Item3 * scale, 1));
        }
    }
}
