using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private void NativeRadialPad(XElement gradient, XElement target, string attribute, BrushRegion region, bool outsideFromLastStop = false) {
        if (region.Width <= 0 || region.Height <= 0) { Set(target, attribute, "none"); return; }
        string id = (string)gradient.Attribute("id")!;
        string cx = (string)gradient.Attribute("cx")!, cy = (string)gradient.Attribute("cy")!, radius = (string)gradient.Attribute("r")!;
        Set(gradient, "cx", (string)gradient.Attribute("fx")!); Set(gradient, "cy", (string)gradient.Attribute("fy")!);
        Set(gradient, "fx", cx); Set(gradient, "fy", cy); Set(gradient, "fr", radius); Set(gradient, "r", "0");
        var stops = gradient.Elements().ToArray();
        if (stops.Length == 0) { Set(target, attribute, "none"); return; }
        // Match SVG's monotonic clamping before reversal, including duplicate stops.
        double previous = 0;
        foreach (var stop in stops) {
            previous = Math.Max(previous, Math.Min(1, Math.Max(0, XpsPackage.Number((string?)stop.Attribute("offset")))));
            Set(stop, "offset", N(1 - previous));
        }
        gradient.RemoveNodes();
        foreach (var stop in Enumerable.Reverse(stops)) gradient.Add(stop);
        var alpha = CloneProjection(gradient);
        Set(alpha, "color-interpolation", "sRGB");
        foreach (var stop in alpha.Elements()) {
            double opacity = Unit((string?)stop.Attribute("stop-opacity") ?? "1");
            string channel = N(opacity * 100) + "%";
            Set(stop, "stop-color", "rgb(" + channel + "," + channel + "," + channel + ")");
            stop.Attribute("stop-opacity")?.Remove();
        }
        foreach (var stop in gradient.Elements()) stop.Attribute("stop-opacity")?.Remove();
        // Native Reflect uses the original first stop outside the cone; after
        // reversal this is the last stop. Repeat and Pad use the original last.
        var outsideColor = outsideFromLastStop ? gradient.Elements().Last() : gradient.Elements().First();
        var outsideAlpha = outsideFromLastStop ? alpha.Elements().Last() : alpha.Elements().First();
        var markup = new StringBuilder();
        markup.AppendNativeRadialPatternDefinition(id, region.X, region.Y, region.Width, region.Height,
            (string)outsideColor.Attribute("stop-color")!, (string)outsideAlpha.Attribute("stop-color")!,
            (output, fieldId) => { Set(gradient, "id", fieldId); output.Append(gradient.ToString(SaveOptions.DisableFormatting)); },
            (output, fieldId) => { Set(alpha, "id", fieldId); output.Append(alpha.ToString(SaveOptions.DisableFormatting)); });
        EnsureOutputCapacity(markup.Length);
        // Charge every composed node/attribute through the native converter's budget.
        _defs.Add(CloneProjection(XElement.Parse(markup.ToString())).Elements().ToArray());
        Set(target, attribute, "url(#" + id + ")");
    }
}
