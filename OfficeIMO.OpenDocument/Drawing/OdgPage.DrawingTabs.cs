using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static OfficeTextTabStops ProjectTabStops(OdfTextParagraph paragraph, double left, double paragraphLeft, HashSet<string> losses,
        OdfConversionReport report, string feature) {
        XElement? defaults = paragraph.Document.Styles.FindDefaultProperties(OdfStyleFamily.Paragraph, OdfNamespaces.Style + "paragraph-properties");
        string? nativeDistance = (string?)defaults?.Attribute(OdfNamespaces.Style + "tab-stop-distance");
        double distance = 36;
        if (nativeDistance == null) report.Add(feature + ":implicit-tab-distance", OdfConversionMappingStatus.Approximated,
            message: "The document omits default tab spacing; the shared layout uses 36 points. Producer defaults can differ.");
        else {
            try { distance = TextLength(OdfLength.Parse(nativeDistance)); if (distance <= 0) throw new NotSupportedException(); }
            catch (Exception e) when (e is ArgumentException or FormatException or NotSupportedException) { losses.Add("tab-distance"); distance = 36; }
        }
        string? relative = (string?)defaults?.Attribute(OdfNamespaces.Text + "relative-tab-stop-position");
        if (relative is not (null or "true" or "false" or "0" or "1")) losses.Add("tab-origin");
        // A generated list-body indent does not move paragraph-declared stops. Express the
        // original paragraph origin relative to the shared frame's new body margin.
        double origin = relative is "false" or "0" ? -left : paragraphLeft - left;
        var stops = new List<OfficeTextTabStop>();
        OdfStyle? tabStyle = paragraph.Styles.FirstOrDefault(s => s.ParagraphProperties?.Element(OdfNamespaces.Style + "tab-stops") != null);
        XElement? container = tabStyle?.ParagraphProperties?.Element(OdfNamespaces.Style + "tab-stops");
        if (container != null) {
            int visited = 0;
            foreach (XElement element in container.Elements()) {
                if (++visited > 256) { losses.Add("tab-stop-limit"); break; }
                try {
                    if (element.Name != OdfNamespaces.Style + "tab-stop") throw new NotSupportedException();
                    OfficeTextTabAlignment type = ((string?)element.Attribute(OdfNamespaces.Style + "type")) switch {
                        null or "left" => OfficeTextTabAlignment.Left, "center" => OfficeTextTabAlignment.Center,
                        "right" => OfficeTextTabAlignment.Right, "char" => OfficeTextTabAlignment.Character,
                        _ => throw new NotSupportedException()
                    };
                    string? position = (string?)element.Attribute(OdfNamespaces.Style + "position");
                    if (position == null) throw new NotSupportedException();
                    string character = (string?)element.Attribute(OdfNamespaces.Style + "char") ?? (type == OfficeTextTabAlignment.Character ? throw new NotSupportedException() : ".");
                    var stop = new OfficeTextTabStop(TextLength(OdfLength.Parse(position)), type, character);
                    if (stops.Any(s => s.Position == stop.Position)) throw new NotSupportedException();
                    stops.Add(ProjectTabLeader(element, stop, paragraph, tabStyle!, losses));
                } catch (Exception e) when (e is ArgumentException or FormatException or NotSupportedException or OverflowException) { losses.Add("tab-stops"); }
            }
        }
        int leaders = stops.Count(stop => !string.IsNullOrWhiteSpace(stop.LeaderText));
        if (leaders > 0) report.Add(feature + ":tab-leaders", OdfConversionMappingStatus.Approximated, leaders,
            "Textual leaders use the measured gap, active formatting and referenced text-style overrides; producer glyph phase and typography can differ. The tested native Draw exporter ignores separate leader styles.");
        int lineLeaders = stops.Count(stop => stop.LeaderText == null && stop.LineLeader?.Style != null && stop.LineLeader.Style != OfficeTextTabLineLeaderStyle.None);
        if (lineLeaders > 0) report.Add(feature + ":tab-line-leaders", OdfConversionMappingStatus.Approximated, lineLeaders,
            "Line patterns use shared bounded outlines, declared colors and absolute widths. Named, integer and percentage widths use the documented font-relative approximation profile; producer rendering and native resave can differ.");
        return new OfficeTextTabStops(stops, distance, origin).WithParagraphAlignment();
    }
}
