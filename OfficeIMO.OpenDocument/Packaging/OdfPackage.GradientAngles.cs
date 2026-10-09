namespace OfficeIMO.OpenDocument;

internal sealed partial class OdfPackage {
    // Old producers used tenths of degrees in unitless native gradient angles.
    // A version-only upgrade can therefore change the color field without changing its XML.
    private void EnsureGradientAnglesSurviveVersionChange(OdfVersion outputVersion) {
        if (Version != OdfVersion.V1_2 || outputVersion == Version) return;
        foreach (string part in new[] { "styles.xml", "content.xml" }) {
            if (!ContainsEntry(part)) continue;
            foreach (XElement definition in GetXml(part).Descendants().Where(element =>
                element.Name == OdfNamespaces.Draw + "gradient" || element.Name == OdfNamespaces.Draw + "opacity")) {
                if ((string?)definition.Attribute(OdfNamespaces.Draw + "style") == "radial") continue;
                string? angle = (string?)definition.Attribute(OdfNamespaces.Draw + "angle");
                if (angle != null && double.TryParse(angle, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) && value != 0)
                    throw new InvalidOperationException("Changing the ODF version of a native gradient or opacity definition with a nonzero unitless angle can change its appearance. Use PreserveSource or replace the definition with an explicit angle unit before saving.");
            }
        }
    }
}
