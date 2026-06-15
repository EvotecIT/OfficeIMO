namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyParagraphControlAttributes(IElement token) {
            Dictionary<string, string> values = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-paragraph-controls"));
            if (values.Count == 0) {
                return;
            }

            RtfParagraph paragraph = EnsureParagraph();
            paragraph.KeepWithNext = ReadBool(values, "keepWithNext") == true;
            paragraph.KeepLinesTogether = ReadBool(values, "keepLinesTogether") == true;
            paragraph.SuppressLineNumbers = ReadBool(values, "suppressLineNumbers") == true;
            paragraph.AutoHyphenation = ReadBool(values, "autoHyphenation");
            paragraph.ContextualSpacing = ReadBool(values, "contextualSpacing");
            paragraph.AdjustRightIndent = ReadBool(values, "adjustRightIndent");
            paragraph.SnapToLineGrid = ReadBool(values, "snapToLineGrid");
            paragraph.WidowControl = ReadBool(values, "widowControl");
            paragraph.SpaceBeforeAuto = ReadBool(values, "spaceBeforeAuto");
            paragraph.SpaceAfterAuto = ReadBool(values, "spaceAfterAuto");

            IReadOnlyList<RtfTabStop> tabStops = ReadTabStops(values, "tab");
            if (tabStops.Count > 0) {
                paragraph.ReplaceTabStops(tabStops);
            }
        }
    }
}
