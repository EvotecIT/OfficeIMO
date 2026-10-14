using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Area anchoring and paragraph alignment are independent. Keep font measurement in Core.
    private static OfficeTextAreaAlignment ProjectTextArea(OdgShape shape, string? area, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        OdfConversionReport report, string feature, HashSet<string> losses, bool shrinkToFit, ref bool wrapText) {
        if (area is null or "justify") return OfficeTextAreaAlignment.FullWidth;
        if (area is not ("left" or "center" or "right")) {
            losses.Add("text-area-alignment");
            return OfficeTextAreaAlignment.FullWidth;
        }
        bool rectangle = shape.ElementName == "rect";
        bool textBox = shape.ElementName == "frame" && shape.TextRoot.Name == OdfNamespaces.Draw + "text-box";
        if (paragraphs.Any(p => p.TabStops != null && !p.Indent.IsEmpty &&
            p.Alignment is OfficeTextAlignment.Center or OfficeTextAlignment.Right)) losses.Add("text-area-tab-indent");
        if (!rectangle && !textBox) losses.Add("text-area-alignment");
        else report.Add(feature + ":text-area-alignment", OdfConversionMappingStatus.Approximated,
            message: "The intrinsic paragraph block is anchored inside the padded shape frame independently of paragraph alignment. Shared render-time fonts determine its width; native font metrics and auto-sizing can differ.");
        // Native shrinking rectangle labels honor wrapping; the ordinary intrinsic rectangle profile does not.
        if (rectangle && wrapText && !shrinkToFit) {
            wrapText = false;
            report.Add(feature + ":text-area-wrapping", OdfConversionMappingStatus.Approximated,
                message: "The qualified native rectangle profile leaves an intrinsic text area unwrapped despite a declared wrapping option. Hard line breaks are retained; text-box frames honor wrapping.");
        }
        return area switch {
            "left" => OfficeTextAreaAlignment.Left,
            "center" => OfficeTextAreaAlignment.Center,
            _ => OfficeTextAreaAlignment.Right
        };
    }
}
