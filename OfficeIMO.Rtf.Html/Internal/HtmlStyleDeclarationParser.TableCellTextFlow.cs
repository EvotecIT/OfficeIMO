namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    private static RtfTableCellTextFlow? ParseWritingMode(string value) {
        switch (value) {
            case "horizontal-tb":
            case "lr-tb":
                return RtfTableCellTextFlow.LeftToRightTopToBottom;
            case "vertical-rl":
            case "tb-rl":
            case "sideways-rl":
                return RtfTableCellTextFlow.TopToBottomRightToLeft;
            case "sideways-lr":
            case "bt-lr":
                return RtfTableCellTextFlow.BottomToTopLeftToRight;
            case "vertical-lr":
                return RtfTableCellTextFlow.LeftToRightTopToBottomVertical;
            default:
                return null;
        }
    }

    private static RtfTableCellTextFlow? ParseRtfTableCellTextFlow(string value) {
        switch (value) {
            case "ltr-tb":
            case "left-to-right-top-to-bottom":
            case "cltxlrtb":
                return RtfTableCellTextFlow.LeftToRightTopToBottom;
            case "tb-rl":
            case "top-to-bottom-right-to-left":
            case "cltxtbrl":
                return RtfTableCellTextFlow.TopToBottomRightToLeft;
            case "bt-lr":
            case "bottom-to-top-left-to-right":
            case "cltxbtlr":
                return RtfTableCellTextFlow.BottomToTopLeftToRight;
            case "ltr-tb-v":
            case "left-to-right-top-to-bottom-vertical":
            case "cltxlrtbv":
                return RtfTableCellTextFlow.LeftToRightTopToBottomVertical;
            case "tb-rl-v":
            case "top-to-bottom-right-to-left-vertical":
            case "cltxtbrlv":
                return RtfTableCellTextFlow.TopToBottomRightToLeftVertical;
            default:
                return null;
        }
    }
}
