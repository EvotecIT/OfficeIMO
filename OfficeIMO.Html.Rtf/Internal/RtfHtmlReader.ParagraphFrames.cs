namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyParagraphFrameAttributes(IElement token) {
            Dictionary<string, string> values = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-paragraph-frame"));
            if (values.Count == 0) {
                return;
            }

            ApplyParagraphFrame(values, "frame", EnsureParagraph().Frame);
        }

        private static void ApplyParagraphFrame(Dictionary<string, string> values, string prefix, RtfParagraphFrame frame) {
            frame.WidthTwips = ReadInt(values, prefix + ".width");
            frame.HeightTwips = ReadInt(values, prefix + ".height");
            frame.HorizontalAnchor = ReadEnum<RtfParagraphFrameHorizontalAnchor>(values, prefix + ".horizontalAnchor");
            frame.VerticalAnchor = ReadEnum<RtfParagraphFrameVerticalAnchor>(values, prefix + ".verticalAnchor");
            frame.HorizontalPosition = ReadEnum<RtfParagraphFrameHorizontalPosition>(values, prefix + ".horizontalPosition");
            frame.HorizontalPositionTwips = ReadInt(values, prefix + ".horizontalTwips");
            frame.VerticalPosition = ReadEnum<RtfParagraphFrameVerticalPosition>(values, prefix + ".verticalPosition");
            frame.VerticalPositionTwips = ReadInt(values, prefix + ".verticalTwips");
            frame.AnchorLocked = ReadBool(values, prefix + ".anchorLocked") == true;
            frame.NoOverlap = ReadBool(values, prefix + ".noOverlap");
            frame.NoWrap = ReadBool(values, prefix + ".noWrap") == true;
            frame.TextWrapDistanceTwips = ReadInt(values, prefix + ".wrapDistance");
            frame.TextWrapDistanceHorizontalTwips = ReadInt(values, prefix + ".wrapDistanceHorizontal");
            frame.TextWrapDistanceVerticalTwips = ReadInt(values, prefix + ".wrapDistanceVertical");
            frame.OverlayText = ReadBool(values, prefix + ".overlayText") == true;
            frame.DropCapLines = ReadInt(values, prefix + ".dropCapLines");
            frame.DropCapKind = ReadEnum<RtfDropCapKind>(values, prefix + ".dropCapKind");
        }
    }
}
