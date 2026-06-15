namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteParagraphFrame(StringBuilder builder, RtfParagraphFrame frame) {
        if (!frame.HasAnyValue) return;

        AppendOptionalTwips(builder, @"\absw", frame.WidthTwips);
        AppendOptionalTwips(builder, @"\absh", frame.HeightTwips);
        if (frame.HorizontalAnchor.HasValue) {
            builder.Append(frame.HorizontalAnchor.Value switch {
                RtfParagraphFrameHorizontalAnchor.Margin => @"\phmrg",
                RtfParagraphFrameHorizontalAnchor.Page => @"\phpg",
                _ => @"\phcol"
            });
        }

        WriteParagraphFrameHorizontalPosition(builder, frame.HorizontalPosition, frame.HorizontalPositionTwips);
        if (frame.VerticalAnchor.HasValue) {
            builder.Append(frame.VerticalAnchor.Value switch {
                RtfParagraphFrameVerticalAnchor.Paragraph => @"\pvpara",
                RtfParagraphFrameVerticalAnchor.Page => @"\pvpg",
                _ => @"\pvmrg"
            });
        }

        WriteParagraphFrameVerticalPosition(builder, frame.VerticalPosition, frame.VerticalPositionTwips);
        if (frame.AnchorLocked) {
            builder.Append(@"\abslock");
        }

        AppendOptionalBinary(builder, @"\absnoovrlp", frame.NoOverlap);
        if (frame.NoWrap) {
            builder.Append(@"\nowrap");
        }

        AppendOptionalTwips(builder, @"\dxfrtext", frame.TextWrapDistanceTwips);
        AppendOptionalTwips(builder, @"\dfrmtxtx", frame.TextWrapDistanceHorizontalTwips);
        AppendOptionalTwips(builder, @"\dfrmtxty", frame.TextWrapDistanceVerticalTwips);
        if (frame.OverlayText) {
            builder.Append(@"\overlay");
        }

        AppendOptionalTwips(builder, @"\dropcapli", frame.DropCapLines);
        if (frame.DropCapKind.HasValue) {
            builder.Append(@"\dropcapt");
            builder.Append(frame.DropCapKind.Value == RtfDropCapKind.Margin ? "2" : "1");
        }
    }

    private static void WriteParagraphFrameHorizontalPosition(StringBuilder builder, RtfParagraphFrameHorizontalPosition? position, int? twips) {
        if (!position.HasValue) return;

        switch (position.Value) {
            case RtfParagraphFrameHorizontalPosition.Absolute:
                AppendOptionalTwips(builder, @"\posx", twips);
                return;
            case RtfParagraphFrameHorizontalPosition.NegativeAbsolute:
                AppendOptionalTwips(builder, @"\posnegx", twips);
                return;
            case RtfParagraphFrameHorizontalPosition.Center:
                builder.Append(@"\posxc");
                return;
            case RtfParagraphFrameHorizontalPosition.Right:
                builder.Append(@"\posxr");
                return;
            case RtfParagraphFrameHorizontalPosition.Inside:
                builder.Append(@"\posxi");
                return;
            case RtfParagraphFrameHorizontalPosition.Outside:
                builder.Append(@"\posxo");
                return;
            default:
                builder.Append(@"\posxl");
                return;
        }
    }

    private static void WriteParagraphFrameVerticalPosition(StringBuilder builder, RtfParagraphFrameVerticalPosition? position, int? twips) {
        if (!position.HasValue) return;

        switch (position.Value) {
            case RtfParagraphFrameVerticalPosition.Absolute:
                AppendOptionalTwips(builder, @"\posy", twips);
                return;
            case RtfParagraphFrameVerticalPosition.NegativeAbsolute:
                AppendOptionalTwips(builder, @"\posnegy", twips);
                return;
            case RtfParagraphFrameVerticalPosition.Center:
                builder.Append(@"\posyc");
                return;
            case RtfParagraphFrameVerticalPosition.Bottom:
                builder.Append(@"\posyb");
                return;
            case RtfParagraphFrameVerticalPosition.Inline:
                builder.Append(@"\posyil");
                return;
            case RtfParagraphFrameVerticalPosition.Inside:
                builder.Append(@"\posyin");
                return;
            case RtfParagraphFrameVerticalPosition.Outside:
                builder.Append(@"\posyout");
                return;
            default:
                builder.Append(@"\posyt");
                return;
        }
    }
}
