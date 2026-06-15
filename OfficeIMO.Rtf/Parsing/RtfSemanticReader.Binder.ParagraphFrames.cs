using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static bool TryApplyParagraphFrameControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "absw":
                    state.Frame.WidthTwips = control.Parameter;
                    return true;
                case "absh":
                    state.Frame.HeightTwips = control.Parameter;
                    return true;
                case "phcol":
                    state.Frame.HorizontalAnchor = RtfParagraphFrameHorizontalAnchor.Column;
                    return true;
                case "phmrg":
                    state.Frame.HorizontalAnchor = RtfParagraphFrameHorizontalAnchor.Margin;
                    return true;
                case "phpg":
                    state.Frame.HorizontalAnchor = RtfParagraphFrameHorizontalAnchor.Page;
                    return true;
                case "posx":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.Absolute, control.Parameter);
                    return true;
                case "posnegx":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.NegativeAbsolute, control.Parameter);
                    return true;
                case "posxl":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.Left, null);
                    return true;
                case "posxc":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.Center, null);
                    return true;
                case "posxr":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.Right, null);
                    return true;
                case "posxi":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.Inside, null);
                    return true;
                case "posxo":
                    SetParagraphFrameHorizontalPosition(state, RtfParagraphFrameHorizontalPosition.Outside, null);
                    return true;
                case "pvmrg":
                    state.Frame.VerticalAnchor = RtfParagraphFrameVerticalAnchor.Margin;
                    return true;
                case "pvpara":
                    state.Frame.VerticalAnchor = RtfParagraphFrameVerticalAnchor.Paragraph;
                    return true;
                case "pvpg":
                    state.Frame.VerticalAnchor = RtfParagraphFrameVerticalAnchor.Page;
                    return true;
                case "posy":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Absolute, control.Parameter);
                    return true;
                case "posnegy":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.NegativeAbsolute, control.Parameter);
                    return true;
                case "posyt":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Top, null);
                    return true;
                case "posyc":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Center, null);
                    return true;
                case "posyb":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Bottom, null);
                    return true;
                case "posyil":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Inline, null);
                    return true;
                case "posyin":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Inside, null);
                    return true;
                case "posyout":
                    SetParagraphFrameVerticalPosition(state, RtfParagraphFrameVerticalPosition.Outside, null);
                    return true;
                case "abslock":
                    state.Frame.AnchorLocked = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "absnoovrlp":
                    state.Frame.NoOverlap = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "nowrap":
                    state.Frame.NoWrap = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "dxfrtext":
                    state.Frame.TextWrapDistanceTwips = control.Parameter;
                    return true;
                case "dfrmtxtx":
                    state.Frame.TextWrapDistanceHorizontalTwips = control.Parameter;
                    return true;
                case "dfrmtxty":
                    state.Frame.TextWrapDistanceVerticalTwips = control.Parameter;
                    return true;
                case "overlay":
                    state.Frame.OverlayText = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "dropcapli":
                    state.Frame.DropCapLines = control.Parameter;
                    return true;
                case "dropcapt":
                    state.Frame.DropCapKind = control.Parameter.GetValueOrDefault() == 2 ? RtfDropCapKind.Margin : RtfDropCapKind.InText;
                    return true;
                default:
                    return false;
            }
        }

        private static void SetParagraphFrameHorizontalPosition(CharacterState state, RtfParagraphFrameHorizontalPosition position, int? twips) {
            state.Frame.HorizontalPosition = position;
            state.Frame.HorizontalPositionTwips = twips;
        }

        private static void SetParagraphFrameVerticalPosition(CharacterState state, RtfParagraphFrameVerticalPosition position, int? twips) {
            state.Frame.VerticalPosition = position;
            state.Frame.VerticalPositionTwips = twips;
        }
    }
}
