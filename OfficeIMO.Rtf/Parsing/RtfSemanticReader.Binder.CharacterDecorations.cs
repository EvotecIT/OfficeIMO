using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private bool TryApplyCharacterDecorationControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "chcbpat":
                    state.CharacterBackgroundColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return true;
                case "chcfpat":
                    state.CharacterShadingForegroundColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return true;
                case "chshdng":
                    state.CharacterShadingPatternPercent = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return true;
                case "chbghoriz":
                    state.CharacterShadingPattern = RtfShadingPattern.Horizontal;
                    return true;
                case "chbgvert":
                    state.CharacterShadingPattern = RtfShadingPattern.Vertical;
                    return true;
                case "chbgfdiag":
                    state.CharacterShadingPattern = RtfShadingPattern.ForwardDiagonal;
                    return true;
                case "chbgbdiag":
                    state.CharacterShadingPattern = RtfShadingPattern.BackwardDiagonal;
                    return true;
                case "chbgcross":
                    state.CharacterShadingPattern = RtfShadingPattern.Cross;
                    return true;
                case "chbgdcross":
                    state.CharacterShadingPattern = RtfShadingPattern.DiagonalCross;
                    return true;
                case "chbgdkhoriz":
                    state.CharacterShadingPattern = RtfShadingPattern.DarkHorizontal;
                    return true;
                case "chbgdkvert":
                    state.CharacterShadingPattern = RtfShadingPattern.DarkVertical;
                    return true;
                case "chbgdkfdiag":
                    state.CharacterShadingPattern = RtfShadingPattern.DarkForwardDiagonal;
                    return true;
                case "chbgdkbdiag":
                    state.CharacterShadingPattern = RtfShadingPattern.DarkBackwardDiagonal;
                    return true;
                case "chbgdkcross":
                    state.CharacterShadingPattern = RtfShadingPattern.DarkCross;
                    return true;
                case "chbgdkdcross":
                    state.CharacterShadingPattern = RtfShadingPattern.DarkDiagonalCross;
                    return true;
                case "chbrdr":
                    BeginCharacterBorder(state);
                    return true;
                default:
                    return false;
            }
        }

        private void BeginCharacterBorder(CharacterState state) {
            state.CurrentCharacterBorderActive = true;
            state.CurrentParagraphBorderSide = null;
            state.CurrentPageBorderSide = null;
            _currentRowBorderSide = null;
            _pendingCellProperties.CurrentBorderSide = null;
            state.CharacterBorder.Style = RtfParagraphBorderStyle.Single;
        }
    }
}
