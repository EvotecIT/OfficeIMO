using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private bool TryApplyTableControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "intbl":
                    _currentParagraphIsInTable = true;
                    return true;
                case "trowd":
                    BeginTableRow();
                    return true;
                case "trhdr":
                    if (_currentRow != null) {
                        _currentRow.RepeatHeader = true;
                    }

                    return true;
                case "trkeep":
                    if (_currentRow != null) {
                        _currentRow.KeepTogether = true;
                    }

                    return true;
                case "trkeepfollow":
                    if (_currentRow != null) {
                        _currentRow.KeepWithNext = true;
                    }

                    return true;
                case "trautofit":
                    if (_currentRow != null) {
                        _currentRow.AutoFit = control.Parameter.GetValueOrDefault(1) != 0;
                    }

                    return true;
                case "ltrrow":
                    if (_currentRow != null) {
                        _currentRow.Direction = RtfTableRowDirection.LeftToRight;
                    }

                    return true;
                case "rtlrow":
                    if (_currentRow != null) {
                        _currentRow.Direction = RtfTableRowDirection.RightToLeft;
                    }

                    return true;
                case "trrh":
                    if (_currentRow != null) {
                        _currentRow.HeightTwips = control.Parameter;
                    }

                    return true;
                case "trgaph":
                    if (_currentRow != null) {
                        _currentRow.CellGapTwips = control.Parameter;
                    }

                    return true;
                case "trleft":
                    if (_currentRow != null) {
                        _currentRow.LeftIndentTwips = control.Parameter;
                    }

                    return true;
                case "trwWidth":
                    if (_currentRow != null) {
                        _currentRow.PreferredWidth = control.Parameter;
                    }

                    return true;
                case "trftsWidth":
                    if (_currentRow != null) {
                        _currentRow.PreferredWidthUnit = ToRtfTableWidthUnit(control.Parameter);
                    }

                    return true;
                case "trcbpat":
                    if (_currentRow != null) {
                        _currentRow.BackgroundColorIndex = control.Parameter;
                    }

                    return true;
                case "trcfpat":
                    if (_currentRow != null) {
                        _currentRow.ShadingForegroundColorIndex = control.Parameter;
                    }

                    return true;
                case "trpat":
                    if (_currentRow != null) {
                        _currentRow.ShadingPatternValue = control.Parameter;
                    }

                    return true;
                case "trshdng":
                    if (_currentRow != null) {
                        _currentRow.ShadingPatternPercent = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    }

                    return true;
                case "trbghoriz":
                    ApplyTableRowShadingPattern(RtfShadingPattern.Horizontal);
                    return true;
                case "trbgvert":
                    ApplyTableRowShadingPattern(RtfShadingPattern.Vertical);
                    return true;
                case "trbgfdiag":
                    ApplyTableRowShadingPattern(RtfShadingPattern.ForwardDiagonal);
                    return true;
                case "trbgbdiag":
                    ApplyTableRowShadingPattern(RtfShadingPattern.BackwardDiagonal);
                    return true;
                case "trbgcross":
                    ApplyTableRowShadingPattern(RtfShadingPattern.Cross);
                    return true;
                case "trbgdcross":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DiagonalCross);
                    return true;
                case "trbgdkhor":
                case "trbgdkhoriz":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DarkHorizontal);
                    return true;
                case "trbgdkvert":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DarkVertical);
                    return true;
                case "trbgdkfdiag":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DarkForwardDiagonal);
                    return true;
                case "trbgdkbdiag":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DarkBackwardDiagonal);
                    return true;
                case "trbgdkcross":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DarkCross);
                    return true;
                case "trbgdkdcross":
                    ApplyTableRowShadingPattern(RtfShadingPattern.DarkDiagonalCross);
                    return true;
                case "trpaddt":
                    SetTableRowPadding(RtfBoxSide.Top, control.Parameter);
                    return true;
                case "trpaddl":
                    SetTableRowPadding(RtfBoxSide.Left, control.Parameter);
                    return true;
                case "trpaddb":
                    SetTableRowPadding(RtfBoxSide.Bottom, control.Parameter);
                    return true;
                case "trpaddr":
                    SetTableRowPadding(RtfBoxSide.Right, control.Parameter);
                    return true;
                case "trpaddft":
                    SetTableRowPaddingUnit(RtfBoxSide.Top, control.Parameter);
                    return true;
                case "trpaddfl":
                    SetTableRowPaddingUnit(RtfBoxSide.Left, control.Parameter);
                    return true;
                case "trpaddfb":
                    SetTableRowPaddingUnit(RtfBoxSide.Bottom, control.Parameter);
                    return true;
                case "trpaddfr":
                    SetTableRowPaddingUnit(RtfBoxSide.Right, control.Parameter);
                    return true;
                case "trspdt":
                    SetTableRowSpacing(RtfBoxSide.Top, control.Parameter);
                    return true;
                case "trspdl":
                    SetTableRowSpacing(RtfBoxSide.Left, control.Parameter);
                    return true;
                case "trspdb":
                    SetTableRowSpacing(RtfBoxSide.Bottom, control.Parameter);
                    return true;
                case "trspdr":
                    SetTableRowSpacing(RtfBoxSide.Right, control.Parameter);
                    return true;
                case "trspdft":
                    SetTableRowSpacingUnit(RtfBoxSide.Top, control.Parameter);
                    return true;
                case "trspdfl":
                    SetTableRowSpacingUnit(RtfBoxSide.Left, control.Parameter);
                    return true;
                case "trspdfb":
                    SetTableRowSpacingUnit(RtfBoxSide.Bottom, control.Parameter);
                    return true;
                case "trspdfr":
                    SetTableRowSpacingUnit(RtfBoxSide.Right, control.Parameter);
                    return true;
                case "tabsnoovrlp":
                    if (_currentRow != null) {
                        _currentRow.NoOverlap = true;
                    }

                    return true;
                case "tphcol":
                    SetTableRowHorizontalAnchor(RtfTableHorizontalAnchor.Column);
                    return true;
                case "tphmrg":
                    SetTableRowHorizontalAnchor(RtfTableHorizontalAnchor.Margin);
                    return true;
                case "tphpg":
                    SetTableRowHorizontalAnchor(RtfTableHorizontalAnchor.Page);
                    return true;
                case "tpvmrg":
                    SetTableRowVerticalAnchor(RtfTableVerticalAnchor.Margin);
                    return true;
                case "tpvpara":
                    SetTableRowVerticalAnchor(RtfTableVerticalAnchor.Paragraph);
                    return true;
                case "tpvpg":
                    SetTableRowVerticalAnchor(RtfTableVerticalAnchor.Page);
                    return true;
                case "tposx":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.Absolute, control.Parameter);
                    return true;
                case "tposnegx":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.NegativeAbsolute, control.Parameter);
                    return true;
                case "tposxl":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.Left, null);
                    return true;
                case "tposxc":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.Center, null);
                    return true;
                case "tposxr":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.Right, null);
                    return true;
                case "tposxi":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.Inside, null);
                    return true;
                case "tposxo":
                    SetTableRowHorizontalPosition(RtfTableHorizontalPosition.Outside, null);
                    return true;
                case "tposy":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Absolute, control.Parameter);
                    return true;
                case "tposnegy":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.NegativeAbsolute, control.Parameter);
                    return true;
                case "tposyt":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Top, null);
                    return true;
                case "tposyc":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Center, null);
                    return true;
                case "tposyb":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Bottom, null);
                    return true;
                case "tposyil":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Inline, null);
                    return true;
                case "tposyin":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Inside, null);
                    return true;
                case "tposyoutv":
                    SetTableRowVerticalPosition(RtfTableVerticalPosition.Outside, null);
                    return true;
                case "tdfrmtxtLeft":
                    if (_currentRow != null) {
                        _currentRow.TextWrapLeftTwips = control.Parameter;
                    }

                    return true;
                case "tdfrmtxtRight":
                    if (_currentRow != null) {
                        _currentRow.TextWrapRightTwips = control.Parameter;
                    }

                    return true;
                case "tdfrmtxtTop":
                    if (_currentRow != null) {
                        _currentRow.TextWrapTopTwips = control.Parameter;
                    }

                    return true;
                case "tdfrmtxtBottom":
                    if (_currentRow != null) {
                        _currentRow.TextWrapBottomTwips = control.Parameter;
                    }

                    return true;
                case "trql":
                    if (_currentRow != null) {
                        _currentRow.Alignment = RtfTableAlignment.Left;
                    }

                    return true;
                case "trqc":
                    if (_currentRow != null) {
                        _currentRow.Alignment = RtfTableAlignment.Center;
                    }

                    return true;
                case "trqr":
                    if (_currentRow != null) {
                        _currentRow.Alignment = RtfTableAlignment.Right;
                    }

                    return true;
                case "trbrdrt":
                    BeginTableRowBorder(state, RtfTableRowBorderSide.Top);
                    return true;
                case "trbrdrl":
                    BeginTableRowBorder(state, RtfTableRowBorderSide.Left);
                    return true;
                case "trbrdrb":
                    BeginTableRowBorder(state, RtfTableRowBorderSide.Bottom);
                    return true;
                case "trbrdrr":
                    BeginTableRowBorder(state, RtfTableRowBorderSide.Right);
                    return true;
                case "trbrdrh":
                    BeginTableRowBorder(state, RtfTableRowBorderSide.Horizontal);
                    return true;
                case "trbrdrv":
                    BeginTableRowBorder(state, RtfTableRowBorderSide.Vertical);
                    return true;
                case "clmgf":
                    _pendingCellProperties.HorizontalMerge = RtfTableCellMerge.First;
                    return true;
                case "clmrg":
                    _pendingCellProperties.HorizontalMerge = RtfTableCellMerge.Continue;
                    return true;
                case "clvmgf":
                    _pendingCellProperties.VerticalMerge = RtfTableCellMerge.First;
                    return true;
                case "clvmrg":
                    _pendingCellProperties.VerticalMerge = RtfTableCellMerge.Continue;
                    return true;
                case "clwWidth":
                    _pendingCellProperties.PreferredWidth = control.Parameter;
                    return true;
                case "clftsWidth":
                    _pendingCellProperties.PreferredWidthUnit = ToRtfTableWidthUnit(control.Parameter);
                    return true;
                case "clhidemark":
                    _pendingCellProperties.HideCellMark = true;
                    return true;
                case "clNoWrap":
                    _pendingCellProperties.NoWrap = true;
                    return true;
                case "clFitText":
                    _pendingCellProperties.FitText = true;
                    return true;
                case "clcbpat":
                    _pendingCellProperties.BackgroundColorIndex = control.Parameter;
                    return true;
                case "clcfpat":
                    _pendingCellProperties.ShadingForegroundColorIndex = control.Parameter;
                    return true;
                case "clshdng":
                    _pendingCellProperties.ShadingPatternPercent = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return true;
                case "clbghoriz":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.Horizontal;
                    return true;
                case "clbgvert":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.Vertical;
                    return true;
                case "clbgfdiag":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.ForwardDiagonal;
                    return true;
                case "clbgbdiag":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.BackwardDiagonal;
                    return true;
                case "clbgcross":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.Cross;
                    return true;
                case "clbgdcross":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DiagonalCross;
                    return true;
                case "clbgdkhor":
                case "clbgdkhoriz":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DarkHorizontal;
                    return true;
                case "clbgdkvert":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DarkVertical;
                    return true;
                case "clbgdkfdiag":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DarkForwardDiagonal;
                    return true;
                case "clbgdkbdiag":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DarkBackwardDiagonal;
                    return true;
                case "clbgdkcross":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DarkCross;
                    return true;
                case "clbgdkdcross":
                    _pendingCellProperties.ShadingPattern = RtfShadingPattern.DarkDiagonalCross;
                    return true;
                case "clpadt":
                    _pendingCellProperties.PaddingTopTwips = control.Parameter;
                    return true;
                case "clpadl":
                    _pendingCellProperties.PaddingLeftTwips = control.Parameter;
                    return true;
                case "clpadb":
                    _pendingCellProperties.PaddingBottomTwips = control.Parameter;
                    return true;
                case "clpadr":
                    _pendingCellProperties.PaddingRightTwips = control.Parameter;
                    return true;
                case "clpadft":
                    _pendingCellProperties.PaddingTopUnit = control.Parameter;
                    return true;
                case "clpadfl":
                    _pendingCellProperties.PaddingLeftUnit = control.Parameter;
                    return true;
                case "clpadfb":
                    _pendingCellProperties.PaddingBottomUnit = control.Parameter;
                    return true;
                case "clpadfr":
                    _pendingCellProperties.PaddingRightUnit = control.Parameter;
                    return true;
                case "clvertalt":
                    _pendingCellProperties.VerticalAlignment = RtfTableCellVerticalAlignment.Top;
                    return true;
                case "clvertalc":
                    _pendingCellProperties.VerticalAlignment = RtfTableCellVerticalAlignment.Center;
                    return true;
                case "clvertalb":
                    _pendingCellProperties.VerticalAlignment = RtfTableCellVerticalAlignment.Bottom;
                    return true;
                case "cltxlrtb":
                    _pendingCellProperties.TextFlow = RtfTableCellTextFlow.LeftToRightTopToBottom;
                    return true;
                case "cltxtbrl":
                    _pendingCellProperties.TextFlow = RtfTableCellTextFlow.TopToBottomRightToLeft;
                    return true;
                case "cltxbtlr":
                    _pendingCellProperties.TextFlow = RtfTableCellTextFlow.BottomToTopLeftToRight;
                    return true;
                case "cltxlrtbv":
                    _pendingCellProperties.TextFlow = RtfTableCellTextFlow.LeftToRightTopToBottomVertical;
                    return true;
                case "cltxtbrlv":
                    _pendingCellProperties.TextFlow = RtfTableCellTextFlow.TopToBottomRightToLeftVertical;
                    return true;
                case "clbrdrt":
                    state.CurrentParagraphBorderSide = null;
                    BeginPendingCellBorder(RtfTableCellBorderSide.Top);
                    return true;
                case "clbrdrl":
                    state.CurrentParagraphBorderSide = null;
                    BeginPendingCellBorder(RtfTableCellBorderSide.Left);
                    return true;
                case "clbrdrb":
                    state.CurrentParagraphBorderSide = null;
                    BeginPendingCellBorder(RtfTableCellBorderSide.Bottom);
                    return true;
                case "clbrdrr":
                    state.CurrentParagraphBorderSide = null;
                    BeginPendingCellBorder(RtfTableCellBorderSide.Right);
                    return true;
                case "cldglu":
                    state.CurrentParagraphBorderSide = null;
                    BeginPendingCellBorder(RtfTableCellBorderSide.TopLeftToBottomRight);
                    return true;
                case "cldgll":
                    state.CurrentParagraphBorderSide = null;
                    BeginPendingCellBorder(RtfTableCellBorderSide.TopRightToBottomLeft);
                    return true;
                case "cellx":
                    AddTableCellBoundary(control.Parameter);
                    return true;
                case "cell":
                    FlushTableCell(state);
                    return true;
                case "row":
                    EndTableRow(state);
                    return true;
                default:
                    return false;
            }
        }
    }
}