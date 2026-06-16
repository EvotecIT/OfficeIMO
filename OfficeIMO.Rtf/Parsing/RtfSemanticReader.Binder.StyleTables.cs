using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static bool TryApplyStyleTableControl(
            RtfControlWord control,
            RtfTableRow tableRowFormat,
            ref PendingTableCellProperties pendingTableCell,
            RowBoxMeasurements rowPadding,
            RowBoxMeasurements rowSpacing,
            ref RtfTableRowBorderSide? currentTableRowBorderSide,
            ref RtfStyleKind kind) {
            switch (control.Name) {
                case "tsrowd":
                case "trowd":
                    kind = RtfStyleKind.Table;
                    currentTableRowBorderSide = null;
                    pendingTableCell.CurrentBorderSide = null;
                    return true;
                case "brdrs":
                    return ApplyStyleTableBorderStyle(tableRowFormat, pendingTableCell, currentTableRowBorderSide, RtfTableCellBorderStyle.Single);
                case "brdrdb":
                    return ApplyStyleTableBorderStyle(tableRowFormat, pendingTableCell, currentTableRowBorderSide, RtfTableCellBorderStyle.Double);
                case "brdrdot":
                    return ApplyStyleTableBorderStyle(tableRowFormat, pendingTableCell, currentTableRowBorderSide, RtfTableCellBorderStyle.Dotted);
                case "brdrdash":
                    return ApplyStyleTableBorderStyle(tableRowFormat, pendingTableCell, currentTableRowBorderSide, RtfTableCellBorderStyle.Dashed);
                case "brdrnil":
                case "brdrnone":
                    return ApplyStyleTableBorderStyle(tableRowFormat, pendingTableCell, currentTableRowBorderSide, RtfTableCellBorderStyle.None);
                case "brdrw":
                    return ApplyStyleTableBorderWidth(tableRowFormat, pendingTableCell, currentTableRowBorderSide, control.Parameter);
                case "brdrcf":
                    return ApplyStyleTableBorderColor(tableRowFormat, pendingTableCell, currentTableRowBorderSide, control.Parameter);
                case "trhdr":
                    tableRowFormat.RepeatHeader = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trkeep":
                    tableRowFormat.KeepTogether = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trkeepfollow":
                    tableRowFormat.KeepWithNext = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trautofit":
                    tableRowFormat.AutoFit = control.Parameter.GetValueOrDefault(1) != 0;
                    kind = RtfStyleKind.Table;
                    return true;
                case "ltrrow":
                    tableRowFormat.Direction = RtfTableRowDirection.LeftToRight;
                    kind = RtfStyleKind.Table;
                    return true;
                case "rtlrow":
                    tableRowFormat.Direction = RtfTableRowDirection.RightToLeft;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trrh":
                    tableRowFormat.HeightTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trgaph":
                    tableRowFormat.CellGapTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trleft":
                    tableRowFormat.LeftIndentTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trwWidth":
                    tableRowFormat.PreferredWidth = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trftsWidth":
                    tableRowFormat.PreferredWidthUnit = ToRtfTableWidthUnit(control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trcbpat":
                    tableRowFormat.BackgroundColorIndex = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trcfpat":
                    tableRowFormat.ShadingForegroundColorIndex = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpat":
                    tableRowFormat.ShadingPatternValue = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trshdng":
                    tableRowFormat.ShadingPatternPercent = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbghoriz":
                case "trbgvert":
                case "trbgfdiag":
                case "trbgbdiag":
                case "trbgcross":
                case "trbgdcross":
                case "trbgdkhor":
                case "trbgdkhoriz":
                case "trbgdkvert":
                case "trbgdkfdiag":
                case "trbgdkbdiag":
                case "trbgdkcross":
                case "trbgdkdcross":
                    tableRowFormat.ShadingPattern = ReadTableRowShadingPattern(control.Name);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddt":
                    SetStyleTableRowBoxValue(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Top, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddl":
                    SetStyleTableRowBoxValue(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Left, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddb":
                    SetStyleTableRowBoxValue(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Bottom, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddr":
                    SetStyleTableRowBoxValue(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Right, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddft":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Top, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddfl":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Left, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddfb":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Bottom, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trpaddfr":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowPadding, isPadding: true, RtfBoxSide.Right, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdt":
                    SetStyleTableRowBoxValue(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Top, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdl":
                    SetStyleTableRowBoxValue(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Left, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdb":
                    SetStyleTableRowBoxValue(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Bottom, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdr":
                    SetStyleTableRowBoxValue(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Right, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdft":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Top, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdfl":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Left, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdfb":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Bottom, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trspdfr":
                    SetStyleTableRowBoxUnit(tableRowFormat, rowSpacing, isPadding: false, RtfBoxSide.Right, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tabsnoovrlp":
                    tableRowFormat.NoOverlap = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tphcol":
                    tableRowFormat.HorizontalAnchor = RtfTableHorizontalAnchor.Column;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tphmrg":
                    tableRowFormat.HorizontalAnchor = RtfTableHorizontalAnchor.Margin;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tphpg":
                    tableRowFormat.HorizontalAnchor = RtfTableHorizontalAnchor.Page;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tpvmrg":
                    tableRowFormat.VerticalAnchor = RtfTableVerticalAnchor.Margin;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tpvpara":
                    tableRowFormat.VerticalAnchor = RtfTableVerticalAnchor.Paragraph;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tpvpg":
                    tableRowFormat.VerticalAnchor = RtfTableVerticalAnchor.Page;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposx":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.Absolute, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposnegx":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.NegativeAbsolute, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposxl":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.Left, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposxc":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.Center, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposxr":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.Right, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposxi":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.Inside, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposxo":
                    SetStyleTableRowHorizontalPosition(tableRowFormat, RtfTableHorizontalPosition.Outside, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposy":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Absolute, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposnegy":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.NegativeAbsolute, control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposyt":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Top, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposyc":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Center, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposyb":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Bottom, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposyil":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Inline, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposyin":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Inside, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tposyoutv":
                    SetStyleTableRowVerticalPosition(tableRowFormat, RtfTableVerticalPosition.Outside, null);
                    kind = RtfStyleKind.Table;
                    return true;
                case "tdfrmtxtLeft":
                    tableRowFormat.TextWrapLeftTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tdfrmtxtRight":
                    tableRowFormat.TextWrapRightTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tdfrmtxtTop":
                    tableRowFormat.TextWrapTopTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "tdfrmtxtBottom":
                    tableRowFormat.TextWrapBottomTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trql":
                    tableRowFormat.Alignment = RtfTableAlignment.Left;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trqc":
                    tableRowFormat.Alignment = RtfTableAlignment.Center;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trqr":
                    tableRowFormat.Alignment = RtfTableAlignment.Right;
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbrdrt":
                    BeginStyleTableRowBorder(tableRowFormat, pendingTableCell, ref currentTableRowBorderSide, RtfTableRowBorderSide.Top);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbrdrl":
                    BeginStyleTableRowBorder(tableRowFormat, pendingTableCell, ref currentTableRowBorderSide, RtfTableRowBorderSide.Left);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbrdrb":
                    BeginStyleTableRowBorder(tableRowFormat, pendingTableCell, ref currentTableRowBorderSide, RtfTableRowBorderSide.Bottom);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbrdrr":
                    BeginStyleTableRowBorder(tableRowFormat, pendingTableCell, ref currentTableRowBorderSide, RtfTableRowBorderSide.Right);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbrdrh":
                    BeginStyleTableRowBorder(tableRowFormat, pendingTableCell, ref currentTableRowBorderSide, RtfTableRowBorderSide.Horizontal);
                    kind = RtfStyleKind.Table;
                    return true;
                case "trbrdrv":
                    BeginStyleTableRowBorder(tableRowFormat, pendingTableCell, ref currentTableRowBorderSide, RtfTableRowBorderSide.Vertical);
                    kind = RtfStyleKind.Table;
                    return true;
                case "clmgf":
                    pendingTableCell.HorizontalMerge = RtfTableCellMerge.First;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clmrg":
                    pendingTableCell.HorizontalMerge = RtfTableCellMerge.Continue;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clvmgf":
                    pendingTableCell.VerticalMerge = RtfTableCellMerge.First;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clvmrg":
                    pendingTableCell.VerticalMerge = RtfTableCellMerge.Continue;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clwWidth":
                    pendingTableCell.PreferredWidth = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clftsWidth":
                    pendingTableCell.PreferredWidthUnit = ToRtfTableWidthUnit(control.Parameter);
                    kind = RtfStyleKind.Table;
                    return true;
                case "clhidemark":
                    pendingTableCell.HideCellMark = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clNoWrap":
                    pendingTableCell.NoWrap = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clFitText":
                    pendingTableCell.FitText = true;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clcbpat":
                    pendingTableCell.BackgroundColorIndex = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clcfpat":
                    pendingTableCell.ShadingForegroundColorIndex = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clshdng":
                    pendingTableCell.ShadingPatternPercent = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clbghoriz":
                case "clbgvert":
                case "clbgfdiag":
                case "clbgbdiag":
                case "clbgcross":
                case "clbgdcross":
                case "clbgdkhor":
                case "clbgdkhoriz":
                case "clbgdkvert":
                case "clbgdkfdiag":
                case "clbgdkbdiag":
                case "clbgdkcross":
                case "clbgdkdcross":
                    pendingTableCell.ShadingPattern = ReadCellShadingPattern(control.Name);
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadt":
                    pendingTableCell.PaddingTopTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadl":
                    pendingTableCell.PaddingLeftTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadb":
                    pendingTableCell.PaddingBottomTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadr":
                    pendingTableCell.PaddingRightTwips = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadft":
                    pendingTableCell.PaddingTopUnit = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadfl":
                    pendingTableCell.PaddingLeftUnit = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadfb":
                    pendingTableCell.PaddingBottomUnit = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clpadfr":
                    pendingTableCell.PaddingRightUnit = control.Parameter;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clvertalt":
                    pendingTableCell.VerticalAlignment = RtfTableCellVerticalAlignment.Top;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clvertalc":
                    pendingTableCell.VerticalAlignment = RtfTableCellVerticalAlignment.Center;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clvertalb":
                    pendingTableCell.VerticalAlignment = RtfTableCellVerticalAlignment.Bottom;
                    kind = RtfStyleKind.Table;
                    return true;
                case "cltxlrtb":
                    pendingTableCell.TextFlow = RtfTableCellTextFlow.LeftToRightTopToBottom;
                    kind = RtfStyleKind.Table;
                    return true;
                case "cltxtbrl":
                    pendingTableCell.TextFlow = RtfTableCellTextFlow.TopToBottomRightToLeft;
                    kind = RtfStyleKind.Table;
                    return true;
                case "cltxbtlr":
                    pendingTableCell.TextFlow = RtfTableCellTextFlow.BottomToTopLeftToRight;
                    kind = RtfStyleKind.Table;
                    return true;
                case "cltxlrtbv":
                    pendingTableCell.TextFlow = RtfTableCellTextFlow.LeftToRightTopToBottomVertical;
                    kind = RtfStyleKind.Table;
                    return true;
                case "cltxtbrlv":
                    pendingTableCell.TextFlow = RtfTableCellTextFlow.TopToBottomRightToLeftVertical;
                    kind = RtfStyleKind.Table;
                    return true;
                case "clbrdrt":
                    BeginStyleTableCellBorder(pendingTableCell, ref currentTableRowBorderSide, RtfTableCellBorderSide.Top);
                    kind = RtfStyleKind.Table;
                    return true;
                case "clbrdrl":
                    BeginStyleTableCellBorder(pendingTableCell, ref currentTableRowBorderSide, RtfTableCellBorderSide.Left);
                    kind = RtfStyleKind.Table;
                    return true;
                case "clbrdrb":
                    BeginStyleTableCellBorder(pendingTableCell, ref currentTableRowBorderSide, RtfTableCellBorderSide.Bottom);
                    kind = RtfStyleKind.Table;
                    return true;
                case "clbrdrr":
                    BeginStyleTableCellBorder(pendingTableCell, ref currentTableRowBorderSide, RtfTableCellBorderSide.Right);
                    kind = RtfStyleKind.Table;
                    return true;
                case "cldglu":
                    BeginStyleTableCellBorder(pendingTableCell, ref currentTableRowBorderSide, RtfTableCellBorderSide.TopLeftToBottomRight);
                    kind = RtfStyleKind.Table;
                    return true;
                case "cldgll":
                    BeginStyleTableCellBorder(pendingTableCell, ref currentTableRowBorderSide, RtfTableCellBorderSide.TopRightToBottomLeft);
                    kind = RtfStyleKind.Table;
                    return true;
                case "cellx":
                    RtfTableCell styleCell = tableRowFormat.AddCell(control.Parameter);
                    pendingTableCell.ApplyTo(styleCell);
                    pendingTableCell = new PendingTableCellProperties();
                    currentTableRowBorderSide = null;
                    kind = RtfStyleKind.Table;
                    return true;
                default:
                    return false;
            }
        }

        private static void BeginStyleTableRowBorder(
            RtfTableRow row,
            PendingTableCellProperties cell,
            ref RtfTableRowBorderSide? currentTableRowBorderSide,
            RtfTableRowBorderSide side) {
            currentTableRowBorderSide = side;
            cell.CurrentBorderSide = null;
            row.GetBorder(side).Style = RtfTableCellBorderStyle.Single;
        }

        private static void BeginStyleTableCellBorder(
            PendingTableCellProperties cell,
            ref RtfTableRowBorderSide? currentTableRowBorderSide,
            RtfTableCellBorderSide side) {
            currentTableRowBorderSide = null;
            cell.CurrentBorderSide = side;
            cell.GetBorder(side).Style = RtfTableCellBorderStyle.Single;
        }

        private static bool ApplyStyleTableBorderStyle(
            RtfTableRow row,
            PendingTableCellProperties cell,
            RtfTableRowBorderSide? rowSide,
            RtfTableCellBorderStyle style) {
            RtfTableCellBorder? cellBorder = cell.GetCurrentBorder();
            if (cellBorder != null) {
                cellBorder.Style = style;
                return true;
            }

            if (!rowSide.HasValue) return false;

            row.GetBorder(rowSide.Value).Style = style;
            return true;
        }

        private static bool ApplyStyleTableBorderWidth(RtfTableRow row, PendingTableCellProperties cell, RtfTableRowBorderSide? rowSide, int? width) {
            RtfTableCellBorder? cellBorder = cell.GetCurrentBorder();
            if (cellBorder != null) {
                cellBorder.Width = width;
                return true;
            }

            if (!rowSide.HasValue) return false;

            row.GetBorder(rowSide.Value).Width = width;
            return true;
        }

        private static bool ApplyStyleTableBorderColor(RtfTableRow row, PendingTableCellProperties cell, RtfTableRowBorderSide? rowSide, int? colorIndex) {
            RtfTableCellBorder? cellBorder = cell.GetCurrentBorder();
            if (cellBorder != null) {
                cellBorder.ColorIndex = colorIndex;
                return true;
            }

            if (!rowSide.HasValue) return false;

            row.GetBorder(rowSide.Value).ColorIndex = colorIndex;
            return true;
        }

        private static RtfShadingPattern ReadTableRowShadingPattern(string controlName) {
            return controlName switch {
                "trbghoriz" => RtfShadingPattern.Horizontal,
                "trbgvert" => RtfShadingPattern.Vertical,
                "trbgfdiag" => RtfShadingPattern.ForwardDiagonal,
                "trbgbdiag" => RtfShadingPattern.BackwardDiagonal,
                "trbgcross" => RtfShadingPattern.Cross,
                "trbgdcross" => RtfShadingPattern.DiagonalCross,
                "trbgdkhor" or "trbgdkhoriz" => RtfShadingPattern.DarkHorizontal,
                "trbgdkvert" => RtfShadingPattern.DarkVertical,
                "trbgdkfdiag" => RtfShadingPattern.DarkForwardDiagonal,
                "trbgdkbdiag" => RtfShadingPattern.DarkBackwardDiagonal,
                "trbgdkcross" => RtfShadingPattern.DarkCross,
                "trbgdkdcross" => RtfShadingPattern.DarkDiagonalCross,
                _ => RtfShadingPattern.None
            };
        }

        private static RtfShadingPattern ReadCellShadingPattern(string controlName) {
            return controlName switch {
                "clbghoriz" => RtfShadingPattern.Horizontal,
                "clbgvert" => RtfShadingPattern.Vertical,
                "clbgfdiag" => RtfShadingPattern.ForwardDiagonal,
                "clbgbdiag" => RtfShadingPattern.BackwardDiagonal,
                "clbgcross" => RtfShadingPattern.Cross,
                "clbgdcross" => RtfShadingPattern.DiagonalCross,
                "clbgdkhor" or "clbgdkhoriz" => RtfShadingPattern.DarkHorizontal,
                "clbgdkvert" => RtfShadingPattern.DarkVertical,
                "clbgdkfdiag" => RtfShadingPattern.DarkForwardDiagonal,
                "clbgdkbdiag" => RtfShadingPattern.DarkBackwardDiagonal,
                "clbgdkcross" => RtfShadingPattern.DarkCross,
                "clbgdkdcross" => RtfShadingPattern.DarkDiagonalCross,
                _ => RtfShadingPattern.None
            };
        }

        private static void SetStyleTableRowBoxValue(RtfTableRow row, RowBoxMeasurements measurements, bool isPadding, RtfBoxSide side, int? value) {
            SetStyleTableRowBox(row, isPadding, side, measurements.SetValue(side, value));
        }

        private static void SetStyleTableRowBoxUnit(RtfTableRow row, RowBoxMeasurements measurements, bool isPadding, RtfBoxSide side, int? unit) {
            SetStyleTableRowBox(row, isPadding, side, measurements.SetUnit(side, unit));
        }

        private static void SetStyleTableRowBox(RtfTableRow row, bool isPadding, RtfBoxSide side, int? value) {
            if (isPadding) {
                switch (side) {
                    case RtfBoxSide.Top:
                        row.PaddingTopTwips = value;
                        return;
                    case RtfBoxSide.Left:
                        row.PaddingLeftTwips = value;
                        return;
                    case RtfBoxSide.Bottom:
                        row.PaddingBottomTwips = value;
                        return;
                    default:
                        row.PaddingRightTwips = value;
                        return;
                }
            }

            switch (side) {
                case RtfBoxSide.Top:
                    row.SpacingTopTwips = value;
                    return;
                case RtfBoxSide.Left:
                    row.SpacingLeftTwips = value;
                    return;
                case RtfBoxSide.Bottom:
                    row.SpacingBottomTwips = value;
                    return;
                default:
                    row.SpacingRightTwips = value;
                    return;
            }
        }

        private static void SetStyleTableRowHorizontalPosition(RtfTableRow row, RtfTableHorizontalPosition position, int? value) {
            row.HorizontalPosition = position;
            row.HorizontalPositionTwips = value;
        }

        private static void SetStyleTableRowVerticalPosition(RtfTableRow row, RtfTableVerticalPosition position, int? value) {
            row.VerticalPosition = position;
            row.VerticalPositionTwips = value;
        }

        private static void CopyStyleTableRowFormat(RtfTableRow source, RtfTableRow destination) {
            destination.RepeatHeader = source.RepeatHeader;
            destination.KeepTogether = source.KeepTogether;
            destination.KeepWithNext = source.KeepWithNext;
            destination.AutoFit = source.AutoFit;
            destination.Direction = source.Direction;
            destination.HeightTwips = source.HeightTwips;
            destination.CellGapTwips = source.CellGapTwips;
            destination.LeftIndentTwips = source.LeftIndentTwips;
            destination.Alignment = source.Alignment;
            destination.PreferredWidth = source.PreferredWidth;
            destination.PreferredWidthUnit = source.PreferredWidthUnit;
            destination.BackgroundColorIndex = source.BackgroundColorIndex;
            destination.ShadingForegroundColorIndex = source.ShadingForegroundColorIndex;
            destination.ShadingPatternValue = source.ShadingPatternValue;
            destination.ShadingPatternPercent = source.ShadingPatternPercent;
            destination.ShadingPattern = source.ShadingPattern;
            destination.PaddingTopTwips = source.PaddingTopTwips;
            destination.PaddingLeftTwips = source.PaddingLeftTwips;
            destination.PaddingBottomTwips = source.PaddingBottomTwips;
            destination.PaddingRightTwips = source.PaddingRightTwips;
            destination.SpacingTopTwips = source.SpacingTopTwips;
            destination.SpacingLeftTwips = source.SpacingLeftTwips;
            destination.SpacingBottomTwips = source.SpacingBottomTwips;
            destination.SpacingRightTwips = source.SpacingRightTwips;
            destination.NoOverlap = source.NoOverlap;
            destination.HorizontalAnchor = source.HorizontalAnchor;
            destination.VerticalAnchor = source.VerticalAnchor;
            destination.HorizontalPosition = source.HorizontalPosition;
            destination.HorizontalPositionTwips = source.HorizontalPositionTwips;
            destination.VerticalPosition = source.VerticalPosition;
            destination.VerticalPositionTwips = source.VerticalPositionTwips;
            destination.TextWrapLeftTwips = source.TextWrapLeftTwips;
            destination.TextWrapRightTwips = source.TextWrapRightTwips;
            destination.TextWrapTopTwips = source.TextWrapTopTwips;
            destination.TextWrapBottomTwips = source.TextWrapBottomTwips;
            CopyStyleTableRowBorder(source.TopBorder, destination.TopBorder);
            CopyStyleTableRowBorder(source.LeftBorder, destination.LeftBorder);
            CopyStyleTableRowBorder(source.BottomBorder, destination.BottomBorder);
            CopyStyleTableRowBorder(source.RightBorder, destination.RightBorder);
            CopyStyleTableRowBorder(source.HorizontalBorder, destination.HorizontalBorder);
            CopyStyleTableRowBorder(source.VerticalBorder, destination.VerticalBorder);

            foreach (RtfTableCell sourceCell in source.Cells) {
                RtfTableCell destinationCell = destination.AddCell(sourceCell.RightBoundaryTwips);
                CopyStyleTableCellFormat(sourceCell, destinationCell);
            }
        }

        private static void CopyStyleTableCellFormat(RtfTableCell source, RtfTableCell destination) {
            destination.HorizontalMerge = source.HorizontalMerge;
            destination.VerticalMerge = source.VerticalMerge;
            destination.BackgroundColorIndex = source.BackgroundColorIndex;
            destination.ShadingForegroundColorIndex = source.ShadingForegroundColorIndex;
            destination.ShadingPatternPercent = source.ShadingPatternPercent;
            destination.ShadingPattern = source.ShadingPattern;
            destination.VerticalAlignment = source.VerticalAlignment;
            destination.TextFlow = source.TextFlow;
            destination.PreferredWidth = source.PreferredWidth;
            destination.PreferredWidthUnit = source.PreferredWidthUnit;
            destination.HideCellMark = source.HideCellMark;
            destination.NoWrap = source.NoWrap;
            destination.FitText = source.FitText;
            destination.PaddingTopTwips = source.PaddingTopTwips;
            destination.PaddingLeftTwips = source.PaddingLeftTwips;
            destination.PaddingBottomTwips = source.PaddingBottomTwips;
            destination.PaddingRightTwips = source.PaddingRightTwips;
            CopyStyleTableCellBorder(source.TopBorder, destination.TopBorder);
            CopyStyleTableCellBorder(source.LeftBorder, destination.LeftBorder);
            CopyStyleTableCellBorder(source.BottomBorder, destination.BottomBorder);
            CopyStyleTableCellBorder(source.RightBorder, destination.RightBorder);
            CopyStyleTableCellBorder(source.TopLeftToBottomRightBorder, destination.TopLeftToBottomRightBorder);
            CopyStyleTableCellBorder(source.TopRightToBottomLeftBorder, destination.TopRightToBottomLeftBorder);
        }

        private static void CopyStyleTableRowBorder(RtfTableRowBorder source, RtfTableRowBorder destination) {
            destination.Style = source.Style;
            destination.Width = source.Width;
            destination.ColorIndex = source.ColorIndex;
        }

        private static void CopyStyleTableCellBorder(RtfTableCellBorder source, RtfTableCellBorder destination) {
            destination.Style = source.Style;
            destination.Width = source.Width;
            destination.ColorIndex = source.ColorIndex;
        }
    }
}
