using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private void BeginTableRow() {
            if (_currentRow != null) {
                _currentCellDefinitionIndex = 0;
                _pendingCellProperties = new PendingTableCellProperties();
                return;
            }

            if (_currentTable == null) {
                _currentTable = new RtfTable();
                AddDocumentBlock(_currentTable);
            }

            _currentRow = _currentTable.AddRow();
            _currentCellIndex = 0;
            _currentCellDefinitionIndex = 0;
            _currentParagraphIsInTable = false;
            _pendingCellProperties = new PendingTableCellProperties();
            _currentRowBorderSide = null;
            _currentRowPadding = new RowBoxMeasurements();
            _currentRowSpacing = new RowBoxMeasurements();
        }

        private void AddTableCellBoundary(int? boundaryTwips) {
            if (_currentRow == null) {
                BeginTableRow();
            }

            RtfTableCell cell;
            if (_currentRow!.Cells.Count > _currentCellDefinitionIndex) {
                cell = _currentRow.Cells[_currentCellDefinitionIndex];
                cell.RightBoundaryTwips = boundaryTwips;
            } else {
                cell = _currentRow.AddCell(boundaryTwips);
            }

            _pendingCellProperties.ApplyTo(cell);
            _pendingCellProperties = new PendingTableCellProperties();
            _currentCellDefinitionIndex++;
        }

        private void FlushTableCell(CharacterState state) {
            ApplyParagraphState(_currentParagraph, state);
            AddParagraphToCurrentCell(_currentParagraph);
            _currentParagraph = new RtfParagraph();
            _currentParagraphIsInTable = false;
            _currentCellIndex++;
        }

        private void EndTableRow(CharacterState state) {
            if (_currentParagraph.Inlines.Count > 0) {
                FlushTableCell(state);
            }

            _currentRow = null;
            _currentCellIndex = 0;
            _currentCellDefinitionIndex = 0;
            _currentParagraphIsInTable = false;
        }

        private void BeginNestedTable(int level, CharacterState state) {
            if (level < 2 || _nestedCellParagraphs != null) return;
            if (_currentRow == null) BeginTableRow();
            if (_currentParagraph.Inlines.Count > 0) {
                ApplyParagraphState(_currentParagraph, state);
                AddParagraphToCurrentCell(_currentParagraph);
                _currentParagraph = new RtfParagraph();
            }

            _nestedCellParagraphs = new List<List<RtfParagraph>> { new List<RtfParagraph>() };
        }

        private void FlushNestedTableCell(CharacterState state) {
            if (_nestedCellParagraphs == null) return;
            List<RtfParagraph> paragraphs = _nestedCellParagraphs[_nestedCellParagraphs.Count - 1];
            if (_currentParagraph.Inlines.Count > 0 || paragraphs.Count == 0) {
                ApplyParagraphState(_currentParagraph, state);
                paragraphs.Add(_currentParagraph);
            }
            _currentParagraph = new RtfParagraph();
            _nestedCellParagraphs.Add(new List<RtfParagraph>());
        }

        private void ReadNestedTableProperties(RtfGroup group, CharacterState state) {
            if (_nestedCellParagraphs == null) return;
            if (_nestedCellParagraphs.Count > 1 && _nestedCellParagraphs[_nestedCellParagraphs.Count - 1].Count == 0) {
                _nestedCellParagraphs.RemoveAt(_nestedCellParagraphs.Count - 1);
            }

            RtfTable? outerTable = _currentTable;
            RtfTableRow? outerRow = _currentRow;
            int outerCellIndex = _currentCellIndex;
            int outerDefinitionIndex = _currentCellDefinitionIndex;
            PendingTableCellProperties outerPending = _pendingCellProperties;
            RtfTableRowBorderSide? outerBorder = _currentRowBorderSide;
            RowBoxMeasurements outerPadding = _currentRowPadding;
            RowBoxMeasurements outerSpacing = _currentRowSpacing;
            bool outerInTable = _currentParagraphIsInTable;
            RtfParagraph outerParagraph = _currentParagraph;

            if (_currentRow == null) BeginTableRow();
            while (_currentRow!.Cells.Count <= _currentCellIndex) _currentRow.AddCell();
            RtfTableCell outerCell = _currentRow.Cells[_currentCellIndex];
            RtfTable? existingNested = FindTrailingNestedTable(outerCell);
            var nested = existingNested ?? new RtfTable();
            _currentTable = nested;
            _currentRow = nested.AddRow();
            _currentCellIndex = 0;
            _currentCellDefinitionIndex = 0;
            _pendingCellProperties = new PendingTableCellProperties();
            _currentRowBorderSide = null;
            _currentRowPadding = new RowBoxMeasurements();
            _currentRowSpacing = new RowBoxMeasurements();
            _currentParagraph = new RtfParagraph();
            _currentParagraphIsInTable = true;

            foreach (RtfControlWord control in group.Children.OfType<RtfControlWord>()) {
                if (control.Name != "trowd" && control.Name != "nestrow") TryApplyTableControl(control, state);
            }

            while (_currentRow.Cells.Count < _nestedCellParagraphs.Count) _currentRow.AddCell();
            for (int cellIndex = 0; cellIndex < _nestedCellParagraphs.Count; cellIndex++) {
                foreach (RtfParagraph paragraph in _nestedCellParagraphs[cellIndex]) {
                    _currentRow.Cells[cellIndex].AddParsedParagraph(paragraph);
                }
            }

            _currentTable = outerTable;
            _currentRow = outerRow;
            _currentCellIndex = outerCellIndex;
            _currentCellDefinitionIndex = outerDefinitionIndex;
            _pendingCellProperties = outerPending;
            _currentRowBorderSide = outerBorder;
            _currentRowPadding = outerPadding;
            _currentRowSpacing = outerSpacing;
            _currentParagraphIsInTable = outerInTable;
            _currentParagraph = outerParagraph;
            if (existingNested == null) outerCell.AddParsedTable(nested);
            _nestedCellParagraphs = null;
        }

        private static RtfTable? FindTrailingNestedTable(RtfTableCell cell) {
            for (int index = cell.Blocks.Count - 1; index >= 0; index--) {
                if (cell.Blocks[index] is RtfTable table) {
                    for (int trailingIndex = index + 1; trailingIndex < cell.Blocks.Count; trailingIndex++) {
                        if (!(cell.Blocks[trailingIndex] is RtfParagraph paragraph) || paragraph.Inlines.Count > 0) {
                            return null;
                        }
                    }

                    return table;
                }

                if (!(cell.Blocks[index] is RtfParagraph trailingParagraph) || trailingParagraph.Inlines.Count > 0) {
                    break;
                }
            }

            return null;
        }

        private void AddParagraphToCurrentCell(RtfParagraph paragraph) {
            if (_currentRow == null) {
                BeginTableRow();
            }

            while (_currentRow!.Cells.Count <= _currentCellIndex) {
                _currentRow.AddCell();
            }

            _currentRow.Cells[_currentCellIndex].AddParsedParagraph(paragraph);
        }

        private void BeginPendingCellBorder(RtfTableCellBorderSide side) {
            _currentRowBorderSide = null;
            _pendingCellProperties.CurrentBorderSide = side;
            _pendingCellProperties.GetBorder(side).Style = RtfTableCellBorderStyle.Single;
        }

        private void BeginTableRowBorder(CharacterState state, RtfTableRowBorderSide side) {
            if (_currentRow == null) {
                BeginTableRow();
            }

            state.CurrentParagraphBorderSide = null;
            state.CurrentPageBorderSide = null;
            _pendingCellProperties.CurrentBorderSide = null;
            _currentRowBorderSide = side;
            _currentRow!.GetBorder(side).Style = RtfTableCellBorderStyle.Single;
        }

        private void ApplyPendingCellBorderStyle(RtfTableCellBorderStyle style) {
            RtfTableCellBorder? border = _pendingCellProperties.GetCurrentBorder();
            if (border != null) {
                border.Style = style;
            }
        }

        private void ApplyPendingCellBorderWidth(int? widthHalfPoints) {
            RtfTableCellBorder? border = _pendingCellProperties.GetCurrentBorder();
            if (border != null) {
                border.Width = widthHalfPoints;
            }
        }

        private void ApplyPendingCellBorderColor(int? colorIndex) {
            RtfTableCellBorder? border = _pendingCellProperties.GetCurrentBorder();
            if (border != null) {
                border.ColorIndex = colorIndex;
            }
        }

        private void BeginParagraphBorder(CharacterState state, RtfParagraphBorderSide side) {
            state.CurrentParagraphBorderSide = side;
            state.CurrentPageBorderSide = null;
            _currentRowBorderSide = null;
            _pendingCellProperties.CurrentBorderSide = null;
            state.GetParagraphBorder(side).Style = RtfParagraphBorderStyle.Single;
        }

        private void ApplyCurrentBorderStyle(CharacterState state, RtfParagraphBorderStyle style) {
            if (state.CurrentPageBorderSide.HasValue && TryApplyCurrentPageBorderStyle(state, ToPageBorderStyle(style))) {
                return;
            }

            RtfParagraphBorder? paragraphBorder = state.GetCurrentParagraphBorder();
            if (paragraphBorder != null) {
                paragraphBorder.Style = style;
                return;
            }

            if (state.CurrentCharacterBorderActive) {
                state.CharacterBorder.Style = style;
                return;
            }

            RtfTableRowBorder? rowBorder = GetCurrentRowBorder();
            if (rowBorder != null) {
                rowBorder.Style = ToCellBorderStyle(style);
                return;
            }

            ApplyPendingCellBorderStyle(ToCellBorderStyle(style));
        }

        private void ApplyCurrentBorderWidth(CharacterState state, int? width) {
            if (state.CurrentPageBorderSide.HasValue && TryApplyCurrentPageBorderWidth(state, width)) {
                return;
            }

            RtfParagraphBorder? paragraphBorder = state.GetCurrentParagraphBorder();
            if (paragraphBorder != null) {
                paragraphBorder.Width = width;
                return;
            }

            if (state.CurrentCharacterBorderActive) {
                state.CharacterBorder.Width = width;
                return;
            }

            RtfTableRowBorder? rowBorder = GetCurrentRowBorder();
            if (rowBorder != null) {
                rowBorder.Width = width;
                return;
            }

            ApplyPendingCellBorderWidth(width);
        }

        private void ApplyCurrentBorderColor(CharacterState state, int? colorIndex) {
            if (state.CurrentPageBorderSide.HasValue && TryApplyCurrentPageBorderColor(state, colorIndex)) {
                return;
            }

            RtfParagraphBorder? paragraphBorder = state.GetCurrentParagraphBorder();
            if (paragraphBorder != null) {
                paragraphBorder.ColorIndex = colorIndex;
                return;
            }

            if (state.CurrentCharacterBorderActive) {
                state.CharacterBorder.ColorIndex = colorIndex;
                return;
            }

            RtfTableRowBorder? rowBorder = GetCurrentRowBorder();
            if (rowBorder != null) {
                rowBorder.ColorIndex = colorIndex;
                return;
            }

            ApplyPendingCellBorderColor(colorIndex);
        }

        private RtfTableRowBorder? GetCurrentRowBorder() {
            return _currentRow != null && _currentRowBorderSide.HasValue
                ? _currentRow.GetBorder(_currentRowBorderSide.Value)
                : null;
        }

        private void ApplyTableRowShadingPattern(RtfShadingPattern pattern) {
            if (_currentRow != null) {
                _currentRow.ShadingPattern = pattern;
            }
        }

        private void SetTableRowPadding(RtfBoxSide side, int? value) {
            if (_currentRow == null) return;
            SetTableRowPaddingProperty(side, _currentRowPadding.SetValue(side, value));
        }

        private void SetTableRowPaddingUnit(RtfBoxSide side, int? unit) {
            if (_currentRow == null) return;
            SetTableRowPaddingProperty(side, _currentRowPadding.SetUnit(side, unit));
        }

        private void SetTableRowPaddingProperty(RtfBoxSide side, int? value) {
            if (_currentRow == null) return;
            switch (side) {
                case RtfBoxSide.Top:
                    _currentRow.PaddingTopTwips = value;
                    return;
                case RtfBoxSide.Left:
                    _currentRow.PaddingLeftTwips = value;
                    return;
                case RtfBoxSide.Bottom:
                    _currentRow.PaddingBottomTwips = value;
                    return;
                default:
                    _currentRow.PaddingRightTwips = value;
                    return;
            }
        }

        private void SetTableRowSpacing(RtfBoxSide side, int? value) {
            if (_currentRow == null) return;
            SetTableRowSpacingProperty(side, _currentRowSpacing.SetValue(side, value));
        }

        private void SetTableRowSpacingUnit(RtfBoxSide side, int? unit) {
            if (_currentRow == null) return;
            SetTableRowSpacingProperty(side, _currentRowSpacing.SetUnit(side, unit));
        }

        private void SetTableRowSpacingProperty(RtfBoxSide side, int? value) {
            if (_currentRow == null) return;
            switch (side) {
                case RtfBoxSide.Top:
                    _currentRow.SpacingTopTwips = value;
                    return;
                case RtfBoxSide.Left:
                    _currentRow.SpacingLeftTwips = value;
                    return;
                case RtfBoxSide.Bottom:
                    _currentRow.SpacingBottomTwips = value;
                    return;
                default:
                    _currentRow.SpacingRightTwips = value;
                    return;
            }
        }

        private void SetTableRowHorizontalAnchor(RtfTableHorizontalAnchor anchor) {
            if (_currentRow != null) {
                _currentRow.HorizontalAnchor = anchor;
            }
        }

        private void SetTableRowVerticalAnchor(RtfTableVerticalAnchor anchor) {
            if (_currentRow != null) {
                _currentRow.VerticalAnchor = anchor;
            }
        }

        private void SetTableRowHorizontalPosition(RtfTableHorizontalPosition position, int? value) {
            if (_currentRow == null) return;
            _currentRow.HorizontalPosition = position;
            _currentRow.HorizontalPositionTwips = value;
        }

        private void SetTableRowVerticalPosition(RtfTableVerticalPosition position, int? value) {
            if (_currentRow == null) return;
            _currentRow.VerticalPosition = position;
            _currentRow.VerticalPositionTwips = value;
        }

        private static RtfTableCellBorderStyle ToCellBorderStyle(RtfParagraphBorderStyle style) {
            switch (style) {
                case RtfParagraphBorderStyle.Double:
                    return RtfTableCellBorderStyle.Double;
                case RtfParagraphBorderStyle.Dotted:
                    return RtfTableCellBorderStyle.Dotted;
                case RtfParagraphBorderStyle.Dashed:
                    return RtfTableCellBorderStyle.Dashed;
                case RtfParagraphBorderStyle.Single:
                    return RtfTableCellBorderStyle.Single;
                default:
                    return RtfTableCellBorderStyle.None;
            }
        }

        private static RtfTableWidthUnit? ToRtfTableWidthUnit(int? value) {
            switch (value) {
                case 1:
                    return RtfTableWidthUnit.Auto;
                case 2:
                    return RtfTableWidthUnit.Percent;
                case 3:
                    return RtfTableWidthUnit.Twips;
                default:
                    return null;
            }
        }

        private static RtfPageBorderStyle ToPageBorderStyle(RtfParagraphBorderStyle style) {
            switch (style) {
                case RtfParagraphBorderStyle.Double:
                    return RtfPageBorderStyle.Double;
                case RtfParagraphBorderStyle.Dotted:
                    return RtfPageBorderStyle.Dotted;
                case RtfParagraphBorderStyle.Dashed:
                    return RtfPageBorderStyle.Dashed;
                case RtfParagraphBorderStyle.None:
                    return RtfPageBorderStyle.None;
                default:
                    return RtfPageBorderStyle.Single;
            }
        }

        private enum RtfBoxSide {
            Top,
            Left,
            Bottom,
            Right
        }

        private sealed class RowBoxMeasurements {
            private int? _topValue;
            private int? _leftValue;
            private int? _bottomValue;
            private int? _rightValue;
            private int? _topUnit;
            private int? _leftUnit;
            private int? _bottomUnit;
            private int? _rightUnit;

            public int? SetValue(RtfBoxSide side, int? value) {
                switch (side) {
                    case RtfBoxSide.Top:
                        _topValue = value;
                        return IsTwipsUnit(_topUnit) ? value : null;
                    case RtfBoxSide.Left:
                        _leftValue = value;
                        return IsTwipsUnit(_leftUnit) ? value : null;
                    case RtfBoxSide.Bottom:
                        _bottomValue = value;
                        return IsTwipsUnit(_bottomUnit) ? value : null;
                    default:
                        _rightValue = value;
                        return IsTwipsUnit(_rightUnit) ? value : null;
                }
            }

            public int? SetUnit(RtfBoxSide side, int? unit) {
                switch (side) {
                    case RtfBoxSide.Top:
                        _topUnit = unit;
                        return IsTwipsUnit(unit) ? _topValue : null;
                    case RtfBoxSide.Left:
                        _leftUnit = unit;
                        return IsTwipsUnit(unit) ? _leftValue : null;
                    case RtfBoxSide.Bottom:
                        _bottomUnit = unit;
                        return IsTwipsUnit(unit) ? _bottomValue : null;
                    default:
                        _rightUnit = unit;
                        return IsTwipsUnit(unit) ? _rightValue : null;
                }
            }

            private static bool IsTwipsUnit(int? unit) => !unit.HasValue || unit.Value == 3;
        }

        private sealed class PendingTableCellProperties {
            public RtfTableCellMerge HorizontalMerge { get; set; }

            public RtfTableCellMerge VerticalMerge { get; set; }

            public int? BackgroundColorIndex { get; set; }

            public int? ShadingForegroundColorIndex { get; set; }

            public int? ShadingPatternPercent { get; set; }

            public RtfShadingPattern ShadingPattern { get; set; } = RtfShadingPattern.None;

            public RtfTableCellVerticalAlignment? VerticalAlignment { get; set; }

            public RtfTableCellTextFlow? TextFlow { get; set; }

            public int? PreferredWidth { get; set; }

            public RtfTableWidthUnit? PreferredWidthUnit { get; set; }

            public bool HideCellMark { get; set; }

            public bool NoWrap { get; set; }

            public bool FitText { get; set; }

            public int? PaddingTopTwips { get; set; }

            public int? PaddingLeftTwips { get; set; }

            public int? PaddingBottomTwips { get; set; }

            public int? PaddingRightTwips { get; set; }

            public int? PaddingTopUnit { get; set; }

            public int? PaddingLeftUnit { get; set; }

            public int? PaddingBottomUnit { get; set; }

            public int? PaddingRightUnit { get; set; }

            public RtfTableCellBorderSide? CurrentBorderSide { get; set; }

            public RtfTableCellBorder TopBorder { get; } = new RtfTableCellBorder();

            public RtfTableCellBorder LeftBorder { get; } = new RtfTableCellBorder();

            public RtfTableCellBorder BottomBorder { get; } = new RtfTableCellBorder();

            public RtfTableCellBorder RightBorder { get; } = new RtfTableCellBorder();

            public RtfTableCellBorder TopLeftToBottomRightBorder { get; } = new RtfTableCellBorder();

            public RtfTableCellBorder TopRightToBottomLeftBorder { get; } = new RtfTableCellBorder();

            public RtfTableCellBorder GetBorder(RtfTableCellBorderSide side) {
                switch (side) {
                    case RtfTableCellBorderSide.Top:
                        return TopBorder;
                    case RtfTableCellBorderSide.Left:
                        return LeftBorder;
                    case RtfTableCellBorderSide.Bottom:
                        return BottomBorder;
                    case RtfTableCellBorderSide.Right:
                        return RightBorder;
                    case RtfTableCellBorderSide.TopLeftToBottomRight:
                        return TopLeftToBottomRightBorder;
                    default:
                        return TopRightToBottomLeftBorder;
                }
            }

            public RtfTableCellBorder? GetCurrentBorder() {
                return CurrentBorderSide.HasValue ? GetBorder(CurrentBorderSide.Value) : null;
            }

            public void ApplyTo(RtfTableCell cell) {
                cell.HorizontalMerge = HorizontalMerge;
                cell.VerticalMerge = VerticalMerge;
                cell.BackgroundColorIndex = BackgroundColorIndex;
                cell.ShadingForegroundColorIndex = ShadingForegroundColorIndex;
                cell.ShadingPatternPercent = ShadingPatternPercent;
                cell.ShadingPattern = ShadingPattern;
                cell.VerticalAlignment = VerticalAlignment;
                cell.TextFlow = TextFlow;
                cell.PreferredWidth = PreferredWidth;
                cell.PreferredWidthUnit = PreferredWidthUnit;
                cell.HideCellMark = HideCellMark;
                cell.NoWrap = NoWrap;
                cell.FitText = FitText;
                cell.PaddingTopTwips = IsTwipsPaddingUnit(PaddingTopUnit) ? PaddingTopTwips : null;
                cell.PaddingLeftTwips = IsTwipsPaddingUnit(PaddingLeftUnit) ? PaddingLeftTwips : null;
                cell.PaddingBottomTwips = IsTwipsPaddingUnit(PaddingBottomUnit) ? PaddingBottomTwips : null;
                cell.PaddingRightTwips = IsTwipsPaddingUnit(PaddingRightUnit) ? PaddingRightTwips : null;
                CopyBorder(TopBorder, cell.TopBorder);
                CopyBorder(LeftBorder, cell.LeftBorder);
                CopyBorder(BottomBorder, cell.BottomBorder);
                CopyBorder(RightBorder, cell.RightBorder);
                CopyBorder(TopLeftToBottomRightBorder, cell.TopLeftToBottomRightBorder);
                CopyBorder(TopRightToBottomLeftBorder, cell.TopRightToBottomLeftBorder);
            }

            private static bool IsTwipsPaddingUnit(int? unit) => !unit.HasValue || unit.Value == 3;

            private static void CopyBorder(RtfTableCellBorder source, RtfTableCellBorder destination) {
                destination.Style = source.Style;
                destination.Width = source.Width;
                destination.ColorIndex = source.ColorIndex;
            }
        }
    }
}
