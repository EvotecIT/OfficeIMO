using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private RtfShape? ReadShape(RtfGroup group, CharacterState state, int depth) {
            var shape = new RtfShape();

            foreach (RtfNode node in group.Children) {
                if (node is RtfGroup childGroup) {
                    switch (childGroup.Destination) {
                        case "shpinst":
                            ReadShapeInstructionGroup(childGroup, shape, state, depth);
                            break;
                        case "sp":
                            ReadShapeProperty(childGroup, shape, state);
                            break;
                        case "shptxt":
                            ReadShapeText(childGroup, shape, state, depth);
                            break;
                    }
                }
            }

            return shape.Instructions.Count > 0 || shape.Properties.Count > 0 || shape.TextBoxParagraphs.Count > 0 ? shape : null;
        }

        private void ReadShapeInstructionGroup(RtfGroup group, RtfShape shape, CharacterState state, int depth) {
            foreach (RtfNode node in group.Children) {
                switch (node) {
                    case RtfControlWord control:
                        if (IsShapeInstructionControl(control.Name)) {
                            shape.AddInstruction(control.Name, control.Parameter, control.HasParameter);
                        }

                        break;
                    case RtfGroup childGroup:
                        switch (childGroup.Destination) {
                            case "sp":
                                ReadShapeProperty(childGroup, shape, state);
                                break;
                            case "shptxt":
                                ReadShapeText(childGroup, shape, state, depth);
                                break;
                            default:
                                ReadShapeInstructionGroup(childGroup, shape, state, depth + 1);
                                break;
                        }

                        break;
                }
            }
        }

        private static bool IsShapeInstructionControl(string name) =>
            name.StartsWith("shp", StringComparison.Ordinal) &&
            name != "shp" &&
            name != "shpinst" &&
            name != "shptxt";

        private static void ReadShapeProperty(RtfGroup group, RtfShape shape, CharacterState state) {
            string? name = null;
            string? value = null;
            foreach (RtfGroup childGroup in group.Children.OfType<RtfGroup>()) {
                switch (childGroup.Destination) {
                    case "sn":
                        name = CollectPlainText(childGroup, state.AnsiCodePage, state.UnicodeSkipCount).Trim();
                        break;
                    case "sv":
                        value = CollectPlainText(childGroup, state.AnsiCodePage, state.UnicodeSkipCount);
                        break;
                }
            }

            if (!string.IsNullOrWhiteSpace(name)) {
                shape.AddProperty(name!, value);
            }
        }

        private void ReadShapeText(RtfGroup group, RtfShape shape, CharacterState state, int depth) {
            RtfParagraph savedParagraph = _currentParagraph;
            RtfTable? savedTable = _currentTable;
            RtfTableRow? savedRow = _currentRow;
            RtfHeaderFooter? savedHeaderFooter = _currentHeaderFooter;
            RtfNote? savedNote = _currentNote;
            RtfShape? savedShape = _currentShape;
            int savedCellIndex = _currentCellIndex;
            bool savedParagraphIsInTable = _currentParagraphIsInTable;

            _currentShape = shape;
            _currentParagraph = new RtfParagraph();
            _currentTable = null;
            _currentRow = null;
            _currentHeaderFooter = null;
            _currentNote = null;
            _currentCellIndex = 0;
            _currentParagraphIsInTable = false;

            var childState = state.Clone();
            try {
                foreach (RtfNode node in group.Children) {
                    switch (node) {
                        case RtfControlWord control when control.Name == "shptxt":
                            break;
                        case RtfGroup childGroup:
                            WalkGroup(childGroup, childState.Clone(), depth + 1, allowDestinationSkip: true);
                            break;
                        case RtfText text:
                            AppendText(ApplySkip(childState, RtfAnsiCodePage.DecodeText(childState.AnsiCodePage, text.Text)), childState);
                            break;
                        case RtfControlWord control:
                            ApplyControlWord(control, childState);
                            break;
                        case RtfControlSymbol symbol:
                            ApplyControlSymbol(symbol, childState);
                            break;
                    }
                }

                FlushParagraphIfNeeded(force: false, childState);
            } finally {
                _currentParagraph = savedParagraph;
                _currentTable = savedTable;
                _currentRow = savedRow;
                _currentHeaderFooter = savedHeaderFooter;
                _currentNote = savedNote;
                _currentShape = savedShape;
                _currentCellIndex = savedCellIndex;
                _currentParagraphIsInTable = savedParagraphIsInTable;
            }
        }

        private void AddShape(RtfShape shape) {
            if (_currentParagraph.Inlines.Count > 0 ||
                _currentParagraphIsInTable ||
                _currentRow != null ||
                _currentNote != null ||
                _currentHeaderFooter != null ||
                _currentShape != null ||
                _inlineCaptureDepth > 0) {
                _currentParagraph.AddShape(shape);
                return;
            }

            AddDocumentBlock(shape);
        }
    }
}
