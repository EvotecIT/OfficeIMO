using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private void ReadHeaderFooter(RtfGroup group, RtfHeaderFooterKind kind, CharacterState state, int depth) {
            RtfParagraph savedParagraph = _currentParagraph;
            RtfTable? savedTable = _currentTable;
            RtfTableRow? savedRow = _currentRow;
            RtfHeaderFooter? savedHeaderFooter = _currentHeaderFooter;
            int savedCellIndex = _currentCellIndex;
            bool savedParagraphIsInTable = _currentParagraphIsInTable;

            var headerFooter = new RtfHeaderFooter(kind);
            _document.AddParsedHeaderFooter(headerFooter);
            _currentHeaderFooter = headerFooter;
            _currentParagraph = new RtfParagraph();
            _currentTable = null;
            _currentRow = null;
            _currentCellIndex = 0;
            _currentParagraphIsInTable = false;

            var childState = state.Clone();
            foreach (RtfNode node in group.Children) {
                switch (node) {
                    case RtfControlWord control when control.Name == group.Destination:
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
            _currentParagraph = savedParagraph;
            _currentTable = savedTable;
            _currentRow = savedRow;
            _currentHeaderFooter = savedHeaderFooter;
            _currentCellIndex = savedCellIndex;
            _currentParagraphIsInTable = savedParagraphIsInTable;
        }

        private static RtfHeaderFooterKind? TryGetHeaderFooterKind(string? destination) {
            switch (destination) {
                case "header":
                    return RtfHeaderFooterKind.Header;
                case "headerl":
                    return RtfHeaderFooterKind.LeftHeader;
                case "headerr":
                    return RtfHeaderFooterKind.RightHeader;
                case "headerf":
                    return RtfHeaderFooterKind.FirstHeader;
                case "footer":
                    return RtfHeaderFooterKind.Footer;
                case "footerl":
                    return RtfHeaderFooterKind.LeftFooter;
                case "footerr":
                    return RtfHeaderFooterKind.RightFooter;
                case "footerf":
                    return RtfHeaderFooterKind.FirstFooter;
                default:
                    return null;
            }
        }
    }
}
