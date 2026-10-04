namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        // Paint callbacks can recursively parse forms and soft masks. Keep the dispatch frame
        // small instead of retaining every operator's graphics-state temporaries on that stack.
        private void ApplyOperator(string op, double paintOrder, bool hasInvalidOperands) {
            _currentPaintOrder = paintOrder;
            switch (op) {
                case "q":
                case "Q":
                case "cm":
                case "w":
                case "J":
                case "j":
                case "d":
                case "gs":
                case "ri":
                    ApplyStateOperator(op);
                    break;
                case "re":
                case "m":
                case "l":
                case "c":
                case "v":
                case "y":
                case "h":
                case "W":
                case "W*":
                case "n":
                case "S":
                case "s":
                case "f":
                case "F":
                case "f*":
                case "B":
                case "B*":
                case "b":
                case "b*":
                    ApplyPathsOperator(op);
                    break;
                case "cs":
                case "CS":
                case "sc":
                case "scn":
                case "SC":
                case "SCN":
                case "rg":
                case "RG":
                case "g":
                case "G":
                case "k":
                case "K":
                    ApplyColorsOperator(op);
                    break;
                case "BT":
                case "ET":
                case "Tf":
                case "Tm":
                case "Td":
                case "TD":
                case "TL":
                case "T*":
                case "Tc":
                case "Tw":
                case "Tz":
                case "Ts":
                case "Tr":
                case "'":
                case "\"":
                case "Tj":
                case "TJ":
                    ApplyTextOperator(op);
                    break;
                case "sh":
                case "Do":
                case "BI":
                    ApplyObjectsOperator(op, paintOrder);
                    break;
                case "BDC":
                case "BMC":
                case "EMC":
                    ApplyMarkedContentOperator(op, hasInvalidOperands);
                    break;
            }
            _args.Clear();
        }
    }
}
