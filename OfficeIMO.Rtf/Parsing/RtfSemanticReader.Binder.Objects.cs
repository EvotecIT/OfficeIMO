using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private RtfObject? ReadObject(RtfGroup group, CharacterState state, int depth) {
            _limits.BeginObject(group.Position);
            var rtfObject = new RtfObject();
            bool hasObjectMetadata = false;

            foreach (RtfNode node in group.Children) {
                if (node is RtfControlWord control) {
                    if (TryApplyObjectControl(rtfObject, control)) {
                        hasObjectMetadata = true;
                    }
                } else if (node is RtfGroup childGroup) {
                    switch (childGroup.Destination) {
                        case "objclass":
                            rtfObject.ClassName = EmptyToNull(CollectPlainText(childGroup, state.AnsiCodePage, state.UnicodeSkipCount));
                            hasObjectMetadata |= rtfObject.ClassName != null;
                            break;
                        case "objname":
                            rtfObject.Name = EmptyToNull(CollectPlainText(childGroup, state.AnsiCodePage, state.UnicodeSkipCount));
                            hasObjectMetadata |= rtfObject.Name != null;
                            break;
                        case "objdata":
                            long objectBytes = 0;
                            rtfObject.Data = ReadObjectData(childGroup, ref objectBytes);
                            hasObjectMetadata |= rtfObject.Data.Length > 0;
                            break;
                        case "result":
                            ReadObjectResult(childGroup, rtfObject, state, depth);
                            hasObjectMetadata |= rtfObject.Result.Inlines.Count > 0 || rtfObject.ResultImage != null;
                            break;
                    }
                }
            }

            return hasObjectMetadata ? rtfObject : null;
        }

        private static bool TryApplyObjectControl(RtfObject rtfObject, RtfControlWord control) {
            switch (control.Name) {
                case "objemb":
                    rtfObject.Kind = RtfObjectKind.Embedded;
                    return true;
                case "objlink":
                    rtfObject.Kind = RtfObjectKind.Linked;
                    return true;
                case "objautlink":
                    rtfObject.Kind = RtfObjectKind.AutoLinked;
                    return true;
                case "objsub":
                    rtfObject.Kind = RtfObjectKind.Subscription;
                    return true;
                case "objpub":
                    rtfObject.Kind = RtfObjectKind.Publisher;
                    return true;
                case "objicemb":
                    rtfObject.Kind = RtfObjectKind.IconEmbedded;
                    return true;
                case "objw":
                    rtfObject.Width = control.Parameter;
                    return control.Parameter.HasValue;
                case "objh":
                    rtfObject.Height = control.Parameter;
                    return control.Parameter.HasValue;
                case "objscalex":
                    rtfObject.ScaleX = control.Parameter;
                    return control.Parameter.HasValue;
                case "objscaley":
                    rtfObject.ScaleY = control.Parameter;
                    return control.Parameter.HasValue;
                default:
                    return false;
            }
        }

        private byte[] ReadObjectData(RtfGroup group, ref long objectBytes) {
            var data = new List<byte>();
            foreach (RtfNode node in group.Children) {
                _limits.CheckCancellation();
                switch (node) {
                    case RtfBinary binary:
                        _limits.AddObjectBytes(ref objectBytes, binary.Data.Length, binary.Position);
                        data.AddRange(binary.Data);
                        break;
                    case RtfText text:
                        long decodedObjectBytes = objectBytes;
                        AppendHexBytes(text.Text, data, count => _limits.AddObjectBytes(ref decodedObjectBytes, count, text.Position));
                        objectBytes = decodedObjectBytes;
                        break;
                    case RtfGroup childGroup:
                        data.AddRange(ReadObjectData(childGroup, ref objectBytes));
                        break;
                }
            }

            return data.ToArray();
        }

        private void ReadObjectResult(RtfGroup group, RtfObject rtfObject, CharacterState state, int depth) {
            RtfParagraph savedParagraph = _currentParagraph;
            RtfTable? savedTable = _currentTable;
            RtfTableRow? savedRow = _currentRow;
            RtfHeaderFooter? savedHeaderFooter = _currentHeaderFooter;
            RtfNote? savedNote = _currentNote;
            RtfShape? savedShape = _currentShape;
            int savedCellIndex = _currentCellIndex;
            bool savedTableState = _currentParagraphIsInTable;

            _currentParagraph = rtfObject.Result;
            _currentTable = null;
            _currentRow = null;
            _currentHeaderFooter = null;
            _currentNote = null;
            _currentShape = null;
            _currentCellIndex = 0;
            _currentParagraphIsInTable = false;
            _inlineCaptureDepth++;
            try {
                var resultState = state.Clone();
                foreach (RtfNode child in group.Children) {
                    switch (child) {
                        case RtfGroup childGroup:
                            if (childGroup.Destination == "pict") {
                                rtfObject.ResultImage = ReadPicture(childGroup);
                                break;
                            }

                            WalkGroup(childGroup, resultState.Clone(), depth + 1, allowDestinationSkip: true);
                            break;
                        case RtfText text:
                            AppendText(ApplySkip(resultState, RtfAnsiCodePage.DecodeText(resultState.AnsiCodePage, text.Text)), resultState);
                            break;
                        case RtfControlWord control:
                            ApplyControlWord(control, resultState);
                            break;
                        case RtfControlSymbol symbol:
                            ApplyControlSymbol(symbol, resultState);
                            break;
                    }
                }
            } finally {
                _inlineCaptureDepth--;
                _currentParagraph = savedParagraph;
                _currentTable = savedTable;
                _currentRow = savedRow;
                _currentHeaderFooter = savedHeaderFooter;
                _currentNote = savedNote;
                _currentShape = savedShape;
                _currentCellIndex = savedCellIndex;
                _currentParagraphIsInTable = savedTableState;
            }
        }

        private static string? EmptyToNull(string value) => string.IsNullOrWhiteSpace(value) ? null : value;
    }
}
