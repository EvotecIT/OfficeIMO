using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        private void ApplyMarkedContentOperator(string op, bool hasInvalidOperands) {
            switch (op) {
                case "BDC":
                    _markedContentMcidStack.Push(GetMcid(_args.Count > 0 ? _args[_args.Count - 1] : null));
                    _artifactContentStack.Push(IsArtifactTag(_args.Count > 1 ? _args[_args.Count - 2] : null));
                    _hiddenContentStack.Push(
                        hasInvalidOperands ||
                        IsHiddenOptionalContent(
                            _args.Count > 1 ? _args[_args.Count - 2] : null,
                            _args.Count > 0 ? _args[_args.Count - 1] : null));
                    break;
                case "BMC":
                    _markedContentMcidStack.Push(null);
                    _artifactContentStack.Push(IsArtifactTag(_args.Count > 0 ? _args[_args.Count - 1] : null));
                    _hiddenContentStack.Push(hasInvalidOperands);
                    break;
                case "EMC":
                    if (_hiddenContentStack.Count > 0) {
                        _hiddenContentStack.Pop();
                    }
                    if (_markedContentMcidStack.Count > 0) {
                        _markedContentMcidStack.Pop();
                    }
                    if (_artifactContentStack.Count > 0) {
                        _artifactContentStack.Pop();
                    }

                    break;
            }
        }
    }
}
