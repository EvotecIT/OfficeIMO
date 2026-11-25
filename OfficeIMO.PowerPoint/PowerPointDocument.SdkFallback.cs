using System.Diagnostics;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The engine reads and edits presentations on its own package model and opens the Open XML SDK only for what that
    /// model does not cover. Each time it does, the document records why in <see cref="Diagnostics"/>.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private readonly List<OfficeDiagnostic> _diagnostics = new();

        /// <summary>
        /// Findings recorded while the document was loaded and used, such as each time it opened the Open XML SDK for
        /// something its own package handling does not cover. Codes come from <see cref="OfficeDiagnosticCodes"/>.
        /// </summary>
        public IReadOnlyList<OfficeDiagnostic> Diagnostics => _diagnostics;

        // A load that needed the SDK before anything was asked of the document.
        private void RecordSdkAtLoad() {
            _diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.SdkFallback, OfficeDiagnosticSeverity.Information,
                "The presentation was opened with the Open XML SDK because its package needs it: it is signed, Strict, " +
                "macro-enabled, a template or show, or was loaded with settings the engine leaves to the SDK.",
                location: null, format: "pptx"));
        }

        // The SDK opened after a load that did without it, for the public member that needed it.
        private void RecordSdkFallback() {
            string? member = FindPublicCaller();
            _diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.SdkFallback, OfficeDiagnosticSeverity.Information,
                member == null
                    ? "The presentation opened the Open XML SDK for an operation the engine does not do on its own package."
                    : "The presentation opened the Open XML SDK for " + member + ", which the engine does not do on its own package.",
                location: member, format: "pptx"));
        }

        // The outermost frame of this assembly on the stack: the call the user made. The trace is read as text, which
        // needs no method metadata; builds that strip frame names record the finding without a member.
        private static string? FindPublicCaller() {
            string? found = null;
            try {
                foreach (string line in new StackTrace(2, false).ToString().Split('\n')) {
                    // The word before the frame is localized ("at" in English), so the frame starts at the namespace.
                    int start = line.IndexOf(" OfficeIMO.PowerPoint.", StringComparison.Ordinal);
                    if (start < 0) continue;
                    string frame = line.Substring(start + 1);
                    int end = frame.IndexOf('(');
                    string qualified = end < 0 ? frame.Trim() : frame.Substring(0, end);
                    int method = qualified.LastIndexOf('.');
                    if (method <= 0) continue;
                    int type = qualified.LastIndexOf('.', method - 1);
                    // The library's public types sit in its root namespace; a frame from a deeper namespace may be the
                    // caller's own code that shares the prefix.
                    if (type != "OfficeIMO.PowerPoint".Length) continue;
                    string typeName = qualified.Substring(type + 1, method - type - 1);
                    string methodName = qualified.Substring(method + 1);
                    // Compiler-generated frames (lambdas, iterators, local functions) are not what the caller called.
                    if (typeName.IndexOf('<') >= 0 || methodName.IndexOf('<') >= 0) continue;
                    if (methodName.StartsWith("get_", StringComparison.Ordinal) || methodName.StartsWith("set_", StringComparison.Ordinal)) methodName = methodName.Substring(4);
                    found = typeName.Replace('+', '.') + "." + methodName;
                }
            } catch (Exception) {
                // A stack the runtime can't describe leaves the finding without a member.
            }
            return found;
        }
    }
}
