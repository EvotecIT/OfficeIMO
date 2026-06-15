using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static bool TryApplyTabStopControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "tql":
                    state.PendingTabAlignment = RtfTabAlignment.Left;
                    return true;
                case "tqc":
                    state.PendingTabAlignment = RtfTabAlignment.Center;
                    return true;
                case "tqr":
                    state.PendingTabAlignment = RtfTabAlignment.Right;
                    return true;
                case "tqdec":
                    state.PendingTabAlignment = RtfTabAlignment.Decimal;
                    return true;
                case "tldot":
                    state.PendingTabLeader = RtfTabLeader.Dots;
                    return true;
                case "tlmdot":
                    state.PendingTabLeader = RtfTabLeader.MiddleDots;
                    return true;
                case "tlhyph":
                    state.PendingTabLeader = RtfTabLeader.Hyphen;
                    return true;
                case "tlul":
                    state.PendingTabLeader = RtfTabLeader.Underline;
                    return true;
                case "tlth":
                    state.PendingTabLeader = RtfTabLeader.ThickLine;
                    return true;
                case "tleq":
                    state.PendingTabLeader = RtfTabLeader.EqualSign;
                    return true;
                case "tx":
                    if (control.Parameter.HasValue && control.Parameter.Value >= 0) {
                        state.AddTabStop(control.Parameter.Value, state.PendingTabAlignment);
                    }

                    return true;
                case "tb":
                    if (control.Parameter.HasValue && control.Parameter.Value >= 0) {
                        state.AddTabStop(control.Parameter.Value, RtfTabAlignment.Bar);
                    }

                    return true;
                default:
                    return false;
            }
        }
    }
}
