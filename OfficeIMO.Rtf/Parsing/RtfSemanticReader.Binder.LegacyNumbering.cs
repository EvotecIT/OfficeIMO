using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static void ReadLegacyNumbering(RtfGroup group, CharacterState state) {
            state.LegacyNumbering.Clear();
            state.LegacyNumbering.Enabled = true;
            state.ListKind = RtfListKind.Decimal;

            foreach (RtfNode node in group.Children) {
                if (node is RtfControlWord control) {
                    TryApplyLegacyNumberingControl(control, state);
                } else if (node is RtfGroup childGroup) {
                    if (childGroup.Destination == "pntxtb") {
                        state.LegacyNumbering.TextBefore = CollectPlainText(childGroup, state.AnsiCodePage, state.UnicodeSkipCount);
                    } else if (childGroup.Destination == "pntxta") {
                        state.LegacyNumbering.TextAfter = CollectPlainText(childGroup, state.AnsiCodePage, state.UnicodeSkipCount);
                    }
                }
            }

            state.PendingLegacyNumberingAfterReset.CopyFrom(state.LegacyNumbering);
            state.HasPendingLegacyNumberingAfterReset = true;
        }

        private static void ApplyLegacyNumberingToState(CharacterState state, RtfLegacyNumbering numbering) {
            state.LegacyNumbering.CopyFrom(numbering);
            switch (numbering.LevelKind) {
                case RtfLegacyNumberingLevelKind.Bullet:
                    state.ListKind = RtfListKind.Bullet;
                    break;
                case RtfLegacyNumberingLevelKind.Continue:
                    state.ListKind = RtfListKind.None;
                    break;
                default:
                    state.ListKind = RtfListKind.Decimal;
                    break;
            }

            if (numbering.Level.HasValue) {
                state.ListLevel = Math.Max(0, numbering.Level.Value - 1);
            }
        }

        private static bool TryApplyLegacyNumberingControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "pn":
                    state.LegacyNumbering.Clear();
                    state.LegacyNumbering.Enabled = true;
                    state.ListKind = RtfListKind.Decimal;
                    return true;
                case "pnlvl":
                    state.LegacyNumbering.Enabled = true;
                    state.LegacyNumbering.LevelKind = RtfLegacyNumberingLevelKind.Level;
                    state.LegacyNumbering.Level = control.Parameter;
                    state.ListLevel = control.Parameter.HasValue ? Math.Max(0, control.Parameter.Value - 1) : state.ListLevel;
                    state.ListKind = RtfListKind.Decimal;
                    return true;
                case "pnlvlblt":
                    state.LegacyNumbering.Enabled = true;
                    state.LegacyNumbering.LevelKind = RtfLegacyNumberingLevelKind.Bullet;
                    state.ListKind = RtfListKind.Bullet;
                    return true;
                case "pnlvlbody":
                    state.LegacyNumbering.Enabled = true;
                    state.LegacyNumbering.LevelKind = RtfLegacyNumberingLevelKind.Body;
                    state.ListKind = RtfListKind.Decimal;
                    return true;
                case "pnlvlcont":
                    state.LegacyNumbering.Enabled = true;
                    state.LegacyNumbering.LevelKind = RtfLegacyNumberingLevelKind.Continue;
                    state.ListKind = RtfListKind.None;
                    return true;
                case "pncard":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.Cardinal;
                    return true;
                case "pndec":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.Decimal;
                    return true;
                case "pnucltr":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.UpperLetter;
                    return true;
                case "pnucrm":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.UpperRoman;
                    return true;
                case "pnlcltr":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.LowerLetter;
                    return true;
                case "pnlcrm":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.LowerRoman;
                    return true;
                case "pnord":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.Ordinal;
                    return true;
                case "pnordt":
                    state.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.OrdinalText;
                    return true;
                case "pnnumonce":
                    state.LegacyNumbering.NumberEachCellOnce = ReadLegacyToggle(control);
                    return true;
                case "pnacross":
                    state.LegacyNumbering.NumberAcrossRows = ReadLegacyToggle(control);
                    return true;
                case "pnhang":
                    state.LegacyNumbering.HangingIndent = ReadLegacyToggle(control);
                    return true;
                case "pnrestart":
                    state.LegacyNumbering.RestartAfterSection = ReadLegacyToggle(control);
                    return true;
                case "pnprev":
                    state.LegacyNumbering.IncludePreviousLevels = ReadLegacyToggle(control);
                    return true;
                case "pnindent":
                    state.LegacyNumbering.IndentTwips = control.Parameter;
                    return true;
                case "pnsp":
                    state.LegacyNumbering.SpaceTwips = control.Parameter;
                    return true;
                case "pnstart":
                    state.LegacyNumbering.StartAt = control.Parameter;
                    return true;
                case "pnql":
                    state.LegacyNumbering.Alignment = RtfLegacyNumberingAlignment.Left;
                    return true;
                case "pnqc":
                    state.LegacyNumbering.Alignment = RtfLegacyNumberingAlignment.Center;
                    return true;
                case "pnqr":
                    state.LegacyNumbering.Alignment = RtfLegacyNumberingAlignment.Right;
                    return true;
                case "pnf":
                    state.LegacyNumbering.FontId = control.Parameter;
                    return true;
                case "pnfs":
                    state.LegacyNumbering.FontSizeHalfPoints = control.Parameter;
                    return true;
                case "pncf":
                    state.LegacyNumbering.ForegroundColorIndex = control.Parameter;
                    return true;
                case "pnb":
                    state.LegacyNumbering.Bold = ReadLegacyToggle(control);
                    return true;
                case "pni":
                    state.LegacyNumbering.Italic = ReadLegacyToggle(control);
                    return true;
                case "pncaps":
                    state.LegacyNumbering.AllCaps = ReadLegacyToggle(control);
                    return true;
                case "pnscaps":
                    state.LegacyNumbering.SmallCaps = ReadLegacyToggle(control);
                    return true;
                case "pnul":
                    state.LegacyNumbering.UnderlineStyle = ReadLegacyToggle(control) == false ? RtfUnderlineStyle.None : RtfUnderlineStyle.Single;
                    return true;
                case "pnuld":
                    state.LegacyNumbering.UnderlineStyle = RtfUnderlineStyle.Dotted;
                    return true;
                case "pnuldb":
                    state.LegacyNumbering.UnderlineStyle = RtfUnderlineStyle.Double;
                    return true;
                case "pnulnone":
                    state.LegacyNumbering.UnderlineStyle = RtfUnderlineStyle.None;
                    return true;
                case "pnulw":
                    state.LegacyNumbering.UnderlineStyle = RtfUnderlineStyle.Words;
                    return true;
                case "pnstrike":
                    state.LegacyNumbering.Strike = ReadLegacyToggle(control);
                    return true;
                default:
                    return false;
            }
        }

        private static bool ReadLegacyToggle(RtfControlWord control) => !control.HasParameter || control.Parameter != 0;
    }
}
