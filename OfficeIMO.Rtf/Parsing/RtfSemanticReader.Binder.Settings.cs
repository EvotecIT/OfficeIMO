using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static void ReadDocumentSettings(RtfGroup root, RtfDocumentSettings settings) {
            foreach (RtfNode child in root.Children) {
                if (!(child is RtfControlWord control)) {
                    continue;
                }

                switch (control.Name) {
                    case "sect":
                    case "sectd":
                    case "pard":
                        return;
                    case "ansi":
                        settings.CharacterSet = RtfDocumentCharacterSet.Ansi;
                        break;
                    case "mac":
                        settings.CharacterSet = RtfDocumentCharacterSet.Mac;
                        break;
                    case "pc":
                        settings.CharacterSet = RtfDocumentCharacterSet.Pc;
                        break;
                    case "pca":
                        settings.CharacterSet = RtfDocumentCharacterSet.Pca;
                        break;
                    case "ansicpg":
                        settings.AnsiCodePage = control.Parameter;
                        break;
                    case "uc":
                        if (control.Parameter.HasValue && control.Parameter.Value >= 0) {
                            settings.UnicodeSkipCount = control.Parameter;
                        }
                        break;
                    case "deff":
                        settings.DefaultFontId = control.Parameter;
                        break;
                    case "deftab":
                        settings.DefaultTabWidthTwips = control.Parameter;
                        break;
                    case "deflang":
                        settings.DefaultLanguageId = control.Parameter;
                        break;
                    case "deflangfe":
                        settings.DefaultFarEastLanguageId = control.Parameter;
                        break;
                    case "adeflang":
                        settings.DefaultAlternateLanguageId = control.Parameter;
                        break;
                    case "viewkind":
                        settings.ViewKind = control.Parameter;
                        break;
                    case "viewscale":
                        settings.ViewScale = control.Parameter;
                        break;
                    case "viewzk":
                        settings.ZoomKind = control.Parameter;
                        break;
                    case "viewbksp":
                        settings.ViewBackspaceBehavior = control.Parameter;
                        break;
                    case "widowctrl":
                        settings.WidowOrphanControl = ReadToggle(control);
                        break;
                    case "hyphauto":
                        settings.AutoHyphenation = ReadToggle(control);
                        break;
                    case "hyphcaps":
                        settings.HyphenateCaps = ReadToggle(control);
                        break;
                    case "hyphconsec":
                        settings.ConsecutiveHyphenLimit = control.Parameter;
                        break;
                    case "hyphhotz":
                        settings.HyphenationZoneTwips = control.Parameter;
                        break;
                    case "facingp":
                        settings.FacingPages = ReadToggle(control);
                        break;
                    case "margmirror":
                        settings.MirrorMargins = ReadToggle(control);
                        break;
                    case "formprot":
                        settings.FormProtection = ReadToggle(control);
                        break;
                    case "revprot":
                        settings.RevisionProtection = ReadToggle(control);
                        break;
                    case "annotprot":
                        settings.AnnotationProtection = ReadToggle(control);
                        break;
                    case "readprot":
                        settings.ReadOnlyProtection = ReadToggle(control);
                        break;
                    case "revisions":
                        settings.TrackRevisions = ReadToggle(control);
                        break;
                    case "revprop":
                        settings.RevisionDisplayStyle = control.Parameter;
                        break;
                    case "revbar":
                        settings.RevisionBarPlacement = control.Parameter;
                        break;
                    case "dghspace":
                        settings.DrawingGridHorizontalSpacingTwips = control.Parameter;
                        break;
                    case "dgvspace":
                        settings.DrawingGridVerticalSpacingTwips = control.Parameter;
                        break;
                    case "dghorigin":
                        settings.DrawingGridHorizontalOriginTwips = control.Parameter;
                        break;
                    case "dgvorigin":
                        settings.DrawingGridVerticalOriginTwips = control.Parameter;
                        break;
                    case "dghshow":
                        settings.DrawingGridHorizontalShow = control.Parameter;
                        break;
                    case "dgvshow":
                        settings.DrawingGridVerticalShow = control.Parameter;
                        break;
                    case "dgsnap":
                        settings.SnapToDrawingGrid = ReadToggle(control);
                        break;
                    case "dgmargin":
                        settings.DrawingGridUsesMargins = ReadToggle(control);
                        break;
                    case "ltrdoc":
                        settings.Direction = RtfTextDirection.LeftToRight;
                        break;
                    case "rtldoc":
                        settings.Direction = RtfTextDirection.RightToLeft;
                        break;
                }
            }
        }

        private static bool ReadToggle(RtfControlWord control) => !control.HasParameter || control.Parameter != 0;
    }
}
