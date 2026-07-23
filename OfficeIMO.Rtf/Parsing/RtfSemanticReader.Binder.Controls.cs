using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static void ApplyCharacterSetDefaultCodePage(CharacterState state, RtfDocumentCharacterSet characterSet) {
            if (!state.HasExplicitAnsiCodePage) {
                state.DocumentAnsiCodePage = RtfAnsiCodePage.GetDefaultCodePage(characterSet);
                state.AnsiCodePage = state.DocumentAnsiCodePage;
            }
        }

        private void ApplyControlWord(RtfControlWord control, CharacterState state) {
            if (TryApplyTabStopControl(control, state)) return;
            if (TryApplyBreakControl(control, state)) return;
            if (TryApplySectionControl(control, state)) return;
            if (TryApplyParagraphFrameControl(control, state)) return;
            if (TryApplyCharacterDecorationControl(control, state)) return;
            if (TryApplyLegacyNumberingControl(control, state)) return;
            if (TryApplyTableControl(control, state)) return;

            switch (control.Name) {
                case "rtf":
                    return;
                case "ansi":
                    _document.Settings.CharacterSet = RtfDocumentCharacterSet.Ansi;
                    ApplyCharacterSetDefaultCodePage(state, RtfDocumentCharacterSet.Ansi);
                    return;
                case "mac":
                    _document.Settings.CharacterSet = RtfDocumentCharacterSet.Mac;
                    ApplyCharacterSetDefaultCodePage(state, RtfDocumentCharacterSet.Mac);
                    return;
                case "pc":
                    _document.Settings.CharacterSet = RtfDocumentCharacterSet.Pc;
                    ApplyCharacterSetDefaultCodePage(state, RtfDocumentCharacterSet.Pc);
                    return;
                case "pca":
                    _document.Settings.CharacterSet = RtfDocumentCharacterSet.Pca;
                    ApplyCharacterSetDefaultCodePage(state, RtfDocumentCharacterSet.Pca);
                    return;
                case "deff":
                    _document.Settings.DefaultFontId = control.Parameter;
                    if (!state.FontId.HasValue) state.AnsiCodePage = ResolveFontCodePage(control.Parameter, state.DocumentAnsiCodePage);
                    return;
                case "deflang":
                    if (control.Parameter.HasValue) {
                        _document.Settings.DefaultLanguageId = control.Parameter;
                        state.DefaultLanguageId = control.Parameter;
                        state.LanguageId = control.Parameter;
                    }
                    return;
                case "deflangfe":
                    _document.Settings.DefaultFarEastLanguageId = control.Parameter;
                    return;
                case "adeflang":
                    _document.Settings.DefaultAlternateLanguageId = control.Parameter;
                    return;
                case "charrsid":
                    state.CharacterRevisionSaveId = control.Parameter;
                    return;
                case "insrsid":
                    state.InsertionRevisionSaveId = control.Parameter;
                    return;
                case "delrsid":
                    state.DeletionRevisionSaveId = control.Parameter;
                    return;
                case "pararsid":
                    state.ParagraphRevisionSaveId = control.Parameter;
                    return;
                case "ansicpg":
                    if (control.Parameter.HasValue) {
                        _document.Settings.AnsiCodePage = control.Parameter;
                        state.DocumentAnsiCodePage = control.Parameter.Value;
                        state.AnsiCodePage = control.Parameter.Value;
                        state.HasExplicitAnsiCodePage = true;
                    }
                    return;
                case "pard":
                    RtfLegacyNumbering? pendingLegacyNumbering = null;
                    if (state.HasPendingLegacyNumberingAfterReset) {
                        pendingLegacyNumbering = new RtfLegacyNumbering();
                        pendingLegacyNumbering.CopyFrom(state.PendingLegacyNumberingAfterReset);
                    }

                    RtfParagraph? pendingListText = state.PendingListTextAfterReset;
                    state.ResetParagraph();
                    state.AnsiCodePage = ResolveFontCodePage(_document.Settings.DefaultFontId, state.DocumentAnsiCodePage);
                    if (pendingLegacyNumbering != null) {
                        ApplyLegacyNumberingToState(state, pendingLegacyNumbering);
                    }

                    if (pendingListText != null) {
                        state.ListText = pendingListText;
                        state.PendingListTextAfterReset = pendingListText;
                    }

                    return;
                case "plain":
                    state.ResetCharacter();
                    state.AnsiCodePage = ResolveFontCodePage(_document.Settings.DefaultFontId, state.DocumentAnsiCodePage);
                    return;
                case "par":
                    FlushParagraphIfNeeded(force: true, state);
                    return;
                case "pagebb":
                    state.PageBreakBefore = !control.HasParameter || control.Parameter != 0;
                    return;
                case "keepn":
                    state.KeepWithNext = !control.HasParameter || control.Parameter != 0;
                    return;
                case "keep":
                    state.KeepLinesTogether = !control.HasParameter || control.Parameter != 0;
                    return;
                case "noline":
                    state.SuppressLineNumbers = !control.HasParameter || control.Parameter != 0;
                    return;
                case "hyphpar":
                    state.AutoHyphenation = !control.HasParameter || control.Parameter != 0;
                    return;
                case "contextualspace":
                    state.ContextualSpacing = !control.HasParameter || control.Parameter != 0;
                    return;
                case "adjustright":
                    state.AdjustRightIndent = !control.HasParameter || control.Parameter != 0;
                    return;
                case "nosnaplinegrid":
                    state.SnapToLineGrid = control.HasParameter && control.Parameter == 0;
                    return;
                case "widctlpar":
                    state.WidowControl = !control.HasParameter || control.Parameter != 0;
                    return;
                case "nowidctlpar":
                    state.WidowControl = false;
                    return;
                case "outlinelevel":
                    state.OutlineLevel = control.Parameter;
                    return;
                case "tab":
                    AppendText("\t", state);
                    return;
                case "chpgn":
                    AppendGeneratedText(RtfGeneratedTextKind.PageNumber, state);
                    return;
                case "sectnum":
                    AppendGeneratedText(RtfGeneratedTextKind.SectionNumber, state);
                    return;
                case "chdate":
                    AppendGeneratedText(RtfGeneratedTextKind.CurrentDate, state);
                    return;
                case "chdpl":
                    AppendGeneratedText(RtfGeneratedTextKind.CurrentDateLong, state);
                    return;
                case "chdpa":
                    AppendGeneratedText(RtfGeneratedTextKind.CurrentDateAbbreviated, state);
                    return;
                case "chtime":
                    AppendGeneratedText(RtfGeneratedTextKind.CurrentTime, state);
                    return;
                case "chftn":
                    AppendGeneratedText(RtfGeneratedTextKind.NoteReference, state);
                    return;
                case "emdash":
                case "endash":
                case "emspace":
                case "enspace":
                case "qmspace":
                case "bullet":
                case "lquote":
                case "rquote":
                case "ldblquote":
                case "rdblquote":
                case "ltrmark":
                case "rtlmark":
                case "zwj":
                case "zwnj":
                    AppendText(GetSpecialCharacterText(control.Name), state);
                    return;
                case "qc":
                    state.Alignment = RtfTextAlignment.Center;
                    return;
                case "qr":
                    state.Alignment = RtfTextAlignment.Right;
                    return;
                case "qj":
                    state.Alignment = RtfTextAlignment.Justify;
                    return;
                case "ql":
                    state.Alignment = RtfTextAlignment.Left;
                    return;
                case "rtlpar":
                    state.ParagraphDirection = RtfTextDirection.RightToLeft;
                    return;
                case "ltrpar":
                    state.ParagraphDirection = RtfTextDirection.LeftToRight;
                    return;
                case "b":
                    state.Bold = !control.HasParameter || control.Parameter != 0;
                    return;
                case "i":
                    state.Italic = !control.HasParameter || control.Parameter != 0;
                    return;
                case "ul":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Single : RtfUnderlineStyle.None;
                    return;
                case "ulw":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Words : RtfUnderlineStyle.None;
                    return;
                case "uldb":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Double : RtfUnderlineStyle.None;
                    return;
                case "uld":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Dotted : RtfUnderlineStyle.None;
                    return;
                case "uldash":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Dash : RtfUnderlineStyle.None;
                    return;
                case "uldashd":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.DashDot : RtfUnderlineStyle.None;
                    return;
                case "uldashdd":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.DashDotDot : RtfUnderlineStyle.None;
                    return;
                case "ulth":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Thick : RtfUnderlineStyle.None;
                    return;
                case "ulthd":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.ThickDotted : RtfUnderlineStyle.None;
                    return;
                case "ulthdash":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.ThickDash : RtfUnderlineStyle.None;
                    return;
                case "ulthdashd":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.ThickDashDot : RtfUnderlineStyle.None;
                    return;
                case "ulthdashdd":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.ThickDashDotDot : RtfUnderlineStyle.None;
                    return;
                case "ulwave":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.Wave : RtfUnderlineStyle.None;
                    return;
                case "ulhwave":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.HeavyWave : RtfUnderlineStyle.None;
                    return;
                case "uldbwave":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.DoubleWave : RtfUnderlineStyle.None;
                    return;
                case "ulldash":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.LongDash : RtfUnderlineStyle.None;
                    return;
                case "ulthldash":
                    state.UnderlineStyle = !control.HasParameter || control.Parameter != 0 ? RtfUnderlineStyle.ThickLongDash : RtfUnderlineStyle.None;
                    return;
                case "ulnone":
                    state.UnderlineStyle = RtfUnderlineStyle.None;
                    return;
                case "ulc":
                    state.UnderlineColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "expndtw":
                    state.CharacterSpacingTwips = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "expnd":
                    state.CharacterSpacingTwips = control.Parameter.HasValue && control.Parameter.Value != 0 ? control.Parameter.Value * 5 : null;
                    return;
                case "charscalex":
                    state.CharacterScalePercent = control.Parameter.GetValueOrDefault() == 100 ? null : control.Parameter;
                    return;
                case "kerning":
                    state.KerningHalfPoints = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "up":
                    state.CharacterOffsetHalfPoints = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "dn":
                    state.CharacterOffsetHalfPoints = control.Parameter.HasValue && control.Parameter.Value != 0 ? -control.Parameter.Value : null;
                    return;
                case "strike":
                    state.Strike = !control.HasParameter || control.Parameter != 0;
                    return;
                case "striked":
                    state.DoubleStrike = !control.HasParameter || control.Parameter != 0;
                    return;
                case "caps":
                    state.CapsStyle = !control.HasParameter || control.Parameter != 0 ? RtfCapsStyle.Caps : RtfCapsStyle.None;
                    return;
                case "scaps":
                    state.CapsStyle = !control.HasParameter || control.Parameter != 0 ? RtfCapsStyle.SmallCaps : RtfCapsStyle.None;
                    return;
                case "v":
                    state.Hidden = !control.HasParameter || control.Parameter != 0;
                    return;
                case "outl":
                    state.Outline = !control.HasParameter || control.Parameter != 0;
                    return;
                case "shad":
                    state.Shadow = !control.HasParameter || control.Parameter != 0;
                    return;
                case "embo":
                    state.Emboss = !control.HasParameter || control.Parameter != 0;
                    return;
                case "impr":
                    state.Imprint = !control.HasParameter || control.Parameter != 0;
                    return;
                case "rtlch":
                    state.Direction = RtfTextDirection.RightToLeft;
                    return;
                case "ltrch":
                    state.Direction = RtfTextDirection.LeftToRight;
                    return;
                case "revised":
                    state.RevisionKind = !control.HasParameter || control.Parameter != 0 ? RtfRevisionKind.Inserted : RtfRevisionKind.None;
                    return;
                case "deleted":
                    state.RevisionKind = !control.HasParameter || control.Parameter != 0 ? RtfRevisionKind.Deleted : RtfRevisionKind.None;
                    return;
                case "revauth":
                    state.RevisionAuthorIndex = control.Parameter;
                    return;
                case "revdttm":
                    state.RevisionTimestampValue = control.Parameter;
                    return;
                case "super":
                    state.VerticalPosition = !control.HasParameter || control.Parameter != 0
                        ? RtfVerticalPosition.Superscript
                        : RtfVerticalPosition.Baseline;
                    return;
                case "sub":
                    state.VerticalPosition = !control.HasParameter || control.Parameter != 0
                        ? RtfVerticalPosition.Subscript
                        : RtfVerticalPosition.Baseline;
                    return;
                case "nosupersub":
                    state.VerticalPosition = RtfVerticalPosition.Baseline;
                    return;
                case "fs":
                    if (control.Parameter.HasValue) {
                        state.FontSize = control.Parameter.Value / 2d;
                    }
                    return;
                case "f":
                    state.FontId = control.Parameter;
                    state.AnsiCodePage = ResolveFontCodePage(control.Parameter, state.DocumentAnsiCodePage);
                    return;
                case "cf":
                    state.ForegroundColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "highlight":
                    state.HighlightColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "cs":
                    state.CharacterStyleId = control.Parameter;
                    return;
                case "lang":
                    state.LanguageId = control.Parameter;
                    return;
                case "s":
                    state.ParagraphStyleId = control.Parameter;
                    return;
                case "ls":
                    state.ListId = control.Parameter;
                    ApplyListOverride(state);
                    return;
                case "ilvl":
                    state.ListLevel = control.Parameter.HasValue
                        ? Math.Min(8, Math.Max(0, control.Parameter.Value))
                        : null;
                    ApplyListLevel(state);
                    if (state.ListKind == RtfListKind.None) {
                        state.ListKind = RtfListKind.Decimal;
                    }
                    return;
                case "pn":
                    state.ListKind = RtfListKind.Decimal;
                    return;
                case "pnlvlblt":
                    state.ListKind = RtfListKind.Bullet;
                    return;
                case "li":
                    state.LeftIndentTwips = control.Parameter;
                    return;
                case "ri":
                    state.RightIndentTwips = control.Parameter;
                    return;
                case "fi":
                    state.FirstLineIndentTwips = control.Parameter;
                    return;
                case "sb":
                    state.SpaceBeforeTwips = control.Parameter;
                    return;
                case "sa":
                    state.SpaceAfterTwips = control.Parameter;
                    return;
                case "sbauto":
                    state.SpaceBeforeAuto = !control.HasParameter || control.Parameter != 0;
                    return;
                case "saauto":
                    state.SpaceAfterAuto = !control.HasParameter || control.Parameter != 0;
                    return;
                case "sl":
                    state.LineSpacingTwips = control.Parameter;
                    return;
                case "slmult":
                    state.LineSpacingMultiple = !control.HasParameter || control.Parameter != 0;
                    return;
                case "cbpat":
                    state.BackgroundColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "cfpat":
                    state.ShadingForegroundColorIndex = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "shading":
                    state.ShadingPatternPercent = control.Parameter.GetValueOrDefault() == 0 ? null : control.Parameter;
                    return;
                case "bghoriz":
                    state.ShadingPattern = RtfShadingPattern.Horizontal;
                    return;
                case "bgvert":
                    state.ShadingPattern = RtfShadingPattern.Vertical;
                    return;
                case "bgfdiag":
                    state.ShadingPattern = RtfShadingPattern.ForwardDiagonal;
                    return;
                case "bgbdiag":
                    state.ShadingPattern = RtfShadingPattern.BackwardDiagonal;
                    return;
                case "bgcross":
                    state.ShadingPattern = RtfShadingPattern.Cross;
                    return;
                case "bgdcross":
                    state.ShadingPattern = RtfShadingPattern.DiagonalCross;
                    return;
                case "bgdkhoriz":
                    state.ShadingPattern = RtfShadingPattern.DarkHorizontal;
                    return;
                case "bgdkvert":
                    state.ShadingPattern = RtfShadingPattern.DarkVertical;
                    return;
                case "bgdkfdiag":
                    state.ShadingPattern = RtfShadingPattern.DarkForwardDiagonal;
                    return;
                case "bgdkbdiag":
                    state.ShadingPattern = RtfShadingPattern.DarkBackwardDiagonal;
                    return;
                case "bgdkcross":
                    state.ShadingPattern = RtfShadingPattern.DarkCross;
                    return;
                case "bgdkdcross":
                    state.ShadingPattern = RtfShadingPattern.DarkDiagonalCross;
                    return;
                case "brdrt":
                    BeginParagraphBorder(state, RtfParagraphBorderSide.Top);
                    return;
                case "brdrl":
                    BeginParagraphBorder(state, RtfParagraphBorderSide.Left);
                    return;
                case "brdrb":
                    BeginParagraphBorder(state, RtfParagraphBorderSide.Bottom);
                    return;
                case "brdrr":
                    BeginParagraphBorder(state, RtfParagraphBorderSide.Right);
                    return;
                case "brdrs":
                    ApplyCurrentBorderStyle(state, RtfParagraphBorderStyle.Single);
                    return;
                case "brdrdb":
                    ApplyCurrentBorderStyle(state, RtfParagraphBorderStyle.Double);
                    return;
                case "brdrdot":
                    ApplyCurrentBorderStyle(state, RtfParagraphBorderStyle.Dotted);
                    return;
                case "brdrdash":
                    ApplyCurrentBorderStyle(state, RtfParagraphBorderStyle.Dashed);
                    return;
                case "brdrnone":
                case "brdrnil":
                    ApplyCurrentBorderStyle(state, RtfParagraphBorderStyle.None);
                    return;
                case "brdrw":
                    ApplyCurrentBorderWidth(state, control.Parameter);
                    return;
                case "brdrcf":
                    ApplyCurrentBorderColor(state, control.Parameter);
                    return;
                case "uc":
                    if (control.Parameter.HasValue && control.Parameter.Value >= 0) {
                        state.UnicodeSkipCount = control.Parameter.Value;
                    }
                    return;
                case "u":
                    if (control.Parameter.HasValue) {
                        AppendUnicodeValue(control.Parameter.Value, state);
                        state.SkipCharacters = state.UnicodeSkipCount;
                    }
                    return;
            }
        }

    }
}
