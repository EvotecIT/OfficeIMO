namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static bool HaveEquivalentHtmlFormatting(RtfRun left, RtfRun right) {
        return left.Note == null &&
               right.Note == null &&
               left.Bold == right.Bold &&
               left.Italic == right.Italic &&
               left.UnderlineStyle == right.UnderlineStyle &&
               left.Strike == right.Strike &&
               left.DoubleStrike == right.DoubleStrike &&
               left.Hidden == right.Hidden &&
               left.Outline == right.Outline &&
               left.Shadow == right.Shadow &&
               left.Emboss == right.Emboss &&
               left.Imprint == right.Imprint &&
               left.CapsStyle == right.CapsStyle &&
               left.VerticalPosition == right.VerticalPosition &&
               left.FontSize == right.FontSize &&
               left.FontId == right.FontId &&
               left.ForegroundColorIndex == right.ForegroundColorIndex &&
               left.HighlightColorIndex == right.HighlightColorIndex &&
               left.CharacterBackgroundColorIndex == right.CharacterBackgroundColorIndex &&
               left.CharacterShadingForegroundColorIndex == right.CharacterShadingForegroundColorIndex &&
               left.CharacterShadingPatternPercent == right.CharacterShadingPatternPercent &&
               left.CharacterShadingPattern == right.CharacterShadingPattern &&
               left.CharacterBorder.Style == right.CharacterBorder.Style &&
               left.CharacterBorder.Width == right.CharacterBorder.Width &&
               left.CharacterBorder.ColorIndex == right.CharacterBorder.ColorIndex &&
               left.UnderlineColorIndex == right.UnderlineColorIndex &&
               left.CharacterSpacingTwips == right.CharacterSpacingTwips &&
               left.CharacterScalePercent == right.CharacterScalePercent &&
               left.KerningHalfPoints == right.KerningHalfPoints &&
               left.CharacterOffsetHalfPoints == right.CharacterOffsetHalfPoints &&
               left.StyleId == right.StyleId &&
               left.Direction == right.Direction &&
               left.LanguageId == right.LanguageId &&
               Equals(left.Hyperlink, right.Hyperlink) &&
               left.RevisionKind == right.RevisionKind &&
               left.RevisionAuthorIndex == right.RevisionAuthorIndex &&
               left.RevisionTimestampValue == right.RevisionTimestampValue &&
               left.CharacterRevisionSaveId == right.CharacterRevisionSaveId &&
               left.InsertionRevisionSaveId == right.InsertionRevisionSaveId &&
               left.DeletionRevisionSaveId == right.DeletionRevisionSaveId;
    }
}
