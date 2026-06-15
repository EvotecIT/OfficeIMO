namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    private static void Apply(HtmlStyleDeclaration declaration, string property, string value, string rawValue) {
        switch (property) {
            case "font-weight":
                declaration.Bold = ParseFontWeight(value);
                break;
            case "font-style":
                declaration.Italic = ParseFontStyle(value);
                break;
            case "font-family":
                declaration.FontFamily = ParseFontFamily(rawValue);
                break;
            case "font-size":
                declaration.FontSizePoints = ParseFontSize(value);
                break;
            case "letter-spacing":
                declaration.CharacterSpacingTwips = ParseCharacterSpacing(value);
                break;
            case "font-stretch":
                declaration.CharacterScalePercent = ParseCharacterScale(value);
                break;
            case "--officeimo-rtf-character-scale":
                declaration.CharacterScalePercent = ParseRtfCharacterScale(value);
                break;
            case "--officeimo-rtf-character-offset":
                declaration.CharacterOffsetHalfPoints = ParseRtfCharacterOffset(value);
                break;
            case "text-decoration":
            case "text-decoration-line":
                ApplyTextDecoration(declaration, value);
                break;
            case "text-decoration-style":
                ApplyTextDecorationStyle(declaration, value);
                break;
            case "text-decoration-color":
                declaration.UnderlineColor = ParseColor(value);
                break;
            case "--officeimo-rtf-underline-style":
                declaration.UnderlineStyle = ParseRtfUnderlineStyle(value);
                break;
            case "--officeimo-rtf-strike-style":
                declaration.DoubleStrike = ParseRtfStrikeStyle(value);
                break;
            case "text-transform":
                declaration.CapsStyle = ParseTextTransform(value);
                break;
            case "font-variant":
            case "font-variant-caps":
                declaration.CapsStyle = ParseFontVariantCaps(value);
                break;
            case "--officeimo-rtf-caps-style":
                declaration.CapsStyle = ParseRtfCapsStyle(value);
                break;
            case "vertical-align":
                declaration.VerticalPosition = ParseVerticalAlign(value);
                declaration.TableCellVerticalAlignment = ParseTableCellVerticalAlign(value);
                declaration.CharacterOffsetHalfPoints = ParseCharacterOffset(value);
                break;
            case "writing-mode":
                declaration.TableCellTextFlow = ParseWritingMode(value);
                break;
            case "--officeimo-rtf-text-flow":
                declaration.TableCellTextFlow = ParseRtfTableCellTextFlow(value);
                break;
            case "direction":
                declaration.Direction = ParseDirection(value);
                break;
            case "--officeimo-rtf-direction":
                declaration.Direction = ParseDirection(value);
                break;
            case "--officeimo-rtf-lang":
                if (TryParseLanguageId(value, out int languageId)) {
                    declaration.LanguageId = languageId;
                }

                break;
            case "text-align":
                declaration.TextAlignment = ParseTextAlign(value);
                break;
            case "width":
                if (TryParseTableWidth(value, out int width, out RtfTableWidthUnit widthUnit)) {
                    declaration.TableWidth = width;
                    declaration.TableWidthUnit = widthUnit;
                }

                break;
            case "height":
                if (TryParseTwips(value, out int heightTwips)) {
                    declaration.TableHeightTwips = heightTwips;
                }

                break;
            case "white-space":
                declaration.NoWrap = ParseWhiteSpace(value);
                break;
            case "visibility":
                declaration.Hidden = ParseVisibility(value);
                break;
            case "text-shadow":
                declaration.Shadow = ParseTextShadow(value);
                break;
            case "--officeimo-rtf-hidden":
                declaration.Hidden = ParseBoolean(value);
                break;
            case "--officeimo-rtf-outline":
                declaration.Outline = ParseBoolean(value);
                break;
            case "--officeimo-rtf-shadow":
                declaration.Shadow = ParseBoolean(value);
                break;
            case "--officeimo-rtf-emboss":
                declaration.Emboss = ParseBoolean(value);
                break;
            case "--officeimo-rtf-imprint":
                declaration.Imprint = ParseBoolean(value);
                break;
            case "padding":
                ApplyPadding(declaration, value);
                break;
            case "padding-top":
                declaration.PaddingTopTwips = ParseTwips(value);
                break;
            case "padding-left":
                declaration.PaddingLeftTwips = ParseTwips(value);
                break;
            case "padding-bottom":
                declaration.PaddingBottomTwips = ParseTwips(value);
                break;
            case "padding-right":
                declaration.PaddingRightTwips = ParseTwips(value);
                break;
            case "border":
                HtmlBorderDeclaration? border = ParseBorder(value);
                if (border != null) {
                    declaration.TopBorder = CloneBorder(border);
                    declaration.LeftBorder = CloneBorder(border);
                    declaration.BottomBorder = CloneBorder(border);
                    declaration.RightBorder = CloneBorder(border);
                }

                break;
            case "border-top":
                declaration.TopBorder = ParseBorder(value);
                break;
            case "border-left":
                declaration.LeftBorder = ParseBorder(value);
                break;
            case "border-bottom":
                declaration.BottomBorder = ParseBorder(value);
                break;
            case "border-right":
                declaration.RightBorder = ParseBorder(value);
                break;
            case "margin-left":
                declaration.LeftIndentTwips = ParseTwips(value);
                break;
            case "margin-right":
                declaration.RightIndentTwips = ParseTwips(value);
                break;
            case "margin-top":
                declaration.SpaceBeforeTwips = ParseTwips(value);
                break;
            case "margin-bottom":
                declaration.SpaceAfterTwips = ParseTwips(value);
                break;
            case "text-indent":
                declaration.FirstLineIndentTwips = ParseTwips(value);
                break;
            case "line-height":
                ApplyLineHeight(declaration, value);
                break;
            case "page-break-before":
            case "break-before":
                declaration.PageBreakBefore = IsPageBreakValue(value);
                break;
            case "page-break-after":
            case "break-after":
                declaration.PageBreakAfter = IsPageBreakValue(value);
                break;
            case "color":
                declaration.ForegroundColor = ParseColor(value);
                break;
            case "background":
            case "background-color":
                declaration.BackgroundColor = ParseColor(value);
                break;
        }
    }
}
