using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyStylesheet(Dictionary<string, string> values) {
            var styles = new List<RtfStyle>();
            for (int index = 0; ; index++) {
                string prefix = "style." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, prefix + ".id");
                string? name = ReadString(values, prefix + ".name");
                if (!id.HasValue || string.IsNullOrWhiteSpace(name)) {
                    break;
                }

                RtfStyleKind kind = ReadEnum(values, prefix + ".kind", RtfStyleKind.Paragraph);
                var style = new RtfStyle(id.Value, name!, kind);
                ApplyStyle(values, prefix, style);
                styles.Add(style);
            }

            if (styles.Count > 0) {
                _document.ReplaceStyles(styles);
            }
        }

        private void ApplyParagraphStyleAttributes(HtmlToken token) {
            int? styleId = ReadStyleIdAttribute(token);
            if (styleId.HasValue) {
                EnsureParagraph().StyleId = styleId.Value;
            }
        }

        private static int? ReadStyleIdAttribute(HtmlToken token) {
            string? value = GetAttribute(token, "data-officeimo-rtf-style-id");
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int styleId) && styleId >= 0
                ? styleId
                : null;
        }

        private static bool IsInlineStyleScope(string tagName) {
            switch (tagName) {
                case "a":
                case "b":
                case "code":
                case "del":
                case "em":
                case "i":
                case "ins":
                case "s":
                case "span":
                case "strike":
                case "strong":
                case "sub":
                case "sup":
                case "u":
                    return true;
                default:
                    return false;
            }
        }

        private static void ApplyStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
            style.BasedOnStyleId = ReadInt(values, prefix + ".basedOn");
            style.NextStyleId = ReadInt(values, prefix + ".next");
            style.LinkedStyleId = ReadInt(values, prefix + ".linked");
            style.KeyCode = ReadKeyCode(values, prefix + ".key");
            style.Additive = ReadBool(values, prefix + ".additive") == true;
            style.AutoUpdate = ReadBool(values, prefix + ".autoUpdate") == true;
            style.Hidden = ReadBool(values, prefix + ".hidden") == true;
            style.Locked = ReadBool(values, prefix + ".locked") == true;
            style.Personal = ReadBool(values, prefix + ".personal") == true;
            style.Compose = ReadBool(values, prefix + ".compose") == true;
            style.Reply = ReadBool(values, prefix + ".reply") == true;
            style.SemiHidden = ReadBool(values, prefix + ".semiHidden") == true;
            style.UnhideWhenUsed = ReadBool(values, prefix + ".unhideWhenUsed") == true;
            style.QuickFormat = ReadBool(values, prefix + ".quickFormat") == true;
            style.Priority = ReadInt(values, prefix + ".priority");
            style.RevisionSaveId = ReadInt(values, prefix + ".revisionSaveId");
            style.Bold = ReadBool(values, prefix + ".bold");
            style.Italic = ReadBool(values, prefix + ".italic");
            style.UnderlineStyle = ReadEnum<RtfUnderlineStyle>(values, prefix + ".underlineStyle");
            style.FontSize = ReadDouble(values, prefix + ".fontSize");
            style.FontId = ReadInt(values, prefix + ".fontId");
            style.ForegroundColorIndex = ReadInt(values, prefix + ".foregroundColor");
            style.HighlightColorIndex = ReadInt(values, prefix + ".highlightColor");
            ApplyParagraphStyle(values, prefix + ".paragraph", style);
            ApplyTableStyle(values, prefix + ".table", style);
        }

        private static void ApplyParagraphStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
            style.ParagraphAlignment = ReadEnum<RtfTextAlignment>(values, prefix + ".alignment");
            style.ParagraphDirection = ReadEnum<RtfTextDirection>(values, prefix + ".direction");
            style.LeftIndentTwips = ReadInt(values, prefix + ".leftIndent");
            style.RightIndentTwips = ReadInt(values, prefix + ".rightIndent");
            style.FirstLineIndentTwips = ReadInt(values, prefix + ".firstLineIndent");
            style.SpaceBeforeTwips = ReadInt(values, prefix + ".spaceBefore");
            style.SpaceAfterTwips = ReadInt(values, prefix + ".spaceAfter");
            style.SpaceBeforeAuto = ReadBool(values, prefix + ".spaceBeforeAuto");
            style.SpaceAfterAuto = ReadBool(values, prefix + ".spaceAfterAuto");
            style.LineSpacingTwips = ReadInt(values, prefix + ".lineSpacing");
            style.LineSpacingMultiple = ReadBool(values, prefix + ".lineSpacingMultiple");
            style.BackgroundColorIndex = ReadInt(values, prefix + ".backgroundColor");
            style.ShadingForegroundColorIndex = ReadInt(values, prefix + ".shadingForeground");
            style.ShadingPatternPercent = ReadInt(values, prefix + ".shadingPercent");
            style.ShadingPattern = ReadEnum(values, prefix + ".shadingPattern", RtfShadingPattern.None);
            style.PageBreakBefore = ReadBool(values, prefix + ".pageBreakBefore");
            style.KeepWithNext = ReadBool(values, prefix + ".keepWithNext");
            style.KeepLinesTogether = ReadBool(values, prefix + ".keepLinesTogether");
            style.SuppressLineNumbers = ReadBool(values, prefix + ".suppressLineNumbers");
            style.AutoHyphenation = ReadBool(values, prefix + ".autoHyphenation");
            style.ContextualSpacing = ReadBool(values, prefix + ".contextualSpacing");
            style.AdjustRightIndent = ReadBool(values, prefix + ".adjustRightIndent");
            style.SnapToLineGrid = ReadBool(values, prefix + ".snapToLineGrid");
            style.WidowControl = ReadBool(values, prefix + ".widowControl");
            style.OutlineLevel = ReadInt(values, prefix + ".outlineLevel");
            ApplyBorder(values, prefix + ".border.top", style.TopBorder);
            ApplyBorder(values, prefix + ".border.left", style.LeftBorder);
            ApplyBorder(values, prefix + ".border.bottom", style.BottomBorder);
            ApplyBorder(values, prefix + ".border.right", style.RightBorder);
            ApplyParagraphFrame(values, prefix + ".frame", style.Frame);
            style.ReplaceTabStops(ReadTabStops(values, prefix + ".tab"));
        }

        private static RtfStyleKeyCode? ReadKeyCode(Dictionary<string, string> values, string prefix) {
            bool shift = ReadBool(values, prefix + ".shift") == true;
            bool control = ReadBool(values, prefix + ".control") == true;
            bool alt = ReadBool(values, prefix + ".alt") == true;
            int? functionKey = ReadInt(values, prefix + ".function");
            string? key = ReadString(values, prefix + ".key");
            if (!shift && !control && !alt && !functionKey.HasValue && string.IsNullOrWhiteSpace(key)) {
                return null;
            }

            return new RtfStyleKeyCode {
                Shift = shift,
                Control = control,
                Alt = alt,
                FunctionKey = functionKey,
                Key = key
            };
        }

        private static void ApplyBorder(Dictionary<string, string> values, string prefix, RtfParagraphBorder border) {
            border.Style = ReadEnum(values, prefix + ".style", RtfParagraphBorderStyle.None);
            border.Width = ReadInt(values, prefix + ".width");
            border.ColorIndex = ReadInt(values, prefix + ".color");
        }

        private static IReadOnlyList<RtfTabStop> ReadTabStops(Dictionary<string, string> values, string prefix) {
            var tabStops = new List<RtfTabStop>();
            for (int index = 0; ; index++) {
                string tabPrefix = prefix + "." + index.ToString(CultureInfo.InvariantCulture);
                int? position = ReadInt(values, tabPrefix + ".position");
                if (!position.HasValue) {
                    break;
                }

                tabStops.Add(new RtfTabStop(
                    position.Value,
                    ReadEnum(values, tabPrefix + ".alignment", RtfTabAlignment.Left),
                    ReadEnum(values, tabPrefix + ".leader", RtfTabLeader.None)));
            }

            return tabStops;
        }

        private static double? ReadDouble(Dictionary<string, string> values, string key) {
            return values.TryGetValue(key, out string? value) &&
                   double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : null;
        }
    }
}
