using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfColor> ReadColorTable(RtfGroup root) {
            RtfGroup? table = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "colortbl");
            if (table == null) return Array.Empty<RtfColor>();

            var colors = new List<RtfColor>();
            var entry = new ColorTableEntry();
            foreach (RtfNode node in table.Children) {
                if (node is RtfControlWord control) {
                    ApplyColorTableControl(control, entry);
                } else if (node is RtfText text && text.Text.Contains(";")) {
                    if (entry.HasAnyValue) {
                        colors.Add(entry.ToColor());
                    }

                    entry = new ColorTableEntry();
                }
            }

            return colors;
        }

        private static void ApplyColorTableControl(RtfControlWord control, ColorTableEntry entry) {
            switch (control.Name) {
                case "red":
                    entry.Red = ClampByte(control.Parameter ?? 0);
                    entry.HasComponent = true;
                    break;
                case "green":
                    entry.Green = ClampByte(control.Parameter ?? 0);
                    entry.HasComponent = true;
                    break;
                case "blue":
                    entry.Blue = ClampByte(control.Parameter ?? 0);
                    entry.HasComponent = true;
                    break;
                case "ctint":
                    entry.Tint = control.Parameter;
                    break;
                case "cshade":
                    entry.Shade = control.Parameter;
                    break;
                default:
                    entry.ThemeColor = ToThemeColor(control.Name) ?? entry.ThemeColor;
                    break;
            }
        }

        private static RtfThemeColor? ToThemeColor(string controlName) {
            switch (controlName) {
                case "cmaindarkone":
                    return RtfThemeColor.MainDarkOne;
                case "cmainlightone":
                    return RtfThemeColor.MainLightOne;
                case "cmaindarktwo":
                    return RtfThemeColor.MainDarkTwo;
                case "cmainlighttwo":
                    return RtfThemeColor.MainLightTwo;
                case "caccentone":
                    return RtfThemeColor.AccentOne;
                case "caccenttwo":
                    return RtfThemeColor.AccentTwo;
                case "caccentthree":
                    return RtfThemeColor.AccentThree;
                case "caccentfour":
                    return RtfThemeColor.AccentFour;
                case "caccentfive":
                    return RtfThemeColor.AccentFive;
                case "caccentsix":
                    return RtfThemeColor.AccentSix;
                case "chyperlink":
                    return RtfThemeColor.Hyperlink;
                case "cfollowedhyperlink":
                    return RtfThemeColor.FollowedHyperlink;
                case "cbackgroundone":
                    return RtfThemeColor.BackgroundOne;
                case "ctextone":
                    return RtfThemeColor.TextOne;
                case "cbackgroundtwo":
                    return RtfThemeColor.BackgroundTwo;
                case "ctexttwo":
                    return RtfThemeColor.TextTwo;
                default:
                    return null;
            }
        }

        private static int ClampByte(int value) => Math.Max(0, Math.Min(255, value));

        private sealed class ColorTableEntry {
            public int Red;
            public int Green;
            public int Blue;
            public bool HasComponent;
            public RtfThemeColor? ThemeColor;
            public int? Tint;
            public int? Shade;

            public bool HasAnyValue => HasComponent || ThemeColor.HasValue || Tint.HasValue || Shade.HasValue;

            public RtfColor ToColor() {
                return new RtfColor((byte)Red, (byte)Green, (byte)Blue) {
                    ThemeColor = ThemeColor,
                    Tint = Tint,
                    Shade = Shade
                };
            }
        }
    }
}
