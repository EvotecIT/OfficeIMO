namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteColorTable(StringBuilder builder, RtfDocument document) {
        if (document.Colors.Count == 0) return;

        builder.Append(@"{\colortbl;");
        foreach (RtfColor color in document.Colors) {
            builder.Append(@"\red");
            builder.Append(color.Red.ToString(CultureInfo.InvariantCulture));
            builder.Append(@"\green");
            builder.Append(color.Green.ToString(CultureInfo.InvariantCulture));
            builder.Append(@"\blue");
            builder.Append(color.Blue.ToString(CultureInfo.InvariantCulture));
            WriteThemeColor(builder, color.ThemeColor);
            AppendOptionalTwips(builder, @"\ctint", color.Tint);
            AppendOptionalTwips(builder, @"\cshade", color.Shade);
            builder.Append(';');
        }

        builder.Append('}');
    }

    private static void WriteThemeColor(StringBuilder builder, RtfThemeColor? themeColor) {
        if (!themeColor.HasValue) return;

        builder.Append(themeColor.Value switch {
            RtfThemeColor.MainDarkOne => @"\cmaindarkone",
            RtfThemeColor.MainLightOne => @"\cmainlightone",
            RtfThemeColor.MainDarkTwo => @"\cmaindarktwo",
            RtfThemeColor.MainLightTwo => @"\cmainlighttwo",
            RtfThemeColor.AccentOne => @"\caccentone",
            RtfThemeColor.AccentTwo => @"\caccenttwo",
            RtfThemeColor.AccentThree => @"\caccentthree",
            RtfThemeColor.AccentFour => @"\caccentfour",
            RtfThemeColor.AccentFive => @"\caccentfive",
            RtfThemeColor.AccentSix => @"\caccentsix",
            RtfThemeColor.Hyperlink => @"\chyperlink",
            RtfThemeColor.FollowedHyperlink => @"\cfollowedhyperlink",
            RtfThemeColor.BackgroundOne => @"\cbackgroundone",
            RtfThemeColor.TextOne => @"\ctextone",
            RtfThemeColor.BackgroundTwo => @"\cbackgroundtwo",
            _ => @"\ctexttwo"
        });
    }
}
