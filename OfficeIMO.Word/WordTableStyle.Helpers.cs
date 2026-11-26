namespace OfficeIMO.Word;

// Provides helper methods for working with table styles.
internal static partial class WordTableStyles {
    private static readonly Lazy<Dictionary<string, WordTableStyle>> StylesById = new(() =>
        global::OfficeIMO.Internal.EnumCompat.GetValues<WordTableStyle>()
            .ToDictionary(GetStyleId, style => style, StringComparer.Ordinal));

    internal static bool TryGetStyle(string styleId, out WordTableStyle style) =>
        StylesById.Value.TryGetValue(styleId, out style);

    /// <summary>
    /// Converts a style name into its corresponding <see cref="WordTableStyle"/> value.
    /// </summary>
    /// <param name="style">The style name as defined by Microsoft Word.</param>
    /// <returns>The matching <see cref="WordTableStyle"/> enumeration value.</returns>
    public static WordTableStyle GetStyle(string style) {
        switch (style) {
            // Grid Tables - Line 1
            case "TableNormal": return WordTableStyle.TableNormal;
            case "TableGrid": return WordTableStyle.TableGrid;
            case WordThemeDefaults.TableStyleId: return WordTableStyle.OfficeIMOModern;
            case "PlainTable1": return WordTableStyle.PlainTable1;
            case "PlainTable2": return WordTableStyle.PlainTable2;
            case "PlainTable3": return WordTableStyle.PlainTable3;
            case "PlainTable4": return WordTableStyle.PlainTable4;
            case "PlainTable5": return WordTableStyle.PlainTable5;
            // Grid Tables - Line 1
            case "GridTable1Light": return WordTableStyle.GridTable1Light;
            case "GridTable1Light-Accent1": return WordTableStyle.GridTable1LightAccent1;
            case "GridTable1Light-Accent2": return WordTableStyle.GridTable1LightAccent2;
            case "GridTable1Light-Accent3": return WordTableStyle.GridTable1LightAccent3;
            case "GridTable1Light-Accent4": return WordTableStyle.GridTable1LightAccent4;
            case "GridTable1Light-Accent5": return WordTableStyle.GridTable1LightAccent5;
            case "GridTable1Light-Accent6": return WordTableStyle.GridTable1LightAccent6;
            //// Grid Tables - Line 2
            case "GridTable2": return WordTableStyle.GridTable2;
            case "GridTable2-Accent1": return WordTableStyle.GridTable2Accent1;
            case "GridTable2-Accent2": return WordTableStyle.GridTable2Accent2;
            case "GridTable2-Accent3": return WordTableStyle.GridTable2Accent3;
            case "GridTable2-Accent4": return WordTableStyle.GridTable2Accent4;
            case "GridTable2-Accent5": return WordTableStyle.GridTable2Accent5;
            case "GridTable2-Accent6": return WordTableStyle.GridTable2Accent6;
            //// Grid Tables - Line 3
            case "GridTable3": return WordTableStyle.GridTable3;
            case "GridTable3-Accent1": return WordTableStyle.GridTable3Accent1;
            case "GridTable3-Accent2": return WordTableStyle.GridTable3Accent2;
            case "GridTable3-Accent3": return WordTableStyle.GridTable3Accent3;
            case "GridTable3-Accent4": return WordTableStyle.GridTable3Accent4;
            case "GridTable3-Accent5": return WordTableStyle.GridTable3Accent5;
            case "GridTable3-Accent6": return WordTableStyle.GridTable3Accent6;
            //// Grid Tables - Line 4
            case "GridTable4": return WordTableStyle.GridTable4;
            case "GridTable4-Accent1": return WordTableStyle.GridTable4Accent1;
            case "GridTable4-Accent2": return WordTableStyle.GridTable4Accent2;
            case "GridTable4-Accent3": return WordTableStyle.GridTable4Accent3;
            case "GridTable4-Accent4": return WordTableStyle.GridTable4Accent4;
            case "GridTable4-Accent5": return WordTableStyle.GridTable4Accent5;
            case "GridTable4-Accent6": return WordTableStyle.GridTable4Accent6;
            //// Grid Tables - Line 5
            case "GridTable5Dark": return WordTableStyle.GridTable5Dark;
            case "GridTable5Dark-Accent1": return WordTableStyle.GridTable5DarkAccent1;
            case "GridTable5Dark-Accent2": return WordTableStyle.GridTable5DarkAccent2;
            case "GridTable5Dark-Accent3": return WordTableStyle.GridTable5DarkAccent3;
            case "GridTable5Dark-Accent4": return WordTableStyle.GridTable5DarkAccent4;
            case "GridTable5Dark-Accent5": return WordTableStyle.GridTable5DarkAccent5;
            case "GridTable5Dark-Accent6": return WordTableStyle.GridTable5DarkAccent6;
            //// Grid Tables - Line 6
            case "GridTable6Colorful": return WordTableStyle.GridTable6Colorful;
            case "GridTable6Colorful-Accent1": return WordTableStyle.GridTable6ColorfulAccent1;
            case "GridTable6Colorful-Accent2": return WordTableStyle.GridTable6ColorfulAccent2;
            case "GridTable6Colorful-Accent3": return WordTableStyle.GridTable6ColorfulAccent3;
            case "GridTable6Colorful-Accent4": return WordTableStyle.GridTable6ColorfulAccent4;
            case "GridTable6Colorful-Accent5": return WordTableStyle.GridTable6ColorfulAccent5;
            case "GridTable6Colorful-Accent6": return WordTableStyle.GridTable6ColorfulAccent6;
            //// Grid Tables - Line 7
            case "GridTable7Colorful": return WordTableStyle.GridTable7Colorful;
            case "GridTable7Colorful-Accent1": return WordTableStyle.GridTable7ColorfulAccent1;
            case "GridTable7Colorful-Accent2": return WordTableStyle.GridTable7ColorfulAccent2;
            case "GridTable7Colorful-Accent3": return WordTableStyle.GridTable7ColorfulAccent3;
            case "GridTable7Colorful-Accent4": return WordTableStyle.GridTable7ColorfulAccent4;
            case "GridTable7Colorful-Accent5": return WordTableStyle.GridTable7ColorfulAccent5;
            case "GridTable7Colorful-Accent6": return WordTableStyle.GridTable7ColorfulAccent6;
            //// Grid Tables - Line 8
            case "ListTable1Light": return WordTableStyle.ListTable1Light;
            case "ListTable1Light-Accent1": return WordTableStyle.ListTable1LightAccent1;
            case "ListTable1Light-Accent2": return WordTableStyle.ListTable1LightAccent2;
            case "ListTable1Light-Accent3": return WordTableStyle.ListTable1LightAccent3;
            case "ListTable1Light-Accent4": return WordTableStyle.ListTable1LightAccent4;
            case "ListTable1Light-Accent5": return WordTableStyle.ListTable1LightAccent5;
            case "ListTable1Light-Accent6": return WordTableStyle.ListTable1LightAccent6;
            //// List Tables - Line 9
            case "ListTable2": return WordTableStyle.ListTable2;
            case "ListTable2-Accent1": return WordTableStyle.ListTable2Accent1;
            case "ListTable2-Accent2": return WordTableStyle.ListTable2Accent2;
            case "ListTable2-Accent3": return WordTableStyle.ListTable2Accent3;
            case "ListTable2-Accent4": return WordTableStyle.ListTable2Accent4;
            case "ListTable2-Accent5": return WordTableStyle.ListTable2Accent5;
            case "ListTable2-Accent6": return WordTableStyle.ListTable2Accent6;

            //// List Tables - Line 10
            case "ListTable3": return WordTableStyle.ListTable3;
            case "ListTable3-Accent1": return WordTableStyle.ListTable3Accent1;
            case "ListTable3-Accent2": return WordTableStyle.ListTable3Accent2;
            case "ListTable3-Accent3": return WordTableStyle.ListTable3Accent3;
            case "ListTable3-Accent4": return WordTableStyle.ListTable3Accent4;
            case "ListTable3-Accent5": return WordTableStyle.ListTable3Accent5;
            case "ListTable3-Accent6": return WordTableStyle.ListTable3Accent6;

            //// List Tables - Line 11
            case "ListTable4": return WordTableStyle.ListTable4;
            case "ListTable4-Accent1": return WordTableStyle.ListTable4Accent1;
            case "ListTable4-Accent2": return WordTableStyle.ListTable4Accent2;
            case "ListTable4-Accent3": return WordTableStyle.ListTable4Accent3;
            case "ListTable4-Accent4": return WordTableStyle.ListTable4Accent4;
            case "ListTable4-Accent5": return WordTableStyle.ListTable4Accent5;
            case "ListTable4-Accent6": return WordTableStyle.ListTable4Accent6;

            //// List Tables - Line 12
            case "ListTable5Dark": return WordTableStyle.ListTable5Dark;
            case "ListTable5Dark-Accent1": return WordTableStyle.ListTable5DarkAccent1;
            case "ListTable5Dark-Accent2": return WordTableStyle.ListTable5DarkAccent2;
            case "ListTable5Dark-Accent3": return WordTableStyle.ListTable5DarkAccent3;
            case "ListTable5Dark-Accent4": return WordTableStyle.ListTable5DarkAccent4;
            case "ListTable5Dark-Accent5": return WordTableStyle.ListTable5DarkAccent5;
            case "ListTable5Dark-Accent6": return WordTableStyle.ListTable5DarkAccent6;

            //// List Tables - Line 13
            case "ListTable6Colorful": return WordTableStyle.ListTable6Colorful;
            case "ListTable6Colorful-Accent1": return WordTableStyle.ListTable6ColorfulAccent1;
            case "ListTable6Colorful-Accent2": return WordTableStyle.ListTable6ColorfulAccent2;
            case "ListTable6Colorful-Accent3": return WordTableStyle.ListTable6ColorfulAccent3;
            case "ListTable6Colorful-Accent4": return WordTableStyle.ListTable6ColorfulAccent4;
            case "ListTable6Colorful-Accent5": return WordTableStyle.ListTable6ColorfulAccent5;
            case "ListTable6Colorful-Accent6": return WordTableStyle.ListTable6ColorfulAccent6;

            //// List Tables - Line 14
            case "ListTable7Colorful": return WordTableStyle.ListTable7Colorful;
            case "ListTable7Colorful-Accent1": return WordTableStyle.ListTable7ColorfulAccent1;
            case "ListTable7Colorful-Accent2": return WordTableStyle.ListTable7ColorfulAccent2;
            case "ListTable7Colorful-Accent3": return WordTableStyle.ListTable7ColorfulAccent3;
            case "ListTable7Colorful-Accent4": return WordTableStyle.ListTable7ColorfulAccent4;
            case "ListTable7Colorful-Accent5": return WordTableStyle.ListTable7ColorfulAccent5;
            case "ListTable7Colorful-Accent6": return WordTableStyle.ListTable7ColorfulAccent6;
        }

        throw new ArgumentOutOfRangeException(nameof(style));
    }

    /// <summary>The style identifier of a built-in table style, as documents name it.</summary>
    internal static string GetStyleId(WordTableStyle style) {
        switch (style) {
            // Grid Tables - Line 1
            case WordTableStyle.OfficeIMOModern: return WordThemeDefaults.TableStyleId;
            case WordTableStyle.TableNormal: return "TableNormal";
            case WordTableStyle.TableGrid: return "TableGrid";
            case WordTableStyle.PlainTable1: return "PlainTable1";
            case WordTableStyle.PlainTable2: return "PlainTable2";
            case WordTableStyle.PlainTable3: return "PlainTable3";
            case WordTableStyle.PlainTable4: return "PlainTable4";
            case WordTableStyle.PlainTable5: return "PlainTable5";
            // Grid Tables - Line 1
            case WordTableStyle.GridTable1Light: return "GridTable1Light";
            case WordTableStyle.GridTable1LightAccent1: return "GridTable1Light-Accent1";
            case WordTableStyle.GridTable1LightAccent2: return "GridTable1Light-Accent2";
            case WordTableStyle.GridTable1LightAccent3: return "GridTable1Light-Accent3";
            case WordTableStyle.GridTable1LightAccent4: return "GridTable1Light-Accent4";
            case WordTableStyle.GridTable1LightAccent5: return "GridTable1Light-Accent5";
            case WordTableStyle.GridTable1LightAccent6: return "GridTable1Light-Accent6";
            // Grid Tables - Line 2
            case WordTableStyle.GridTable2: return "GridTable2";
            case WordTableStyle.GridTable2Accent1: return "GridTable2-Accent1";
            case WordTableStyle.GridTable2Accent2: return "GridTable2-Accent2";
            case WordTableStyle.GridTable2Accent3: return "GridTable2-Accent3";
            case WordTableStyle.GridTable2Accent4: return "GridTable2-Accent4";
            case WordTableStyle.GridTable2Accent5: return "GridTable2-Accent5";
            case WordTableStyle.GridTable2Accent6: return "GridTable2-Accent6";
            // Grid Tables - Line 3
            case WordTableStyle.GridTable3: return "GridTable3";
            case WordTableStyle.GridTable3Accent1: return "GridTable3-Accent1";
            case WordTableStyle.GridTable3Accent2: return "GridTable3-Accent2";
            case WordTableStyle.GridTable3Accent3: return "GridTable3-Accent3";
            case WordTableStyle.GridTable3Accent4: return "GridTable3-Accent4";
            case WordTableStyle.GridTable3Accent5: return "GridTable3-Accent5";
            case WordTableStyle.GridTable3Accent6: return "GridTable3-Accent6";
            // Grid Tables - Line 4
            case WordTableStyle.GridTable4: return "GridTable4";
            case WordTableStyle.GridTable4Accent1: return "GridTable4-Accent1";
            case WordTableStyle.GridTable4Accent2: return "GridTable4-Accent2";
            case WordTableStyle.GridTable4Accent3: return "GridTable4-Accent3";
            case WordTableStyle.GridTable4Accent4: return "GridTable4-Accent4";
            case WordTableStyle.GridTable4Accent5: return "GridTable4-Accent5";
            case WordTableStyle.GridTable4Accent6: return "GridTable4-Accent6";
            // Grid Tables - Line 5
            case WordTableStyle.GridTable5Dark: return "GridTable5Dark";
            case WordTableStyle.GridTable5DarkAccent1: return "GridTable5Dark-Accent1";
            case WordTableStyle.GridTable5DarkAccent2: return "GridTable5Dark-Accent2";
            case WordTableStyle.GridTable5DarkAccent3: return "GridTable5Dark-Accent3";
            case WordTableStyle.GridTable5DarkAccent4: return "GridTable5Dark-Accent4";
            case WordTableStyle.GridTable5DarkAccent5: return "GridTable5Dark-Accent5";
            case WordTableStyle.GridTable5DarkAccent6: return "GridTable5Dark-Accent6";
            // Grid Tables - Line 6
            case WordTableStyle.GridTable6Colorful: return "GridTable6Colorful";
            case WordTableStyle.GridTable6ColorfulAccent1: return "GridTable6Colorful-Accent1";
            case WordTableStyle.GridTable6ColorfulAccent2: return "GridTable6Colorful-Accent2";
            case WordTableStyle.GridTable6ColorfulAccent3: return "GridTable6Colorful-Accent3";
            case WordTableStyle.GridTable6ColorfulAccent4: return "GridTable6Colorful-Accent4";
            case WordTableStyle.GridTable6ColorfulAccent5: return "GridTable6Colorful-Accent5";
            case WordTableStyle.GridTable6ColorfulAccent6: return "GridTable6Colorful-Accent6";
            // Grid Tables - Line 7
            case WordTableStyle.GridTable7Colorful: return "GridTable7Colorful";
            case WordTableStyle.GridTable7ColorfulAccent1: return "GridTable7Colorful-Accent1";
            case WordTableStyle.GridTable7ColorfulAccent2: return "GridTable7Colorful-Accent2";
            case WordTableStyle.GridTable7ColorfulAccent3: return "GridTable7Colorful-Accent3";
            case WordTableStyle.GridTable7ColorfulAccent4: return "GridTable7Colorful-Accent4";
            case WordTableStyle.GridTable7ColorfulAccent5: return "GridTable7Colorful-Accent5";
            case WordTableStyle.GridTable7ColorfulAccent6: return "GridTable7Colorful-Accent6";
            // Grid Tables - Line 8
            case WordTableStyle.ListTable1Light: return "ListTable1Light";
            case WordTableStyle.ListTable1LightAccent1: return "ListTable1Light-Accent1";
            case WordTableStyle.ListTable1LightAccent2: return "ListTable1Light-Accent2";
            case WordTableStyle.ListTable1LightAccent3: return "ListTable1Light-Accent3";
            case WordTableStyle.ListTable1LightAccent4: return "ListTable1Light-Accent4";
            case WordTableStyle.ListTable1LightAccent5: return "ListTable1Light-Accent5";
            case WordTableStyle.ListTable1LightAccent6: return "ListTable1Light-Accent6";
            // Grid Tables - Line 9
            case WordTableStyle.ListTable2: return "ListTable2";
            case WordTableStyle.ListTable2Accent1: return "ListTable2-Accent1";
            case WordTableStyle.ListTable2Accent2: return "ListTable2-Accent2";
            case WordTableStyle.ListTable2Accent3: return "ListTable2-Accent3";
            case WordTableStyle.ListTable2Accent4: return "ListTable2-Accent4";
            case WordTableStyle.ListTable2Accent5: return "ListTable2-Accent5";
            case WordTableStyle.ListTable2Accent6: return "ListTable2-Accent6";
            // Grid Tables - Line 10
            case WordTableStyle.ListTable3: return "ListTable3";
            case WordTableStyle.ListTable3Accent1: return "ListTable3-Accent1";
            case WordTableStyle.ListTable3Accent2: return "ListTable3-Accent2";
            case WordTableStyle.ListTable3Accent3: return "ListTable3-Accent3";
            case WordTableStyle.ListTable3Accent4: return "ListTable3-Accent4";
            case WordTableStyle.ListTable3Accent5: return "ListTable3-Accent5";
            case WordTableStyle.ListTable3Accent6: return "ListTable3-Accent6";
            // Grid Tables - Line 11
            case WordTableStyle.ListTable4: return "ListTable4";
            case WordTableStyle.ListTable4Accent1: return "ListTable4-Accent1";
            case WordTableStyle.ListTable4Accent2: return "ListTable4-Accent2";
            case WordTableStyle.ListTable4Accent3: return "ListTable4-Accent3";
            case WordTableStyle.ListTable4Accent4: return "ListTable4-Accent4";
            case WordTableStyle.ListTable4Accent5: return "ListTable4-Accent5";
            case WordTableStyle.ListTable4Accent6: return "ListTable4-Accent6";
            // Grid Tables - Line 12
            case WordTableStyle.ListTable5Dark: return "ListTable5Dark";
            case WordTableStyle.ListTable5DarkAccent1: return "ListTable5Dark-Accent1";
            case WordTableStyle.ListTable5DarkAccent2: return "ListTable5Dark-Accent2";
            case WordTableStyle.ListTable5DarkAccent3: return "ListTable5Dark-Accent3";
            case WordTableStyle.ListTable5DarkAccent4: return "ListTable5Dark-Accent4";
            case WordTableStyle.ListTable5DarkAccent5: return "ListTable5Dark-Accent5";
            case WordTableStyle.ListTable5DarkAccent6: return "ListTable5Dark-Accent6";
            // Grid Tables - Line 13
            case WordTableStyle.ListTable6Colorful: return "ListTable6Colorful";
            case WordTableStyle.ListTable6ColorfulAccent1: return "ListTable6Colorful-Accent1";
            case WordTableStyle.ListTable6ColorfulAccent2: return "ListTable6Colorful-Accent2";
            case WordTableStyle.ListTable6ColorfulAccent3: return "ListTable6Colorful-Accent3";
            case WordTableStyle.ListTable6ColorfulAccent4: return "ListTable6Colorful-Accent4";
            case WordTableStyle.ListTable6ColorfulAccent5: return "ListTable6Colorful-Accent5";
            case WordTableStyle.ListTable6ColorfulAccent6: return "ListTable6Colorful-Accent6";
            // Grid Tables - Line 14
            case WordTableStyle.ListTable7Colorful: return "ListTable7Colorful";
            case WordTableStyle.ListTable7ColorfulAccent1: return "ListTable7Colorful-Accent1";
            case WordTableStyle.ListTable7ColorfulAccent2: return "ListTable7Colorful-Accent2";
            case WordTableStyle.ListTable7ColorfulAccent3: return "ListTable7Colorful-Accent3";
            case WordTableStyle.ListTable7ColorfulAccent4: return "ListTable7Colorful-Accent4";
            case WordTableStyle.ListTable7ColorfulAccent5: return "ListTable7Colorful-Accent5";
            case WordTableStyle.ListTable7ColorfulAccent6: return "ListTable7Colorful-Accent6";
        }

        throw new ArgumentOutOfRangeException(nameof(style));
    }
}
