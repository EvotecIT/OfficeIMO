using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Provides helper methods for working with table styles.
/// </summary>
public static partial class WordTableStyles {
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

    /// <summary>
    /// Creates a <see cref="TableStyle"/> element representing the specified table style.
    /// </summary>
    /// <param name="style">The <see cref="WordTableStyle"/> to convert.</param>
    /// <returns>A <see cref="TableStyle"/> element configured for the given style.</returns>
    public static TableStyle GetStyle(WordTableStyle style) {
        switch (style) {
            // Grid Tables - Line 1
            case WordTableStyle.TableNormal: return new TableStyle() { Val = "TableNormal" };
            case WordTableStyle.TableGrid: return new TableStyle() { Val = "TableGrid" };
            case WordTableStyle.PlainTable1: return new TableStyle() { Val = "PlainTable1" };
            case WordTableStyle.PlainTable2: return new TableStyle() { Val = "PlainTable2" };
            case WordTableStyle.PlainTable3: return new TableStyle() { Val = "PlainTable3" };
            case WordTableStyle.PlainTable4: return new TableStyle() { Val = "PlainTable4" };
            case WordTableStyle.PlainTable5: return new TableStyle() { Val = "PlainTable5" };
            // Grid Tables - Line 1
            case WordTableStyle.GridTable1Light: return new TableStyle() { Val = "GridTable1Light" };
            case WordTableStyle.GridTable1LightAccent1: return new TableStyle() { Val = "GridTable1Light-Accent1" };
            case WordTableStyle.GridTable1LightAccent2: return new TableStyle() { Val = "GridTable1Light-Accent2" };
            case WordTableStyle.GridTable1LightAccent3: return new TableStyle() { Val = "GridTable1Light-Accent3" };
            case WordTableStyle.GridTable1LightAccent4: return new TableStyle() { Val = "GridTable1Light-Accent4" };
            case WordTableStyle.GridTable1LightAccent5: return new TableStyle() { Val = "GridTable1Light-Accent5" };
            case WordTableStyle.GridTable1LightAccent6: return new TableStyle() { Val = "GridTable1Light-Accent6" };
            // Grid Tables - Line 2
            case WordTableStyle.GridTable2: return new TableStyle() { Val = "GridTable2" };
            case WordTableStyle.GridTable2Accent1: return new TableStyle() { Val = "GridTable2-Accent1" };
            case WordTableStyle.GridTable2Accent2: return new TableStyle() { Val = "GridTable2-Accent2" };
            case WordTableStyle.GridTable2Accent3: return new TableStyle() { Val = "GridTable2-Accent3" };
            case WordTableStyle.GridTable2Accent4: return new TableStyle() { Val = "GridTable2-Accent4" };
            case WordTableStyle.GridTable2Accent5: return new TableStyle() { Val = "GridTable2-Accent5" };
            case WordTableStyle.GridTable2Accent6: return new TableStyle() { Val = "GridTable2-Accent6" };
            // Grid Tables - Line 3
            case WordTableStyle.GridTable3: return new TableStyle() { Val = "GridTable3" };
            case WordTableStyle.GridTable3Accent1: return new TableStyle() { Val = "GridTable3-Accent1" };
            case WordTableStyle.GridTable3Accent2: return new TableStyle() { Val = "GridTable3-Accent2" };
            case WordTableStyle.GridTable3Accent3: return new TableStyle() { Val = "GridTable3-Accent3" };
            case WordTableStyle.GridTable3Accent4: return new TableStyle() { Val = "GridTable3-Accent4" };
            case WordTableStyle.GridTable3Accent5: return new TableStyle() { Val = "GridTable3-Accent5" };
            case WordTableStyle.GridTable3Accent6: return new TableStyle() { Val = "GridTable3-Accent6" };
            // Grid Tables - Line 4
            case WordTableStyle.GridTable4: return new TableStyle() { Val = "GridTable4" };
            case WordTableStyle.GridTable4Accent1: return new TableStyle() { Val = "GridTable4-Accent1" };
            case WordTableStyle.GridTable4Accent2: return new TableStyle() { Val = "GridTable4-Accent2" };
            case WordTableStyle.GridTable4Accent3: return new TableStyle() { Val = "GridTable4-Accent3" };
            case WordTableStyle.GridTable4Accent4: return new TableStyle() { Val = "GridTable4-Accent4" };
            case WordTableStyle.GridTable4Accent5: return new TableStyle() { Val = "GridTable4-Accent5" };
            case WordTableStyle.GridTable4Accent6: return new TableStyle() { Val = "GridTable4-Accent6" };
            // Grid Tables - Line 5
            case WordTableStyle.GridTable5Dark: return new TableStyle() { Val = "GridTable5Dark" };
            case WordTableStyle.GridTable5DarkAccent1: return new TableStyle() { Val = "GridTable5Dark-Accent1" };
            case WordTableStyle.GridTable5DarkAccent2: return new TableStyle() { Val = "GridTable5Dark-Accent2" };
            case WordTableStyle.GridTable5DarkAccent3: return new TableStyle() { Val = "GridTable5Dark-Accent3" };
            case WordTableStyle.GridTable5DarkAccent4: return new TableStyle() { Val = "GridTable5Dark-Accent4" };
            case WordTableStyle.GridTable5DarkAccent5: return new TableStyle() { Val = "GridTable5Dark-Accent5" };
            case WordTableStyle.GridTable5DarkAccent6: return new TableStyle() { Val = "GridTable5Dark-Accent6" };
            // Grid Tables - Line 6
            case WordTableStyle.GridTable6Colorful: return new TableStyle() { Val = "GridTable6Colorful" };
            case WordTableStyle.GridTable6ColorfulAccent1: return new TableStyle() { Val = "GridTable6Colorful-Accent1" };
            case WordTableStyle.GridTable6ColorfulAccent2: return new TableStyle() { Val = "GridTable6Colorful-Accent2" };
            case WordTableStyle.GridTable6ColorfulAccent3: return new TableStyle() { Val = "GridTable6Colorful-Accent3" };
            case WordTableStyle.GridTable6ColorfulAccent4: return new TableStyle() { Val = "GridTable6Colorful-Accent4" };
            case WordTableStyle.GridTable6ColorfulAccent5: return new TableStyle() { Val = "GridTable6Colorful-Accent5" };
            case WordTableStyle.GridTable6ColorfulAccent6: return new TableStyle() { Val = "GridTable6Colorful-Accent6" };
            // Grid Tables - Line 7
            case WordTableStyle.GridTable7Colorful: return new TableStyle() { Val = "GridTable7Colorful" };
            case WordTableStyle.GridTable7ColorfulAccent1: return new TableStyle() { Val = "GridTable7Colorful-Accent1" };
            case WordTableStyle.GridTable7ColorfulAccent2: return new TableStyle() { Val = "GridTable7Colorful-Accent2" };
            case WordTableStyle.GridTable7ColorfulAccent3: return new TableStyle() { Val = "GridTable7Colorful-Accent3" };
            case WordTableStyle.GridTable7ColorfulAccent4: return new TableStyle() { Val = "GridTable7Colorful-Accent4" };
            case WordTableStyle.GridTable7ColorfulAccent5: return new TableStyle() { Val = "GridTable7Colorful-Accent5" };
            case WordTableStyle.GridTable7ColorfulAccent6: return new TableStyle() { Val = "GridTable7Colorful-Accent6" };
            // Grid Tables - Line 8
            case WordTableStyle.ListTable1Light: return new TableStyle() { Val = "ListTable1Light" };
            case WordTableStyle.ListTable1LightAccent1: return new TableStyle() { Val = "ListTable1Light-Accent1" };
            case WordTableStyle.ListTable1LightAccent2: return new TableStyle() { Val = "ListTable1Light-Accent2" };
            case WordTableStyle.ListTable1LightAccent3: return new TableStyle() { Val = "ListTable1Light-Accent3" };
            case WordTableStyle.ListTable1LightAccent4: return new TableStyle() { Val = "ListTable1Light-Accent4" };
            case WordTableStyle.ListTable1LightAccent5: return new TableStyle() { Val = "ListTable1Light-Accent5" };
            case WordTableStyle.ListTable1LightAccent6: return new TableStyle() { Val = "ListTable1Light-Accent6" };
            // Grid Tables - Line 9
            case WordTableStyle.ListTable2: return new TableStyle() { Val = "ListTable2" };
            case WordTableStyle.ListTable2Accent1: return new TableStyle() { Val = "ListTable2-Accent1" };
            case WordTableStyle.ListTable2Accent2: return new TableStyle() { Val = "ListTable2-Accent2" };
            case WordTableStyle.ListTable2Accent3: return new TableStyle() { Val = "ListTable2-Accent3" };
            case WordTableStyle.ListTable2Accent4: return new TableStyle() { Val = "ListTable2-Accent4" };
            case WordTableStyle.ListTable2Accent5: return new TableStyle() { Val = "ListTable2-Accent5" };
            case WordTableStyle.ListTable2Accent6: return new TableStyle() { Val = "ListTable2-Accent6" };
            // Grid Tables - Line 10
            case WordTableStyle.ListTable3: return new TableStyle() { Val = "ListTable3" };
            case WordTableStyle.ListTable3Accent1: return new TableStyle() { Val = "ListTable3-Accent1" };
            case WordTableStyle.ListTable3Accent2: return new TableStyle() { Val = "ListTable3-Accent2" };
            case WordTableStyle.ListTable3Accent3: return new TableStyle() { Val = "ListTable3-Accent3" };
            case WordTableStyle.ListTable3Accent4: return new TableStyle() { Val = "ListTable3-Accent4" };
            case WordTableStyle.ListTable3Accent5: return new TableStyle() { Val = "ListTable3-Accent5" };
            case WordTableStyle.ListTable3Accent6: return new TableStyle() { Val = "ListTable3-Accent6" };
            // Grid Tables - Line 11
            case WordTableStyle.ListTable4: return new TableStyle() { Val = "ListTable4" };
            case WordTableStyle.ListTable4Accent1: return new TableStyle() { Val = "ListTable4-Accent1" };
            case WordTableStyle.ListTable4Accent2: return new TableStyle() { Val = "ListTable4-Accent2" };
            case WordTableStyle.ListTable4Accent3: return new TableStyle() { Val = "ListTable4-Accent3" };
            case WordTableStyle.ListTable4Accent4: return new TableStyle() { Val = "ListTable4-Accent4" };
            case WordTableStyle.ListTable4Accent5: return new TableStyle() { Val = "ListTable4-Accent5" };
            case WordTableStyle.ListTable4Accent6: return new TableStyle() { Val = "ListTable4-Accent6" };
            // Grid Tables - Line 12
            case WordTableStyle.ListTable5Dark: return new TableStyle() { Val = "ListTable5Dark" };
            case WordTableStyle.ListTable5DarkAccent1: return new TableStyle() { Val = "ListTable5Dark-Accent1" };
            case WordTableStyle.ListTable5DarkAccent2: return new TableStyle() { Val = "ListTable5Dark-Accent2" };
            case WordTableStyle.ListTable5DarkAccent3: return new TableStyle() { Val = "ListTable5Dark-Accent3" };
            case WordTableStyle.ListTable5DarkAccent4: return new TableStyle() { Val = "ListTable5Dark-Accent4" };
            case WordTableStyle.ListTable5DarkAccent5: return new TableStyle() { Val = "ListTable5Dark-Accent5" };
            case WordTableStyle.ListTable5DarkAccent6: return new TableStyle() { Val = "ListTable5Dark-Accent6" };
            // Grid Tables - Line 13
            case WordTableStyle.ListTable6Colorful: return new TableStyle() { Val = "ListTable6Colorful" };
            case WordTableStyle.ListTable6ColorfulAccent1: return new TableStyle() { Val = "ListTable6Colorful-Accent1" };
            case WordTableStyle.ListTable6ColorfulAccent2: return new TableStyle() { Val = "ListTable6Colorful-Accent2" };
            case WordTableStyle.ListTable6ColorfulAccent3: return new TableStyle() { Val = "ListTable6Colorful-Accent3" };
            case WordTableStyle.ListTable6ColorfulAccent4: return new TableStyle() { Val = "ListTable6Colorful-Accent4" };
            case WordTableStyle.ListTable6ColorfulAccent5: return new TableStyle() { Val = "ListTable6Colorful-Accent5" };
            case WordTableStyle.ListTable6ColorfulAccent6: return new TableStyle() { Val = "ListTable6Colorful-Accent6" };
            // Grid Tables - Line 14
            case WordTableStyle.ListTable7Colorful: return new TableStyle() { Val = "ListTable7Colorful" };
            case WordTableStyle.ListTable7ColorfulAccent1: return new TableStyle() { Val = "ListTable7Colorful-Accent1" };
            case WordTableStyle.ListTable7ColorfulAccent2: return new TableStyle() { Val = "ListTable7Colorful-Accent2" };
            case WordTableStyle.ListTable7ColorfulAccent3: return new TableStyle() { Val = "ListTable7Colorful-Accent3" };
            case WordTableStyle.ListTable7ColorfulAccent4: return new TableStyle() { Val = "ListTable7Colorful-Accent4" };
            case WordTableStyle.ListTable7ColorfulAccent5: return new TableStyle() { Val = "ListTable7Colorful-Accent5" };
            case WordTableStyle.ListTable7ColorfulAccent6: return new TableStyle() { Val = "ListTable7Colorful-Accent6" };
        }

        throw new ArgumentOutOfRangeException(nameof(style));
    }

    /// <summary>
    /// Verifies whether table style is available in document or not
    /// </summary>
    /// <param name="styles"></param>
    /// <param name="style"></param>
    /// <returns></returns>
    internal static bool IsAvailableStyle(Styles styles, WordTableStyle style) {
        var listCurrentStyles = styles.OfType<Style>().ToList();
        // Compare against style ID from the style definition to avoid duplicate styles (#85)
        var styleDefinition = GetStyleDefinition(style);
        var styleIdValue = styleDefinition.StyleId?.Value;
        if (styleIdValue == null) {
            return false;
        }
        foreach (var currentStyle in listCurrentStyles) {
            if (currentStyle.StyleId == styleIdValue) {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Gets the underlying <see cref="Style"/> definition for a given table style.
    /// </summary>
    /// <param name="style">The style to retrieve.</param>
    /// <returns>The <see cref="Style"/> definition that represents the provided enumeration value.</returns>
    public static Style GetStyleDefinition(WordTableStyle style) {
        switch (style) {
            // Grid Tables - Line 1
            case WordTableStyle.TableNormal: return StyleTableNormal;
            case WordTableStyle.TableGrid: return StyleTableGrid;
            case WordTableStyle.PlainTable1: return StylePlainTable1;
            case WordTableStyle.PlainTable2: return StylePlainTable2;
            case WordTableStyle.PlainTable3: return StylePlainTable3;
            case WordTableStyle.PlainTable4: return StylePlainTable4;
            case WordTableStyle.PlainTable5: return StylePlainTable5;
            // Grid Tables - Line 1
            case WordTableStyle.GridTable1Light: return StyleGridTable1Light;
            case WordTableStyle.GridTable1LightAccent1: return StyleGridTable1LightAccent1;
            case WordTableStyle.GridTable1LightAccent2: return StyleGridTable1LightAccent2;
            case WordTableStyle.GridTable1LightAccent3: return StyleGridTable1LightAccent3;
            case WordTableStyle.GridTable1LightAccent4: return StyleGridTable1LightAccent4;
            case WordTableStyle.GridTable1LightAccent5: return StyleGridTable1LightAccent5;
            case WordTableStyle.GridTable1LightAccent6: return StyleGridTable1LightAccent6;
            // Grid Tables - Line 2
            case WordTableStyle.GridTable2: return StyleGridTable2;
            case WordTableStyle.GridTable2Accent1: return StyleGridTable2Accent1;
            case WordTableStyle.GridTable2Accent2: return StyleGridTable2Accent2;
            case WordTableStyle.GridTable2Accent3: return StyleGridTable2Accent3;
            case WordTableStyle.GridTable2Accent4: return StyleGridTable2Accent4;
            case WordTableStyle.GridTable2Accent5: return StyleGridTable2Accent5;
            case WordTableStyle.GridTable2Accent6: return StyleGridTable2Accent6;
            // Grid Tables - Line 3
            case WordTableStyle.GridTable3: return StyleGridTable3;
            case WordTableStyle.GridTable3Accent1: return StyleGridTable3Accent1;
            case WordTableStyle.GridTable3Accent2: return StyleGridTable3Accent2;
            case WordTableStyle.GridTable3Accent3: return StyleGridTable3Accent3;
            case WordTableStyle.GridTable3Accent4: return StyleGridTable3Accent4;
            case WordTableStyle.GridTable3Accent5: return StyleGridTable3Accent5;
            case WordTableStyle.GridTable3Accent6: return StyleGridTable3Accent6;
            // Grid Tables - Line 4
            case WordTableStyle.GridTable4: return StyleGridTable4;
            case WordTableStyle.GridTable4Accent1: return StyleGridTable4Accent1;
            case WordTableStyle.GridTable4Accent2: return StyleGridTable4Accent2;
            case WordTableStyle.GridTable4Accent3: return StyleGridTable4Accent3;
            case WordTableStyle.GridTable4Accent4: return StyleGridTable4Accent4;
            case WordTableStyle.GridTable4Accent5: return StyleGridTable4Accent5;
            case WordTableStyle.GridTable4Accent6: return StyleGridTable4Accent6;
            // Grid Tables - Line 5
            case WordTableStyle.GridTable5Dark: return StyleGridTable5Dark;
            case WordTableStyle.GridTable5DarkAccent1: return StyleGridTable5DarkAccent1;
            case WordTableStyle.GridTable5DarkAccent2: return StyleGridTable5DarkAccent2;
            case WordTableStyle.GridTable5DarkAccent3: return StyleGridTable5DarkAccent3;
            case WordTableStyle.GridTable5DarkAccent4: return StyleGridTable5DarkAccent4;
            case WordTableStyle.GridTable5DarkAccent5: return StyleGridTable5DarkAccent5;
            case WordTableStyle.GridTable5DarkAccent6: return StyleGridTable5DarkAccent6;
            // Grid Tables - Line 6
            case WordTableStyle.GridTable6Colorful: return StyleGridTable6Colorful;
            case WordTableStyle.GridTable6ColorfulAccent1: return StyleGridTable6ColorfulAccent1;
            case WordTableStyle.GridTable6ColorfulAccent2: return StyleGridTable6ColorfulAccent2;
            case WordTableStyle.GridTable6ColorfulAccent3: return StyleGridTable6ColorfulAccent3;
            case WordTableStyle.GridTable6ColorfulAccent4: return StyleGridTable6ColorfulAccent4;
            case WordTableStyle.GridTable6ColorfulAccent5: return StyleGridTable6ColorfulAccent5;
            case WordTableStyle.GridTable6ColorfulAccent6: return StyleGridTable6ColorfulAccent6;
            // Grid Tables - Line 7
            case WordTableStyle.GridTable7Colorful: return StyleGridTable7Colorful;
            case WordTableStyle.GridTable7ColorfulAccent1: return StyleGridTable7ColorfulAccent1;
            case WordTableStyle.GridTable7ColorfulAccent2: return StyleGridTable7ColorfulAccent2;
            case WordTableStyle.GridTable7ColorfulAccent3: return StyleGridTable7ColorfulAccent3;
            case WordTableStyle.GridTable7ColorfulAccent4: return StyleGridTable7ColorfulAccent4;
            case WordTableStyle.GridTable7ColorfulAccent5: return StyleGridTable7ColorfulAccent5;
            case WordTableStyle.GridTable7ColorfulAccent6: return StyleGridTable7ColorfulAccent6;
            // Grid Tables - Line 8
            case WordTableStyle.ListTable1Light: return StyleListTable1Light;
            case WordTableStyle.ListTable1LightAccent1: return StyleListTable1LightAccent1;
            case WordTableStyle.ListTable1LightAccent2: return StyleListTable1LightAccent2;
            case WordTableStyle.ListTable1LightAccent3: return StyleListTable1LightAccent3;
            case WordTableStyle.ListTable1LightAccent4: return StyleListTable1LightAccent4;
            case WordTableStyle.ListTable1LightAccent5: return StyleListTable1LightAccent5;
            case WordTableStyle.ListTable1LightAccent6: return StyleListTable1LightAccent6;
            // Grid Tables - Line 9
            case WordTableStyle.ListTable2: return StyleListTable2;
            case WordTableStyle.ListTable2Accent1: return StyleListTable2Accent1;
            case WordTableStyle.ListTable2Accent2: return StyleListTable2Accent2;
            case WordTableStyle.ListTable2Accent3: return StyleListTable2Accent3;
            case WordTableStyle.ListTable2Accent4: return StyleListTable2Accent4;
            case WordTableStyle.ListTable2Accent5: return StyleListTable2Accent5;
            case WordTableStyle.ListTable2Accent6: return StyleListTable2Accent6;
            // Grid Tables - Line 10
            case WordTableStyle.ListTable3: return StyleListTable3;
            case WordTableStyle.ListTable3Accent1: return StyleListTable3Accent1;
            case WordTableStyle.ListTable3Accent2: return StyleListTable3Accent2;
            case WordTableStyle.ListTable3Accent3: return StyleListTable3Accent3;
            case WordTableStyle.ListTable3Accent4: return StyleListTable3Accent4;
            case WordTableStyle.ListTable3Accent5: return StyleListTable3Accent5;
            case WordTableStyle.ListTable3Accent6: return StyleListTable3Accent6;
            // Grid Tables - Line 11
            case WordTableStyle.ListTable4: return StyleListTable4;
            case WordTableStyle.ListTable4Accent1: return StyleListTable4Accent1;
            case WordTableStyle.ListTable4Accent2: return StyleListTable4Accent2;
            case WordTableStyle.ListTable4Accent3: return StyleListTable4Accent3;
            case WordTableStyle.ListTable4Accent4: return StyleListTable4Accent4;
            case WordTableStyle.ListTable4Accent5: return StyleListTable4Accent5;
            case WordTableStyle.ListTable4Accent6: return StyleListTable4Accent6;
            // Grid Tables - Line 12
            case WordTableStyle.ListTable5Dark: return StyleListTable5Dark;
            case WordTableStyle.ListTable5DarkAccent1: return StyleListTable5DarkAccent1;
            case WordTableStyle.ListTable5DarkAccent2: return StyleListTable5DarkAccent2;
            case WordTableStyle.ListTable5DarkAccent3: return StyleListTable5DarkAccent3;
            case WordTableStyle.ListTable5DarkAccent4: return StyleListTable5DarkAccent4;
            case WordTableStyle.ListTable5DarkAccent5: return StyleListTable5DarkAccent5;
            case WordTableStyle.ListTable5DarkAccent6: return StyleListTable5DarkAccent6;
            // Grid Tables - Line 13
            case WordTableStyle.ListTable6Colorful: return StyleListTable6Colorful;
            case WordTableStyle.ListTable6ColorfulAccent1: return StyleListTable6ColorfulAccent1;
            case WordTableStyle.ListTable6ColorfulAccent2: return StyleListTable6ColorfulAccent2;
            case WordTableStyle.ListTable6ColorfulAccent3: return StyleListTable6ColorfulAccent3;
            case WordTableStyle.ListTable6ColorfulAccent4: return StyleListTable6ColorfulAccent4;
            case WordTableStyle.ListTable6ColorfulAccent5: return StyleListTable6ColorfulAccent5;
            case WordTableStyle.ListTable6ColorfulAccent6: return StyleListTable6ColorfulAccent6;
            // Grid Tables - Line 14
            case WordTableStyle.ListTable7Colorful: return StyleListTable7Colorful;
            case WordTableStyle.ListTable7ColorfulAccent1: return StyleListTable7ColorfulAccent1;
            case WordTableStyle.ListTable7ColorfulAccent2: return StyleListTable7ColorfulAccent2;
            case WordTableStyle.ListTable7ColorfulAccent3: return StyleListTable7ColorfulAccent3;
            case WordTableStyle.ListTable7ColorfulAccent4: return StyleListTable7ColorfulAccent4;
            case WordTableStyle.ListTable7ColorfulAccent5: return StyleListTable7ColorfulAccent5;
            case WordTableStyle.ListTable7ColorfulAccent6: return StyleListTable7ColorfulAccent6;
        }

        throw new ArgumentOutOfRangeException(nameof(style));
    }

}
