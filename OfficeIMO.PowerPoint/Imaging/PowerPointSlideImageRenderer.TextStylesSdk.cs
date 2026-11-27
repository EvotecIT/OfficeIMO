using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// What image export still reads through the SDK: the effective run style resolver, which takes the SDK list,
    /// master and table text styles and owner parts, and the table style members the HTML converter and Google Slides
    /// read as SDK elements. The rest of the renderer reads the parts through <see cref="OfficeXmlNode"/>.
    /// </summary>
    internal static partial class PowerPointSlideImageRenderer {
        private static IReadOnlyList<PowerPointEffectiveTextSegment> ResolveTextSegments(PowerPointTextRun run,
            PowerPointParagraph paragraph, PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveSegments(run, paragraph, textBox.TextBody?.ListStyle, textBox.MasterTextStyle);

        private static PowerPointEffectiveRunStyle ResolveParagraphStyle(PowerPointParagraph paragraph, PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveParagraph(paragraph, textBox.TextBody?.ListStyle, textBox.MasterTextStyle);

        private static PowerPointEffectiveRunStyle ResolveParagraphStyle(PowerPointParagraph paragraph, PowerPointTableCell cell) =>
            PowerPointEffectiveRunStyleResolver.ResolveParagraph(paragraph, cell.Cell.TextBody?.ListStyle,
                cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles?.OtherStyle);

        /// <summary>The run's inherited highlight in the text box's theme context, or null.</summary>
        private static OfficeColor? ResolveTextBoxHighlightColor(PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveTextColorElementForExport(
                PowerPointEffectiveRunStyleResolver.ResolveHighlightElement(run, paragraph, textBox.TextBody?.ListStyle, textBox.MasterTextStyle),
                textBox);

        /// <summary>The run's inherited <c>a:highlight</c> in a table cell, or none.</summary>
        private static OfficeXmlNode ResolveTableCellHighlightElement(PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTableCell cell) {
            OpenXmlCompositeElement? masterStyle = cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles?.OtherStyle;
            return OfficeXmlNode.FromSdk(PowerPointEffectiveRunStyleResolver.ResolveHighlightElement(run, paragraph,
                cell.Cell.TextBody?.ListStyle, masterStyle));
        }

        /// <summary>Resolves a theme font reference at the part that holds the run, its paragraph or the text box.</summary>
        private static string? ResolveThemeFontReference(string? typeface, PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTextBox textBox) {
            OpenXmlPartContainer? owner = run?.OwnerPart ?? paragraph?.OwnerPart ?? textBox.RelationshipOwnerPart;
            return PowerPointEffectiveRunStyleResolver.ResolveThemeFontReference(typeface, owner);
        }

        /// <summary>Resolves a theme font reference at the part that holds the run, its paragraph or the cell.</summary>
        private static string? ResolveThemeFontReference(string? typeface, PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTableCell cell) {
            OpenXmlPartContainer? owner = run?.OwnerPart ?? paragraph?.OwnerPart ?? cell.OwnerPart;
            return PowerPointEffectiveRunStyleResolver.ResolveThemeFontReference(typeface, owner);
        }

        /// <summary>The color scheme of a table's slide, or of the layout or master that holds it.</summary>
        private static OfficeXmlNode ResolveTableColorScheme(PowerPointTable table) =>
            table.OwnerSlide == null
                ? OfficeXmlNode.FromSdk(PowerPointEffectiveRunStyleResolver.ResolveColorScheme(table.RelationshipOwnerPart))
                : GetSlideColorScheme(table.OwnerSlide);

        /// <summary>The table's style as an SDK element, for the HTML converter.</summary>
        internal static A.TableStyleEntry? ResolveTableStyle(PowerPointTable table) =>
            ToSdkElement(ResolveTableStyleElement(table), static xml => new A.TableStyleEntry(xml));

        internal static OfficeColor ResolveTableCellFillColorForExport(
            PowerPointTable table, int row, int column, A.TableStyleEntry? tableStyle) =>
            ResolveTableCellFillColorForExport(table, row, column, OfficeXmlNode.FromSdk(tableStyle));

        internal static OfficeBorderBox ResolveTableCellBordersForExport(
            PowerPointTable table, int row, int column, A.TableStyleEntry? tableStyle) =>
            ResolveTableCellBordersForExport(table, row, column, OfficeXmlNode.FromSdk(tableStyle));

        /// <summary>The table text styles that apply to a cell, nearest first, as the effective run style resolver reads them.</summary>
        internal static IReadOnlyList<A.TableCellTextStyle> ResolveTableCellTextStylesForExport(
            PowerPointTable table,
            int row,
            int column) => ResolveTableCellTextStylesForExport(table, row, column, ResolveTableStyleElement(table));

        internal static IReadOnlyList<A.TableCellTextStyle> ResolveTableCellTextStylesForExport(
            PowerPointTable table,
            int row,
            int column,
            A.TableStyleEntry? tableStyle) => ResolveTableCellTextStylesForExport(table, row, column, OfficeXmlNode.FromSdk(tableStyle));

        private static IReadOnlyList<A.TableCellTextStyle> ResolveTableCellTextStylesForExport(
            PowerPointTable table,
            int row,
            int column,
            OfficeXmlNode tableStyle) => ResolveTableCellTextStyleElements(table, row, column, tableStyle)
                .Select(static style => ToSdkElement(style, static xml => new A.TableCellTextStyle(xml))!)
                .ToArray();

        // An element held in a tree (a package part or the built-in styles) is read into a detached SDK element.
        private static T? ToSdkElement<T>(OfficeXmlNode element, Func<string, T> parse) where T : OpenXmlElement =>
            element.IsNone ? null : element.SdkElement as T ?? parse(Encoding.UTF8.GetString(element.OuterXml()));
    }
}
