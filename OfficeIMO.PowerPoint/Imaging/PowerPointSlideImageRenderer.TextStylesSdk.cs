using System.Text;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Table style adapters for the HTML and Google Slides converters.</summary>
    internal static partial class PowerPointSlideImageRenderer {
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
