using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static IReadOnlyList<PowerPointEffectiveTextSegment> ResolveTextSegments(PowerPointTextRun run,
            PowerPointParagraph paragraph, PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveSegments(run, paragraph,
                PowerPointEffectiveRunStyleResolver.TextBoxListStyle(textBox), PowerPointEffectiveRunStyleResolver.TextBoxMasterStyle(textBox));

        private static PowerPointEffectiveRunStyle ResolveParagraphStyle(PowerPointParagraph paragraph, PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveParagraph(paragraph,
                PowerPointEffectiveRunStyleResolver.TextBoxListStyle(textBox), PowerPointEffectiveRunStyleResolver.TextBoxMasterStyle(textBox));

        private static PowerPointEffectiveRunStyle ResolveParagraphStyle(PowerPointParagraph paragraph, PowerPointTableCell cell) =>
            PowerPointEffectiveRunStyleResolver.ResolveParagraph(paragraph, TableCellListStyle(cell), TableCellMasterStyle(cell));

        private static OfficeColor? ResolveTextBoxHighlightColor(PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveTextColorElementForExport(
                PowerPointEffectiveRunStyleResolver.ResolveHighlightElement(run, paragraph,
                    PowerPointEffectiveRunStyleResolver.TextBoxListStyle(textBox), PowerPointEffectiveRunStyleResolver.TextBoxMasterStyle(textBox)),
                textBox);

        private static OfficeXmlNode ResolveTableCellHighlightElement(PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTableCell cell) => PowerPointEffectiveRunStyleResolver.ResolveHighlightElement(run, paragraph,
                TableCellListStyle(cell), TableCellMasterStyle(cell));

        private static OfficeXmlNode TableCellListStyle(PowerPointTableCell cell) =>
            cell.Element.Child(D, "txBody").Child(D, "lstStyle");

        private static OfficeXmlNode TableCellMasterStyle(PowerPointTableCell cell) =>
            PowerPointEffectiveRunStyleResolver.MasterTextStyles(cell.SlidePartRef).Child(Pml, "otherStyle");

        private static string? ResolveThemeFontReference(string? typeface, PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTextBox textBox) =>
            PowerPointEffectiveRunStyleResolver.ResolveThemeFontReference(typeface, TextOwner(run, paragraph, textBox.OwnerPartRef));

        private static string? ResolveThemeFontReference(string? typeface, PowerPointTextRun? run, PowerPointParagraph? paragraph,
            PowerPointTableCell cell) =>
            PowerPointEffectiveRunStyleResolver.ResolveThemeFontReference(typeface, TextOwner(run, paragraph, cell.OwnerPartRef));

        private static PowerPointPartRef TextOwner(PowerPointTextRun? run, PowerPointParagraph? paragraph, PowerPointPartRef fallback) {
            PowerPointPartRef owner = run?.OwnerPartRef ?? default;
            if (owner.IsNone) owner = paragraph?.OwnerPartRef ?? default;
            return owner.IsNone ? fallback : owner;
        }

        private static OfficeXmlNode ResolveTableColorScheme(PowerPointTable table) =>
            table.OwnerSlide == null ? PowerPointThemeScope.For(table.OwnerPartRef).ColorScheme : GetSlideColorScheme(table.OwnerSlide);
    }
}
