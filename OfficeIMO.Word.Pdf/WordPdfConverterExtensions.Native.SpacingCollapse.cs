using System.Collections.Generic;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private sealed class NativeSpacingCollapseFlow : INativePdfFlow {
            private readonly INativePdfFlow _inner;
            private double? _pendingSpacingAfter;

            public NativeSpacingCollapseFlow(INativePdfFlow inner) {
                _inner = inner;
            }

            public void PageBreak() {
                _inner.PageBreak();
                ResetSpacingCollapse();
            }

            public void Spacer(double height) {
                _inner.Spacer(height);
                ResetSpacingCollapse();
            }

            public void Bookmark(string name) => _inner.Bookmark(name);

            public void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) {
                _inner.HR(thickness, color, spacingBefore, spacingAfter, style);
                ResetSpacingCollapse();
            }

            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) {
                PdfCore.PdfParagraphStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = CollapseSpacingBefore(style.SpacingBefore);
                }

                _inner.Paragraph(build, align, defaultColor, collapsedStyle);
                _pendingSpacingAfter = style?.SpacingAfter;
            }

            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null) {
                PdfCore.PanelStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = CollapseSpacingBefore(style.SpacingBefore);
                }

                _inner.PanelParagraph(build, collapsedStyle, align, defaultColor);
                _pendingSpacingAfter = style?.SpacingAfter;
            }

            public void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents) {
                PdfCore.PdfHeadingStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = CollapseSpacingBefore(style.SpacingBefore);
                }

                _inner.Heading(level, text, align, color, collapsedStyle, linkUri, linkDestinationName, linkContents);
                _pendingSpacingAfter = style?.SpacingAfter;
            }

            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style) {
                PdfCore.PdfListStyle? collapsedStyle = CollapseListStyle(style);
                _inner.RichNumbered(items, align, color, startNumber, collapsedStyle);
                _pendingSpacingAfter = style?.SpacingAfter;
            }

            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style) {
                PdfCore.PdfListStyle? collapsedStyle = CollapseListStyle(style);
                _inner.RichBullets(items, align, color, collapsedStyle);
                _pendingSpacingAfter = style?.SpacingAfter;
            }

            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style = null) {
                _inner.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style);
                ResetSpacingCollapse();
            }

            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style = null) {
                _inner.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
                ResetSpacingCollapse();
            }

            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName = "Yes", PdfCore.PdfFormFieldStyle? style = null) {
                _inner.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
                ResetSpacingCollapse();
            }

            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
                _inner.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
                ResetSpacingCollapse();
            }

            public void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
                _inner.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
                ResetSpacingCollapse();
            }

            public void Canvas(Action<PdfCore.PdfPageCanvas> build) {
                _inner.Canvas(build);
                ResetSpacingCollapse();
            }

            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) {
                _inner.Table(rows, align, style);
                ResetSpacingCollapse();
            }

            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) {
                _inner.Image(bytes, width, height, align);
                ResetSpacingCollapse();
            }

            private PdfCore.PdfListStyle? CollapseListStyle(PdfCore.PdfListStyle? style) {
                if (style == null) {
                    return null;
                }

                PdfCore.PdfListStyle collapsedStyle = style.Clone();
                collapsedStyle.SpacingBefore = CollapseSpacingBefore(style.SpacingBefore);
                return collapsedStyle;
            }

            private double CollapseSpacingBefore(double spacingBefore) =>
                _pendingSpacingAfter.HasValue
                    ? Math.Max(0D, spacingBefore - _pendingSpacingAfter.Value)
                    : spacingBefore;

            private void ResetSpacingCollapse() => _pendingSpacingAfter = null;
        }
    }
}
