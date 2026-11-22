using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private sealed class NativeSpacingCollapseFlow : INativePdfFlow {
            public void ContinuedFlow(Action<PdfCore.PdfContentBuilder> build, Action<PdfCore.PdfContentBuilder>? separator,
                Action<PdfCore.PdfContentBuilder>? continuationSeparator, Action<PdfCore.PdfContentBuilder>? continuationNotice,
                bool reserveFinalNotice, bool allowContinuationSeparatorInBottomMargin) {
                _inner.ContinuedFlow(build, separator, continuationSeparator, continuationNotice, reserveFinalNotice,
                    allowContinuationSeparatorInBottomMargin);
                ResetSpacingCollapse();
            }
            public PdfCore.PdfFootnoteDefinition DefineFootnote(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfFootnoteOptions options) => _inner.DefineFootnote(build, options);
            public bool SupportsPositionedTables => _inner.SupportsPositionedTables;
            private readonly INativePdfFlow _inner;
            private double? _pendingSpacingAfter;
            private bool _preserveInitialSpacingBefore;
            private double _preservedSpacingAfter;

            public NativeSpacingCollapseFlow(INativePdfFlow inner, bool preserveInitialSpacingBefore = false) {
                _inner = inner;
                _preserveInitialSpacingBefore = preserveInitialSpacingBefore;
            }

            // Contextual spacing removes the rendered trailing space, but Word
            // still subtracts that declared value from the next paragraph's
            // before spacing when collapsing their shared boundary.
            internal void RetainSuppressedSpacingAfter(double? declaredSpacingAfter) =>
                _pendingSpacingAfter = declaredSpacingAfter;

            public PdfCore.PdfPageSize PageSize => _inner.PageSize;

            /// <summary>The trailing space after of the last block, while the next block can still collapse with it.</summary>
            internal double? PendingSpacingAfter => _pendingSpacingAfter;

            // Word keeps the space before of the first paragraph on a page that starts a
            // new section (or follows a legacy page-break-before), reduced by the space
            // after that ended the previous page, instead of dropping it at the page top.
            internal void PreserveNextSpacingBefore(double? previousSpacingAfter) {
                _pendingSpacingAfter = null;
                _preserveInitialSpacingBefore = true;
                _preservedSpacingAfter = Math.Max(0D, previousSpacingAfter ?? 0D);
            }

            internal void RestorePendingSpacingAfter(double? spacingAfter) => _pendingSpacingAfter = spacingAfter;

            // Out-of-flow content, such as notes rendered at the end of a section, must neither
            // read nor replace the body's collapse state.
            internal (double? Pending, bool Preserve, double Carried) CaptureState() =>
                (_pendingSpacingAfter, _preserveInitialSpacingBefore, _preservedSpacingAfter);

            internal void RestoreState((double? Pending, bool Preserve, double Carried) state) {
                _pendingSpacingAfter = state.Pending;
                _preserveInitialSpacingBefore = state.Preserve;
                _preservedSpacingAfter = state.Carried;
            }

            /// <summary>Drops a carried section-start spacing that another flow has already used.</summary>
            internal void ClearPreservedSpacingBefore() {
                _preserveInitialSpacingBefore = false;
                _preservedSpacingAfter = 0D;
            }

            internal void PageBreakPreservingSpacingBefore() {
                double? previousSpacingAfter = _pendingSpacingAfter;
                _inner.PageBreak();
                PreserveNextSpacingBefore(previousSpacingAfter);
            }

            public void PageBreak(bool preserveEmptyPage = false) {
                _inner.PageBreak(preserveEmptyPage);
                ResetSpacingCollapse();
            }

            public void ColumnBreak() {
                _inner.ColumnBreak();
                ResetSpacingCollapse();
            }

            public void Spacer(double height) {
                _inner.Spacer(height);
                ResetSpacingCollapse();
            }

            public void ParagraphSpacingBefore(double height) =>
                _inner.ParagraphSpacingBefore(CollapseSpacingBefore(height));

            public void ParagraphSpacingAfter(double height) {
                _inner.ParagraphSpacingAfter(height);
                _pendingSpacingAfter = height;
            }

            public void Bookmark(string name) => _inner.Bookmark(name);
            public void TableOfContents(PdfCore.PdfTableOfContentsOptions options) { _inner.TableOfContents(options); ResetSpacingCollapse(); }

            public void HR(double? thickness = null, global::OfficeIMO.OfficeColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) {
                _inner.HR(thickness, color, spacingBefore, spacingAfter, style);
                ResetSpacingCollapse();
            }

            public void Paragraph(Action<PdfCore.PdfParagraph> build, global::OfficeIMO.OfficeAlignment align = global::OfficeIMO.OfficeAlignment.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) {
                PdfCore.PdfParagraphStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(CollapseSpacingBefore(style.SpacingBefore.ToPoints()));
                }

                _inner.Paragraph(build, align, defaultColor, collapsedStyle);
                _pendingSpacingAfter = style?.SpacingAfter?.ToPoints();
            }

            public void Panel(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfPanelStyle? style = null) {
                PdfCore.PdfPanelStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(CollapseSpacingBefore(style.SpacingBefore.ToPoints()));
                }

                _inner.Panel(build, collapsedStyle);
                _pendingSpacingAfter = (style?.SpacingAfter.ToPoints());
            }

            public void PanelParagraph(Action<PdfCore.PdfParagraph> build, PdfCore.PdfPanelStyle? style = null, global::OfficeIMO.OfficeAlignment align = global::OfficeIMO.OfficeAlignment.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? paragraphStyle = null) {
                PdfCore.PdfPanelStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(CollapseSpacingBefore(style.SpacingBefore.ToPoints()));
                }

                _inner.PanelParagraph(build, collapsedStyle, align, defaultColor, paragraphStyle);
                _pendingSpacingAfter = paragraphStyle?.SpacingAfter?.ToPoints() ?? (style?.SpacingAfter.ToPoints());
            }

            public void Heading(int level, string text, Action<PdfCore.PdfParagraph> build, global::OfficeIMO.OfficeAlignment align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfHeadingStyle? style) {
                PdfCore.PdfHeadingStyle? collapsedStyle = style;
                if (style != null) {
                    collapsedStyle = style.Clone();
                    collapsedStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(CollapseSpacingBefore(style.SpacingBefore.ToPoints()));
                }

                _inner.Heading(level, text, build, align, color, collapsedStyle);
                _pendingSpacingAfter = (style?.SpacingAfter?.ToPoints());
            }

            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, global::OfficeIMO.OfficeAlignment align, global::OfficeIMO.OfficeColor? color, int startNumber, PdfCore.PdfListStyle? style) {
                PdfCore.PdfListStyle? collapsedStyle = CollapseListStyle(style);
                _inner.RichNumbered(items, align, color, startNumber, collapsedStyle);
                _pendingSpacingAfter = ((style?.SpacingAfter)?.ToPoints());
            }

            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, global::OfficeIMO.OfficeAlignment align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfListStyle? style) {
                PdfCore.PdfListStyle? collapsedStyle = CollapseListStyle(style);
                _inner.RichBullets(items, align, color, collapsedStyle);
                _pendingSpacingAfter = ((style?.SpacingAfter)?.ToPoints());
            }

            public void TextField(string name, double width, double height, string value, global::OfficeIMO.OfficeAlignment align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style = null) {
                _inner.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style);
                ResetSpacingCollapse();
            }

            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, global::OfficeIMO.OfficeAlignment align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style = null) {
                _inner.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
                ResetSpacingCollapse();
            }

            public void CheckBox(string name, bool isChecked, double size, global::OfficeIMO.OfficeAlignment align, double spacingBefore, double spacingAfter, string checkedValueName = "Yes", PdfCore.PdfFormFieldStyle? style = null) {
                _inner.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
                ResetSpacingCollapse();
            }

            public void Shape(OfficeShape shape, global::OfficeIMO.OfficeAlignment? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
                _inner.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
                ResetSpacingCollapse();
            }

            public void Drawing(OfficeDrawing drawing, global::OfficeIMO.OfficeAlignment? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
                _inner.Drawing(drawing, align, spacingBefore.HasValue ? CollapseSpacingBefore(spacingBefore.Value) : null, spacingAfter, style, linkUri, linkContents);
                _pendingSpacingAfter = spacingAfter;
            }

            public IReadOnlyList<PdfCore.IPdfBlock> CaptureBlocks(Action<PdfCore.PdfContentBuilder> build) => _inner.CaptureBlocks(build);
            public void Canvas(Action<PdfCore.PdfPageCanvas> build) {
                _inner.Canvas(build);
                ResetSpacingCollapse();
            }

            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, global::OfficeIMO.OfficeAlignment align, PdfCore.PdfTableStyle? style) {
                _inner.Table(rows, align, style);
                ResetSpacingCollapse();
            }

            public void Image(byte[] bytes, double width, double height, global::OfficeIMO.OfficeAlignment? align = null, PdfCore.PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
                _inner.Image(bytes, width, height, align, style, linkUri, linkContents);
                ResetSpacingCollapse();
            }

            private PdfCore.PdfListStyle? CollapseListStyle(PdfCore.PdfListStyle? style) {
                if (style == null) {
                    return null;
                }

                PdfCore.PdfListStyle collapsedStyle = style.Clone();
                collapsedStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(CollapseSpacingBefore(style.SpacingBefore.ToPoints()));
                return collapsedStyle;
            }

            private double CollapseSpacingBefore(double spacingBefore) {
                // Running Word stories preserve authored initial spacing. An
                // explicit spacer avoids generic page-start suppression while
                // keeping measurement and rendering on the same flow path.
                bool preserve = _preserveInitialSpacingBefore;
                double preservedSpacingAfter = _preservedSpacingAfter;
                _preserveInitialSpacingBefore = false;
                _preservedSpacingAfter = 0D;
                if (preserve) {
                    double preservedSpacingBefore = spacingBefore - preservedSpacingAfter;
                    if (preservedSpacingBefore > 0D) _inner.Spacer(preservedSpacingBefore);
                    return 0D;
                }
                return _pendingSpacingAfter.HasValue
                    ? Math.Max(0D, spacingBefore - _pendingSpacingAfter.Value)
                    : spacingBefore;
            }

            private void ResetSpacingCollapse() {
                _pendingSpacingAfter = null;
                _preserveInitialSpacingBefore = false;
                _preservedSpacingAfter = 0D;
            }
        }
    }
}
