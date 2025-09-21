using System;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Navigation helpers for SheetComposer (sections with anchors, header/footer).
    /// </summary>
    public sealed partial class SheetComposer
    {
        /// <summary>
        /// Inserts a section header and a back-to-top link (explicit A1 link).
        /// </summary>
        public SheetComposer SectionWithAnchor(string text, string? anchorName = null, bool backToTopLink = true, string backToTopText = "↑ Top")
        {
            Section(text);
            if (backToTopLink)
            {
                try {
                    string topName = SanitizeName($"top_{Sheet.Name}");
                    Sheet.SetInternalLink(_row, 1, Sheet, "A1", backToTopText);
                    _row++;
                } catch { }
            }
            return this;
        }

        /// <summary>Configures header/footer content and images via a builder.</summary>
        public SheetComposer HeaderFooter(Action<HeaderFooterBuilder> configure)
        {
            if (configure == null) return this;
            var b = new HeaderFooterBuilder();
            configure(b);
            b.Apply(Sheet);
            return this;
        }

        /// <summary>Applies optional autofit operations and returns the composer.</summary>
        public SheetComposer Finish(bool autoFitColumns = true, bool autoFitRows = false)
        {
            if (autoFitColumns) Sheet.AutoFitColumns();
            if (autoFitRows) Sheet.AutoFitRows();
            return this;
        }
    }
}
