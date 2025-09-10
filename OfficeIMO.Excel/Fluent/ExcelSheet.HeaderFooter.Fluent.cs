using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel
{
    public partial class ExcelSheet
    {
        /// <summary>
        /// Fluent header/footer configuration using the same builder as SheetComposer.
        /// </summary>
        public void HeaderFooter(Action<OfficeIMO.Excel.Fluent.HeaderFooterBuilder> configure)
        {
            if (configure == null) return;
            var b = new OfficeIMO.Excel.Fluent.HeaderFooterBuilder();
            configure(b);
            b.Apply(this);
        }

        /// <summary>
        /// Convenience: sets a header logo from URL in the given position. Optional page text can be supplied.
        /// </summary>
        public void HeaderLogoUrl(string url, HeaderFooterPosition position = HeaderFooterPosition.Right,
                                  double? widthPoints = null, double? heightPoints = null,
                                  string? leftText = null, string? centerText = null, string? rightText = null)
        {
            SetHeaderFooter(leftText, centerText, rightText);
            SetHeaderImageUrl(position, url, widthPoints, heightPoints);
        }
    }
}
