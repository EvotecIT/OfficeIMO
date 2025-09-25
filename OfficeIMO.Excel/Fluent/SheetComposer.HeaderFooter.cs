namespace OfficeIMO.Excel.Fluent {
    public sealed partial class SheetComposer {
        /// <summary>
        /// Convenience: sets a header logo from URL and optional header text using the fluent builder.
        /// </summary>
        public SheetComposer HeaderLogoUrl(string url, HeaderFooterPosition position = HeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null) {
            HeaderFooter(h => {
                if (!string.IsNullOrEmpty(leftText)) h.Left(leftText!);
                if (!string.IsNullOrEmpty(centerText)) h.Center(centerText!);
                if (!string.IsNullOrEmpty(rightText)) h.Right(rightText!);
                switch (position) {
                    case HeaderFooterPosition.Left: h.LeftImageUrl(url, widthPoints, heightPoints); break;
                    case HeaderFooterPosition.Center: h.CenterImageUrl(url, widthPoints, heightPoints); break;
                    default: h.RightImageUrl(url, widthPoints, heightPoints); break;
                }
            });
            return this;
        }

        /// <summary>
        /// Convenience: sets a footer logo from URL and optional footer text using the fluent builder.
        /// </summary>
        public SheetComposer FooterLogoUrl(string url, HeaderFooterPosition position = HeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null) {
            HeaderFooter(h => {
                if (!string.IsNullOrEmpty(leftText)) h.FooterLeft(leftText!);
                if (!string.IsNullOrEmpty(centerText)) h.FooterCenter(centerText!);
                if (!string.IsNullOrEmpty(rightText)) h.FooterRight(rightText!);
                switch (position) {
                    case HeaderFooterPosition.Left: h.FooterLeftImageUrl(url, widthPoints, heightPoints); break;
                    case HeaderFooterPosition.Center: h.FooterCenterImageUrl(url, widthPoints, heightPoints); break;
                    default: h.FooterRightImageUrl(url, widthPoints, heightPoints); break;
                }
            });
            return this;
        }
    }
}
