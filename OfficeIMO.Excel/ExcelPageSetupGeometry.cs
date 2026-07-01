using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static class ExcelPageSetupGeometry {
        internal const double DefaultMarginLeftInches = 0.7D;
        internal const double DefaultMarginRightInches = 0.7D;
        internal const double DefaultMarginTopInches = 0.75D;
        internal const double DefaultMarginBottomInches = 0.75D;

        internal static OfficePageSize ResolvePageSize(ExcelSheetPageSetup? pageSetup, OfficePageSize fallbackPageSize) {
            OfficePageSize pageSize = pageSetup != null && TryResolvePageSize(pageSetup.PaperSize, out OfficePageSize resolvedPageSize)
                ? resolvedPageSize
                : fallbackPageSize;

            if (pageSetup?.Orientation == ExcelPageOrientation.Landscape) {
                return pageSize.Landscape();
            }

            return pageSize.Portrait();
        }

        internal static bool TryResolvePageSize(ExcelPaperSize? paperSize, out OfficePageSize pageSize) {
            switch (paperSize) {
                case ExcelPaperSize.Letter:
                case ExcelPaperSize.LetterSmall:
                    pageSize = OfficePageSizes.Letter;
                    return true;
                case ExcelPaperSize.Tabloid:
                    pageSize = OfficePageSizes.Tabloid;
                    return true;
                case ExcelPaperSize.Ledger:
                    pageSize = OfficePageSizes.Ledger;
                    return true;
                case ExcelPaperSize.Legal:
                    pageSize = OfficePageSizes.Legal;
                    return true;
                case ExcelPaperSize.Statement:
                    pageSize = OfficePageSizes.Statement;
                    return true;
                case ExcelPaperSize.Executive:
                    pageSize = OfficePageSizes.Executive;
                    return true;
                case ExcelPaperSize.A3:
                    pageSize = OfficePageSizes.A3;
                    return true;
                case ExcelPaperSize.A4:
                case ExcelPaperSize.A4Small:
                    pageSize = OfficePageSizes.A4;
                    return true;
                case ExcelPaperSize.A5:
                    pageSize = OfficePageSizes.A5;
                    return true;
                case ExcelPaperSize.B4Jis:
                    pageSize = OfficePageSizes.B4Jis;
                    return true;
                case ExcelPaperSize.B5Jis:
                    pageSize = OfficePageSizes.B5Jis;
                    return true;
                default:
                    pageSize = default;
                    return false;
            }
        }

        internal static bool HasFitToPageScale(ExcelSheetPageSetup? pageSetup) =>
            pageSetup?.FitToWidth.HasValue == true || pageSetup?.FitToHeight.HasValue == true;

        internal static bool HasUnsupportedFitToPageScale(ExcelSheetPageSetup? pageSetup) =>
            IsUnsupportedFitDimension(pageSetup?.FitToWidth) ||
            IsUnsupportedFitDimension(pageSetup?.FitToHeight);

        internal static double ResolveContentScale(
            ExcelSheetPageSetup pageSetup,
            int contentWidth,
            int contentHeight,
            double printableWidth,
            double printableHeight,
            double minimumScale,
            double maximumScale) {
            if (!HasFitToPageScale(pageSetup)) {
                return ClampScale((pageSetup.Scale ?? 100U) / 100D, minimumScale, maximumScale);
            }

            double scale = 1D;
            if (IsSupportedFitDimension(pageSetup.FitToWidth)) {
                scale = System.Math.Min(scale, printableWidth / System.Math.Max(1, contentWidth));
            }

            if (IsSupportedFitDimension(pageSetup.FitToHeight)) {
                scale = System.Math.Min(scale, printableHeight / System.Math.Max(1, contentHeight));
            }

            return ClampScale(scale, minimumScale, maximumScale);
        }

        internal static double ClampMargin(double margin, double pageSize) {
            if (double.IsNaN(margin) || double.IsInfinity(margin) || margin <= 0D) {
                return 0D;
            }

            return System.Math.Min(margin, System.Math.Max(0D, pageSize - 1D));
        }

        private static bool IsSupportedFitDimension(uint? value) =>
            value.HasValue && value.Value == 1U;

        private static bool IsUnsupportedFitDimension(uint? value) =>
            value.HasValue && value.Value > 1U;

        private static double ClampScale(double scale, double minimumScale, double maximumScale) =>
            System.Math.Max(minimumScale, System.Math.Min(maximumScale, scale));
    }
}
