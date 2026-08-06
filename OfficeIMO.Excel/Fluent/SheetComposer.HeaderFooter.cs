using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel.Fluent {
    public sealed partial class SheetComposer {
        /// <summary>
        /// Asynchronously sets a header logo from a URL and optional header text.
        /// </summary>
        public async Task<SheetComposer> HeaderLogoFromUrlAsync(string url, ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            HeaderFooter(h => {
                if (!string.IsNullOrEmpty(leftText)) h.Left(leftText!);
                if (!string.IsNullOrEmpty(centerText)) h.Center(centerText!);
                if (!string.IsNullOrEmpty(rightText)) h.Right(rightText!);
                switch (position) {
                    case ExcelHeaderFooterPosition.Left: h.LeftImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    case ExcelHeaderFooterPosition.Center: h.CenterImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    default: h.RightImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                }
            });
            return this;
        }

        /// <summary>Sets a header logo using an explicit remote-network policy.</summary>
        public async Task<SheetComposer> HeaderLogoFromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions,
                                           ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(url, remoteImageOptions, cancellationToken).ConfigureAwait(false);
            HeaderFooter(h => {
                if (!string.IsNullOrEmpty(leftText)) h.Left(leftText!);
                if (!string.IsNullOrEmpty(centerText)) h.Center(centerText!);
                if (!string.IsNullOrEmpty(rightText)) h.Right(rightText!);
                switch (position) {
                    case ExcelHeaderFooterPosition.Left: h.LeftImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    case ExcelHeaderFooterPosition.Center: h.CenterImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    default: h.RightImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                }
            });
            return this;
        }

        /// <summary>
        /// Asynchronously sets a footer logo from a URL and optional footer text.
        /// </summary>
        public async Task<SheetComposer> FooterLogoFromUrlAsync(string url, ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            HeaderFooter(h => {
                if (!string.IsNullOrEmpty(leftText)) h.FooterLeft(leftText!);
                if (!string.IsNullOrEmpty(centerText)) h.FooterCenter(centerText!);
                if (!string.IsNullOrEmpty(rightText)) h.FooterRight(rightText!);
                switch (position) {
                    case ExcelHeaderFooterPosition.Left: h.FooterLeftImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    case ExcelHeaderFooterPosition.Center: h.FooterCenterImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    default: h.FooterRightImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                }
            });
            return this;
        }

        /// <summary>Sets a footer logo using an explicit remote-network policy.</summary>
        public async Task<SheetComposer> FooterLogoFromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions,
                                           ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(url, remoteImageOptions, cancellationToken).ConfigureAwait(false);
            HeaderFooter(h => {
                if (!string.IsNullOrEmpty(leftText)) h.FooterLeft(leftText!);
                if (!string.IsNullOrEmpty(centerText)) h.FooterCenter(centerText!);
                if (!string.IsNullOrEmpty(rightText)) h.FooterRight(rightText!);
                switch (position) {
                    case ExcelHeaderFooterPosition.Left: h.FooterLeftImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    case ExcelHeaderFooterPosition.Center: h.FooterCenterImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                    default: h.FooterRightImage(remote.ToBytes(), remote.ContentType, widthPoints, heightPoints); break;
                }
            });
            return this;
        }
    }
}
