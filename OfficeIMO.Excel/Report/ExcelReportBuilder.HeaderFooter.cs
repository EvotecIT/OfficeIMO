using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelReportBuilder {
        /// <summary>
        /// Asynchronously sets a header logo from a URL and optional header text.
        /// </summary>
        public async Task<ExcelReportBuilder> HeaderLogoFromUrlAsync(string url, ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            Sheet.SetHeaderFooter(headerLeft: leftText, headerCenter: centerText, headerRight: rightText);
            Sheet.SetHeaderImage(position, remote.ToBytes(), remote.ContentType, widthPoints, heightPoints);
            return this;
        }

        /// <summary>Sets a header logo using an explicit remote-network policy.</summary>
        public async Task<ExcelReportBuilder> HeaderLogoFromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions,
                                           ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(url, remoteImageOptions, cancellationToken).ConfigureAwait(false);
            Sheet.SetHeaderFooter(headerLeft: leftText, headerCenter: centerText, headerRight: rightText);
            Sheet.SetHeaderImage(position, remote.ToBytes(), remote.ContentType, widthPoints, heightPoints);
            return this;
        }

        /// <summary>
        /// Asynchronously sets a footer logo from a URL and optional footer text.
        /// </summary>
        public async Task<ExcelReportBuilder> FooterLogoFromUrlAsync(string url, ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            Sheet.SetHeaderFooter(footerLeft: leftText, footerCenter: centerText, footerRight: rightText);
            Sheet.SetFooterImage(position, remote.ToBytes(), remote.ContentType, widthPoints, heightPoints);
            return this;
        }

        /// <summary>Sets a footer logo using an explicit remote-network policy.</summary>
        public async Task<ExcelReportBuilder> FooterLogoFromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions,
                                           ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                           double? widthPoints = null, double? heightPoints = null,
                                           string? leftText = null, string? centerText = null, string? rightText = null,
                                           CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(url, remoteImageOptions, cancellationToken).ConfigureAwait(false);
            Sheet.SetHeaderFooter(footerLeft: leftText, footerCenter: centerText, footerRight: rightText);
            Sheet.SetFooterImage(position, remote.ToBytes(), remote.ContentType, widthPoints, heightPoints);
            return this;
        }
    }
}
