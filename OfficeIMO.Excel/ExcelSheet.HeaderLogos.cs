using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Asynchronously sets a header logo from a URL in the given position. Optional page text can be supplied.
        /// </summary>
        public async Task HeaderLogoFromUrlAsync(string url, ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                  double? widthPoints = null, double? heightPoints = null,
                                  string? leftText = null, string? centerText = null, string? rightText = null,
                                  CancellationToken cancellationToken = default) {
            SetHeaderFooter(leftText, centerText, rightText);
            await SetHeaderImageFromUrlAsync(position, url, widthPoints, heightPoints, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Sets a header logo using an explicit remote-network policy.</summary>
        public async Task HeaderLogoFromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions,
                                  ExcelHeaderFooterPosition position = ExcelHeaderFooterPosition.Right,
                                  double? widthPoints = null, double? heightPoints = null,
                                  string? leftText = null, string? centerText = null, string? rightText = null,
                                  CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            SetHeaderFooter(leftText, centerText, rightText);
            await SetHeaderImageFromUrlAsync(position, url, remoteImageOptions, widthPoints, heightPoints,
                cancellationToken).ConfigureAwait(false);
        }
    }
}
