using System;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Tracks an SVG composition's UTF-8 size and rejects fragments before they are
/// appended to the aggregate document.
/// </summary>
internal sealed class OfficeSvgUtf8CompositionBudget {
    private readonly long _maximumBytes;
    private long _encodedBytes;

    internal OfficeSvgUtf8CompositionBudget(long maximumBytes) {
        if (maximumBytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
    }

    internal long RemainingBytes => _maximumBytes - _encodedBytes;

    internal void Append(StringBuilder destination, string fragment) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (fragment == null) throw new ArgumentNullException(nameof(fragment));

        long fragmentBytes = Encoding.UTF8.GetByteCount(fragment);
        long actual = fragmentBytes > long.MaxValue - _encodedBytes
            ? long.MaxValue
            : _encodedBytes + fragmentBytes;
        if (actual > _maximumBytes) {
            throw new OfficeImageExportBatchLimitException(
                nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes),
                actual,
                _maximumBytes);
        }

        destination.Append(fragment);
        _encodedBytes = actual;
    }
}
