using System;

namespace OfficeIMO.Drawing;

/// <summary>Thrown when a raster export exceeds a configured allocation or encoder limit.</summary>
public sealed class OfficeImageExportLimitException : InvalidOperationException {
    internal OfficeImageExportLimitException(
        double requestedScale,
        long requestedPixels,
        long maximumPixels,
        int maximumDimension)
        : base(CreateMessage(requestedScale, requestedPixels, maximumPixels, maximumDimension)) {
        RequestedScale = requestedScale;
        RequestedPixels = requestedPixels;
        MaximumPixels = maximumPixels;
        MaximumDimension = maximumDimension;
    }

    /// <summary>Caller-requested scale.</summary>
    public double RequestedScale { get; }

    /// <summary>Requested pixel count, or <see cref="long.MaxValue"/> when it exceeded numeric bounds.</summary>
    public long RequestedPixels { get; }

    /// <summary>Effective pixel-count ceiling.</summary>
    public long MaximumPixels { get; }

    /// <summary>Effective per-dimension ceiling.</summary>
    public int MaximumDimension { get; }

    private static string CreateMessage(
        double requestedScale,
        long requestedPixels,
        long maximumPixels,
        int maximumDimension) =>
        "The requested raster export at scale " +
        requestedScale.ToString("0.########", System.Globalization.CultureInfo.InvariantCulture) +
        " requires " +
        (requestedPixels == long.MaxValue
            ? "more pixels than can be represented"
            : requestedPixels.ToString(System.Globalization.CultureInfo.InvariantCulture) + " pixels") +
        ", exceeding the effective limit of " +
        maximumPixels.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        " pixels and " +
        maximumDimension.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        " pixels per dimension.";
}
