namespace OfficeIMO.Word;

public partial class WordImage {
    /// <summary>Gets or sets the Top source crop as a fraction of the original image; one means its full dimension.</summary>
    /// <remarks>Negative fractions extend the source window. This changes the source window while retaining the displayed frame. Use <see cref="CropTop"/> for a physical crop that also adjusts the frame and anchor.</remarks>
    public double? CropTopRatio {
        get => CropTopPercentage / CropPercentageUnits;
        set => CropTopPercentage = ToCropPercentage(value);
    }

    /// <summary>Gets or sets the Bottom source crop as a fraction of the original image; one means its full dimension.</summary>
    /// <remarks>Negative fractions extend the source window. This changes the source window while retaining the displayed frame. Use <see cref="CropBottom"/> for a physical crop that also adjusts the frame and anchor.</remarks>
    public double? CropBottomRatio {
        get => CropBottomPercentage / CropPercentageUnits;
        set => CropBottomPercentage = ToCropPercentage(value);
    }

    /// <summary>Gets or sets the Left source crop as a fraction of the original image; one means its full dimension.</summary>
    /// <remarks>Negative fractions extend the source window. This changes the source window while retaining the displayed frame. Use <see cref="CropLeft"/> for a physical crop that also adjusts the frame and anchor.</remarks>
    public double? CropLeftRatio {
        get => CropLeftPercentage / CropPercentageUnits;
        set => CropLeftPercentage = ToCropPercentage(value);
    }

    /// <summary>Gets or sets the Right source crop as a fraction of the original image; one means its full dimension.</summary>
    /// <remarks>Negative fractions extend the source window. This changes the source window while retaining the displayed frame. Use <see cref="CropRight"/> for a physical crop that also adjusts the frame and anchor.</remarks>
    public double? CropRightRatio {
        get => CropRightPercentage / CropPercentageUnits;
        set => CropRightPercentage = ToCropPercentage(value);
    }

    private static int? ToCropPercentage(double? fraction) {
        if (!fraction.HasValue) return null;
        double percentage = Math.Round(fraction.Value * CropPercentageUnits, MidpointRounding.AwayFromZero);
        if (double.IsNaN(percentage) || double.IsInfinity(percentage) || percentage < int.MinValue || percentage > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(fraction), "The source crop must fit the DrawingML percentage range.");
        return checked((int)percentage);
    }
}
