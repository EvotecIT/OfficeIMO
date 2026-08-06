using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word;

public partial class WordImage {
    /// <summary>Gets or sets the horizontal reference used by a floating image.</summary>
    public WordHorizontalRelativePosition HorizontalPositionRelativeFrom {
        get => (horizontalPosition.RelativeFrom?.Value ?? HorizontalRelativePositionValues.Page).ToOfficeEnum();
        set => horizontalPosition.RelativeFrom = value.ToOpenXml();
    }

    /// <summary>Gets or sets the horizontal image offset in English Metric Units.</summary>
    public long? HorizontalPositionOffset {
        get => ReadPositionOffset(horizontalPosition.PositionOffset);
        set => horizontalPosition.PositionOffset = value.HasValue
            ? new PositionOffset { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
            : null;
    }

    /// <summary>Gets or sets the vertical reference used by a floating image.</summary>
    public WordVerticalRelativePosition VerticalPositionRelativeFrom {
        get => (verticalPosition.RelativeFrom?.Value ?? VerticalRelativePositionValues.Page).ToOfficeEnum();
        set => verticalPosition.RelativeFrom = value.ToOpenXml();
    }

    /// <summary>Gets or sets the vertical image offset in English Metric Units.</summary>
    public long? VerticalPositionOffset {
        get => ReadPositionOffset(verticalPosition.PositionOffset);
        set => verticalPosition.PositionOffset = value.HasValue
            ? new PositionOffset { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
            : null;
    }

    private static long? ReadPositionOffset(PositionOffset? offset) =>
        long.TryParse(offset?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) ? value : null;
}
