using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word;

public partial class WordImage {
    /// <summary>Gets or sets the horizontal reference used by a floating image.</summary>
    public WordHorizontalRelativePosition HorizontalPositionRelativeFrom {
        get => (horizontalPosition.RelativeFrom?.Value ?? HorizontalRelativePositionValues.Page).ToOfficeEnum();
        set => horizontalPosition.RelativeFrom = value.ToOpenXml();
    }

    /// <summary>Gets or sets the horizontal image offset.</summary>
    public OfficeLength? HorizontalPositionOffset {
        get => ReadPositionOffset(horizontalPosition.PositionOffset);
        set => horizontalPosition.PositionOffset = value.HasValue
            ? new PositionOffset { Text = value.Value.ToEmus().ToString(CultureInfo.InvariantCulture) }
            : null;
    }

    /// <summary>Gets or sets the vertical reference used by a floating image.</summary>
    public WordVerticalRelativePosition VerticalPositionRelativeFrom {
        get => (verticalPosition.RelativeFrom?.Value ?? VerticalRelativePositionValues.Page).ToOfficeEnum();
        set => verticalPosition.RelativeFrom = value.ToOpenXml();
    }

    /// <summary>Gets or sets the vertical image offset.</summary>
    public OfficeLength? VerticalPositionOffset {
        get => ReadPositionOffset(verticalPosition.PositionOffset);
        set => verticalPosition.PositionOffset = value.HasValue
            ? new PositionOffset { Text = value.Value.ToEmus().ToString(CultureInfo.InvariantCulture) }
            : null;
    }

    private static OfficeLength? ReadPositionOffset(PositionOffset? offset) =>
        long.TryParse(offset?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) ? OfficeLength.Emus(value) : null;
}
