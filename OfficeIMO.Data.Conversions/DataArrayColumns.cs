namespace OfficeIMO.Data.Conversions;

/// <summary>How arrays inside a row become table columns.</summary>
public enum DataArrayColumns {
    /// <summary>One column per item, named with the item index (<c>tags[0]</c>, <c>items[1].sku</c>). Lossless.</summary>
    Indexed = 0,
    /// <summary>Scalar items joined into one cell with <see cref="DataConversionOptions.ArraySeparator"/>. Readable, but lossy and reported.</summary>
    Join
}
