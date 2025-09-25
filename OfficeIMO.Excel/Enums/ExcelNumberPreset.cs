namespace OfficeIMO.Excel {
    /// <summary>
    /// Common number format presets for convenience. For full control, use a custom format string.
    /// </summary>
    public enum ExcelNumberPreset {
        /// <summary>General format with no specific pattern.</summary>
        General,
        /// <summary>Whole number with thousands separators.</summary>
        Integer,
        /// <summary>Decimal number with configurable fraction digits.</summary>
        Decimal,
        /// <summary>Percentage format.</summary>
        Percent,
        /// <summary>Currency format using culture-specific symbol.</summary>
        Currency,
        /// <summary>Scientific notation.</summary>
        Scientific,
        /// <summary>Short date format (yyyy-mm-dd).</summary>
        DateShort,
        /// <summary>Long date/time format (yyyy-mm-dd hh:mm).</summary>
        DateLong,
        /// <summary>Time format (h:mm:ss).</summary>
        Time,
        /// <summary>Date and time format (yyyy-mm-dd hh:mm:ss).</summary>
        DateTime,
        /// <summary>Duration in hours format ([h]:mm:ss).</summary>
        DurationHours,
        /// <summary>Text format.</summary>
        Text
    }
}

