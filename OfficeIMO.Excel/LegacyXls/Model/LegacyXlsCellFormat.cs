namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a parsed legacy XLS XF cell format.
    /// </summary>
    public sealed class LegacyXlsCellFormat {
        /// <summary>
        /// Creates a parsed legacy XLS cell format.
        /// </summary>
        /// <param name="styleIndex">Zero-based XF index.</param>
        /// <param name="fontIndex">Legacy FontIndex referenced by the XF record.</param>
        /// <param name="numberFormatId">Legacy IFmt number format identifier.</param>
        /// <param name="isStyle">Whether this XF represents a cell style XF rather than a cell XF.</param>
        /// <param name="parentStyleIndex">Legacy parent style XF index for cell XF inheritance.</param>
        /// <param name="applyNumberFormat">Whether the number format facet is owned by this cell XF.</param>
        /// <param name="applyFont">Whether the font facet is owned by this cell XF.</param>
        /// <param name="applyFill">Whether the fill facet is owned by this cell XF.</param>
        /// <param name="fillPattern">Legacy fill pattern code.</param>
        /// <param name="fillForegroundColorIndex">Legacy foreground fill color index.</param>
        /// <param name="fillBackgroundColorIndex">Legacy background fill color index.</param>
        /// <param name="applyAlignment">Whether alignment fields should be projected.</param>
        /// <param name="horizontalAlignment">Legacy horizontal alignment code.</param>
        /// <param name="verticalAlignment">Legacy vertical alignment code.</param>
        /// <param name="wrapText">Whether text wrapping is enabled.</param>
        /// <param name="textRotation">Legacy text rotation code.</param>
        /// <param name="indent">Legacy indentation level.</param>
        /// <param name="shrinkToFit">Whether shrink-to-fit is enabled.</param>
        /// <param name="readingOrder">Legacy reading order code.</param>
        /// <param name="applyProtection">Whether protection fields should be projected.</param>
        /// <param name="locked">Whether locked protection is enabled.</param>
        /// <param name="formulaHidden">Whether formula display is hidden when the worksheet is protected.</param>
        /// <param name="quotePrefix">Whether the cell has a Lotus 1-2-3 quote-prefix marker.</param>
        /// <param name="border">Resolved border formatting, when present.</param>
        /// <param name="numberFormatCode">Resolved Excel number format code.</param>
        /// <param name="isBuiltInNumberFormat">Whether the number format id is built in.</param>
        /// <param name="isDateLike">Whether the number format should be treated as a date/time format.</param>
        public LegacyXlsCellFormat(
            ushort styleIndex,
            ushort fontIndex,
            ushort numberFormatId,
            bool isStyle,
            ushort parentStyleIndex,
            bool applyNumberFormat,
            bool applyFont,
            bool applyFill,
            byte fillPattern,
            ushort fillForegroundColorIndex,
            ushort fillBackgroundColorIndex,
            bool applyAlignment,
            byte horizontalAlignment,
            byte verticalAlignment,
            bool wrapText,
            byte textRotation,
            byte indent,
            bool shrinkToFit,
            byte readingOrder,
            bool applyProtection,
            bool locked,
            bool formulaHidden,
            bool quotePrefix,
            LegacyXlsBorder? border,
            string? numberFormatCode,
            bool isBuiltInNumberFormat,
            bool isDateLike) {
            StyleIndex = styleIndex;
            FontIndex = fontIndex;
            NumberFormatId = numberFormatId;
            IsStyle = isStyle;
            ParentStyleIndex = parentStyleIndex;
            ApplyNumberFormat = applyNumberFormat;
            ApplyFont = applyFont;
            ApplyFill = applyFill;
            FillPattern = fillPattern;
            FillForegroundColorIndex = fillForegroundColorIndex;
            FillBackgroundColorIndex = fillBackgroundColorIndex;
            ApplyAlignment = applyAlignment;
            HorizontalAlignment = horizontalAlignment;
            VerticalAlignment = verticalAlignment;
            WrapText = wrapText;
            TextRotation = textRotation;
            Indent = indent;
            ShrinkToFit = shrinkToFit;
            ReadingOrder = readingOrder;
            ApplyProtection = applyProtection;
            Locked = locked;
            FormulaHidden = formulaHidden;
            QuotePrefix = quotePrefix;
            Border = border;
            NumberFormatCode = numberFormatCode;
            IsBuiltInNumberFormat = isBuiltInNumberFormat;
            IsDateLike = isDateLike;
        }

        /// <summary>
        /// Gets the zero-based legacy XF index referenced by cells.
        /// </summary>
        public ushort StyleIndex { get; }

        /// <summary>
        /// Gets the legacy FontIndex referenced by the XF record.
        /// </summary>
        public ushort FontIndex { get; }

        /// <summary>
        /// Gets the legacy IFmt number format identifier.
        /// </summary>
        public ushort NumberFormatId { get; }

        /// <summary>
        /// Gets whether this XF represents a cell style XF rather than a cell XF.
        /// </summary>
        public bool IsStyle { get; }

        /// <summary>
        /// Gets the parent style XF index referenced by a cell XF.
        /// </summary>
        public ushort ParentStyleIndex { get; }

        /// <summary>
        /// Gets whether the number format facet is owned by this cell XF instead of inherited from the parent style XF.
        /// </summary>
        public bool ApplyNumberFormat { get; }

        /// <summary>
        /// Gets whether the font facet is owned by this cell XF instead of inherited from the parent style XF.
        /// </summary>
        public bool ApplyFont { get; }

        /// <summary>
        /// Gets whether the fill facet is owned by this cell XF instead of inherited from the parent style XF.
        /// </summary>
        public bool ApplyFill { get; }

        /// <summary>
        /// Gets the legacy fill pattern code.
        /// </summary>
        public byte FillPattern { get; }

        /// <summary>
        /// Gets the legacy foreground fill color index.
        /// </summary>
        public ushort FillForegroundColorIndex { get; }

        /// <summary>
        /// Gets the legacy background fill color index.
        /// </summary>
        public ushort FillBackgroundColorIndex { get; }

        /// <summary>
        /// Gets whether alignment fields should be projected.
        /// </summary>
        public bool ApplyAlignment { get; }

        /// <summary>
        /// Gets the legacy horizontal alignment code.
        /// </summary>
        public byte HorizontalAlignment { get; }

        /// <summary>
        /// Gets the legacy vertical alignment code.
        /// </summary>
        public byte VerticalAlignment { get; }

        /// <summary>
        /// Gets whether text wrapping is enabled.
        /// </summary>
        public bool WrapText { get; }

        /// <summary>
        /// Gets the legacy text rotation code.
        /// </summary>
        public byte TextRotation { get; }

        /// <summary>
        /// Gets the legacy indentation level.
        /// </summary>
        public byte Indent { get; }

        /// <summary>
        /// Gets whether shrink-to-fit is enabled.
        /// </summary>
        public bool ShrinkToFit { get; }

        /// <summary>
        /// Gets the legacy reading order code.
        /// </summary>
        public byte ReadingOrder { get; }

        /// <summary>
        /// Gets whether protection fields should be projected.
        /// </summary>
        public bool ApplyProtection { get; }

        /// <summary>
        /// Gets whether locked protection is enabled.
        /// </summary>
        public bool Locked { get; }

        /// <summary>
        /// Gets whether formula display is hidden when the worksheet is protected.
        /// </summary>
        public bool FormulaHidden { get; }

        /// <summary>
        /// Gets whether the cell has a Lotus 1-2-3 quote-prefix marker.
        /// </summary>
        public bool QuotePrefix { get; }

        /// <summary>
        /// Gets parsed border formatting, when present.
        /// </summary>
        public LegacyXlsBorder? Border { get; }

        /// <summary>
        /// Gets the resolved Excel number format code, when known.
        /// </summary>
        public string? NumberFormatCode { get; }

        /// <summary>
        /// Gets whether the number format identifier is built in.
        /// </summary>
        public bool IsBuiltInNumberFormat { get; }

        /// <summary>
        /// Gets whether this cell format should be treated as date/time-like.
        /// </summary>
        public bool IsDateLike { get; }
    }
}
