using System;
using System.Globalization;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    /// <summary>A worksheet dimension expressed in Excel character units or as a physical length.</summary>
    /// <remarks>
    /// Character column widths retain the stored Excel value, including cell padding.
    /// Physical column widths resolve through the workbook default font and round to the nearest
    /// 96-DPI pixel. Row heights use the physical length directly; character heights use the
    /// default font's maximum digit advance, without column padding. A zero value clears a custom dimension.
    /// </remarks>
    public readonly struct ExcelColumnWidth : IEquatable<ExcelColumnWidth> {
        private readonly double _characters;
        private readonly OfficeLength _length;
        private readonly bool _isLength;

        private ExcelColumnWidth(double characters) { _characters = characters; _length = default; _isLength = false; }
        private ExcelColumnWidth(OfficeLength length) { _characters = 0D; _length = length; _isLength = true; }

        /// <summary>Creates a dimension in Excel character units. Non-finite values are rejected.</summary>
        public static ExcelColumnWidth Characters(double value) {
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value));
            return new ExcelColumnWidth(value);
        }

        /// <summary>Creates a dimension from an explicit physical length.</summary>
        public static ExcelColumnWidth Length(OfficeLength value) => new ExcelColumnWidth(value);

        /// <summary>Gets whether the value was specified as a physical length.</summary>
        public bool IsPhysical => _isLength;

        /// <summary>Gets the stored column width for a worksheet, resolving physical values through its default font.</summary>
        /// <remarks>The physical width rounds to whole pixels and then to the stored 1/256-character grid.</remarks>
        public double ToCharacters(ExcelSheet sheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!_isLength) return _characters;
            double pixels = Math.Round(_length.ToPixels(), MidpointRounding.AwayFromZero);
            double digit = ExcelColumnGeometry.GetMaximumDigitWidth(sheet.WorksheetPart);
            return Math.Truncate(pixels / digit * 256D) / 256D;
        }

        /// <summary>Gets the physical column width for a worksheet, including native column padding and pixel rounding.</summary>
        public OfficeLength ToLength(ExcelSheet sheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (_isLength) return _length;
            double digit = ExcelColumnGeometry.GetMaximumDigitWidth(sheet.WorksheetPart);
            return OfficeLength.Pixels(_characters <= 0D ? _characters * digit
                : Math.Truncate((256D * _characters + Math.Truncate(128D / digit)) / 256D * digit));
        }

        /// <summary>Resolves a row height without applying column padding or column-width rounding.</summary>
        internal double ToRowPoints(ExcelSheet sheet) => _isLength ? _length.ToPoints()
            : _characters * ExcelColumnGeometry.GetMaximumDigitWidth(sheet.WorksheetPart) * 72D / 96D;

        /// <inheritdoc />
        public bool Equals(ExcelColumnWidth other) => _isLength == other._isLength
            && (_isLength ? _length == other._length : _characters.Equals(other._characters));
        /// <inheritdoc />
        public override bool Equals(object? obj) => obj is ExcelColumnWidth other && Equals(other);
        /// <inheritdoc />
        public override int GetHashCode() => _isLength ? _length.GetHashCode() ^ int.MinValue : _characters.GetHashCode();
        /// <summary>Formats the specified dimension using invariant units.</summary>
        public override string ToString() => _isLength ? _length.ToString() : _characters.ToString("G", CultureInfo.InvariantCulture) + "ch";
        /// <summary>Compares dimensions in the same authored unit. Conversion requires a worksheet.</summary>
        public static bool operator ==(ExcelColumnWidth left, ExcelColumnWidth right) => left.Equals(right);
        /// <summary>Compares dimensions in the same authored unit. Conversion requires a worksheet.</summary>
        public static bool operator !=(ExcelColumnWidth left, ExcelColumnWidth right) => !left.Equals(right);
    }
}
