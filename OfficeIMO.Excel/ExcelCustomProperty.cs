using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a custom workbook property value.
    /// </summary>
    public sealed class ExcelCustomProperty {
        /// <summary>
        /// Gets or sets the raw value of the custom property.
        /// </summary>
        public object? Value { get; set; }

        /// <summary>
        /// Gets the custom property value type.
        /// </summary>
        public ExcelCustomPropertyType PropertyType { get; set; }

        /// <summary>
        /// Gets the value as a date/time when the property type is a date.
        /// </summary>
        public DateTime? Date => Value is DateTime value ? value : null;

        /// <summary>
        /// Gets the value as an integer when the property type is an integer.
        /// </summary>
        public int? NumberInteger => Value is int value ? value : null;

        /// <summary>
        /// Gets the value as a double when the property type is a floating point number.
        /// </summary>
        public double? NumberDouble => Value is double value ? value : null;

        /// <summary>
        /// Gets the value as text when the property type is textual.
        /// </summary>
        public string? Text => Value is string value ? value : null;

        /// <summary>
        /// Gets the value as a boolean when the property type is YesNo.
        /// </summary>
        public bool? Bool => Value is bool value ? value : null;

        /// <summary>
        /// Creates an empty custom property.
        /// </summary>
        public ExcelCustomProperty() {
            Value = string.Empty;
            PropertyType = ExcelCustomPropertyType.Text;
        }

        /// <summary>
        /// Creates a custom property with the specified value and type.
        /// </summary>
        public ExcelCustomProperty(object? value, ExcelCustomPropertyType propertyType) {
            Value = value;
            PropertyType = propertyType;
        }

        /// <summary>
        /// Creates a string custom property.
        /// </summary>
        public ExcelCustomProperty(string value) : this(value, ExcelCustomPropertyType.Text) { }

        /// <summary>
        /// Creates a boolean custom property.
        /// </summary>
        public ExcelCustomProperty(bool value) : this(value, ExcelCustomPropertyType.YesNo) { }

        /// <summary>
        /// Creates a date/time custom property.
        /// </summary>
        public ExcelCustomProperty(DateTime value) : this(value, ExcelCustomPropertyType.DateTime) { }

        /// <summary>
        /// Creates an integer custom property.
        /// </summary>
        public ExcelCustomProperty(int value) : this(value, ExcelCustomPropertyType.NumberInteger) { }

        /// <summary>
        /// Creates a floating point custom property.
        /// </summary>
        public ExcelCustomProperty(double value) : this(value, ExcelCustomPropertyType.NumberDouble) { }

        internal ExcelCustomProperty(CustomDocumentProperty property) {
            if (property.VTInt32 != null) {
                Value = int.Parse(property.VTInt32.Text, CultureInfo.InvariantCulture);
                PropertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTInt64 != null) {
                long value = long.Parse(property.VTInt64.Text, CultureInfo.InvariantCulture);
                Value = value >= int.MinValue && value <= int.MaxValue ? (int)value : value;
                PropertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTFileTime != null) {
                Value = DateTime.Parse(property.VTFileTime.Text, CultureInfo.InvariantCulture).ToUniversalTime();
                PropertyType = ExcelCustomPropertyType.DateTime;
            } else if (property.VTDate != null) {
                Value = DateTime.Parse(property.VTDate.Text, CultureInfo.InvariantCulture).ToUniversalTime();
                PropertyType = ExcelCustomPropertyType.DateTime;
            } else if (property.VTFloat != null) {
                Value = double.Parse(property.VTFloat.Text, CultureInfo.InvariantCulture);
                PropertyType = ExcelCustomPropertyType.NumberDouble;
            } else if (property.VTDouble != null) {
                Value = double.Parse(property.VTDouble.Text, CultureInfo.InvariantCulture);
                PropertyType = ExcelCustomPropertyType.NumberDouble;
            } else if (property.VTLPWSTR != null) {
                Value = property.VTLPWSTR.Text;
                PropertyType = ExcelCustomPropertyType.Text;
            } else if (property.VTBool != null) {
                Value = ParseBooleanProperty(property.VTBool.Text);
                PropertyType = ExcelCustomPropertyType.YesNo;
            } else {
                Value = string.Empty;
                PropertyType = ExcelCustomPropertyType.Text;
            }
        }

        internal CustomDocumentProperty ToOpenXml(string name) {
            var property = new CustomDocumentProperty {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                Name = name
            };

            switch (PropertyType) {
                case ExcelCustomPropertyType.DateTime:
                    property.VTFileTime = new VTFileTime(string.Format(CultureInfo.InvariantCulture, "{0:s}Z", Convert.ToDateTime(Value, CultureInfo.InvariantCulture).ToUniversalTime()));
                    break;
                case ExcelCustomPropertyType.NumberInteger:
                    long integer = Convert.ToInt64(Value, CultureInfo.InvariantCulture);
                    if (integer >= int.MinValue && integer <= int.MaxValue) {
                        property.VTInt32 = new VTInt32(integer.ToString(CultureInfo.InvariantCulture));
                    } else {
                        property.VTInt64 = new VTInt64(integer.ToString(CultureInfo.InvariantCulture));
                    }
                    break;
                case ExcelCustomPropertyType.NumberDouble:
                    property.VTDouble = new VTDouble(Convert.ToDouble(Value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
                    break;
                case ExcelCustomPropertyType.YesNo:
                    property.VTBool = new VTBool(Convert.ToBoolean(Value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture).ToLowerInvariant());
                    break;
                default:
                    property.VTLPWSTR = new VTLPWSTR(Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty);
                    break;
            }

            return property;
        }

        private static bool ParseBooleanProperty(string text) {
            if (bool.TryParse(text, out bool result)) {
                return result;
            }

            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numeric) && (numeric == 0 || numeric == 1)) {
                return numeric == 1;
            }

            throw new FormatException($"The custom property boolean value '{text}' is not valid.");
        }
    }
}
