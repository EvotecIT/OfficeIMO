using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a custom workbook property value.
    /// </summary>
    public sealed class ExcelCustomProperty {
        private object? _value;
        private ExcelCustomPropertyType _propertyType;
        private Action? _changed;

        /// <summary>
        /// Gets or sets the raw value of the custom property.
        /// </summary>
        public object? Value {
            get => _value;
            set {
                if (!Equals(_value, value)) {
                    _value = value;
                    MarkChanged();
                }
            }
        }

        /// <summary>
        /// Gets the custom property value type.
        /// </summary>
        public ExcelCustomPropertyType PropertyType {
            get => _propertyType;
            set {
                if (_propertyType != value) {
                    _propertyType = value;
                    MarkChanged();
                }
            }
        }

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
        /// Gets the value as binary data when the property type is Binary.
        /// </summary>
        public byte[]? Binary => Value is byte[] value ? (byte[])value.Clone() : null;

        /// <summary>
        /// Creates an empty custom property.
        /// </summary>
        public ExcelCustomProperty() {
            _value = string.Empty;
            _propertyType = ExcelCustomPropertyType.Text;
        }

        /// <summary>
        /// Creates a custom property with the specified value and type.
        /// </summary>
        public ExcelCustomProperty(object? value, ExcelCustomPropertyType propertyType) {
            _value = value;
            _propertyType = propertyType;
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

        /// <summary>
        /// Creates a binary custom property.
        /// </summary>
        public ExcelCustomProperty(byte[] value) : this(value == null ? throw new ArgumentNullException(nameof(value)) : (byte[])value.Clone(), ExcelCustomPropertyType.Binary) { }

        internal ExcelCustomProperty(CustomDocumentProperty property) {
            if (property.VTInt32 != null) {
                _value = int.Parse(property.VTInt32.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTInt64 != null) {
                long value = long.Parse(property.VTInt64.Text, CultureInfo.InvariantCulture);
                _value = value >= int.MinValue && value <= int.MaxValue ? (int)value : value;
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTUnsignedByte != null) {
                _value = byte.Parse(property.VTUnsignedByte.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTUnsignedShort != null) {
                _value = ushort.Parse(property.VTUnsignedShort.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTUnsignedInt32 != null) {
                _value = uint.Parse(property.VTUnsignedInt32.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTUnsignedInteger != null) {
                _value = ulong.Parse(property.VTUnsignedInteger.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTUnsignedInt64 != null) {
                _value = ulong.Parse(property.VTUnsignedInt64.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberInteger;
            } else if (property.VTFileTime != null) {
                _value = DateTime.Parse(property.VTFileTime.Text, CultureInfo.InvariantCulture).ToUniversalTime();
                _propertyType = ExcelCustomPropertyType.DateTime;
            } else if (property.VTDate != null) {
                _value = DateTime.Parse(property.VTDate.Text, CultureInfo.InvariantCulture).ToUniversalTime();
                _propertyType = ExcelCustomPropertyType.DateTime;
            } else if (property.VTFloat != null) {
                _value = double.Parse(property.VTFloat.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberDouble;
            } else if (property.VTDouble != null) {
                _value = double.Parse(property.VTDouble.Text, CultureInfo.InvariantCulture);
                _propertyType = ExcelCustomPropertyType.NumberDouble;
            } else if (property.VTLPWSTR != null) {
                _value = property.VTLPWSTR.Text;
                _propertyType = ExcelCustomPropertyType.Text;
            } else if (property.VTBool != null) {
                _value = ParseBooleanProperty(property.VTBool.Text);
                _propertyType = ExcelCustomPropertyType.YesNo;
            } else if (property.VTBlob != null) {
                _value = ParseBinaryProperty(property.VTBlob.Text);
                _propertyType = ExcelCustomPropertyType.Binary;
            } else if (property.VTOBlob != null) {
                _value = ParseBinaryProperty(property.VTOBlob.Text);
                _propertyType = ExcelCustomPropertyType.Binary;
            } else {
                _value = string.Empty;
                _propertyType = ExcelCustomPropertyType.Text;
            }
        }

        internal void SetChangeHandler(Action? changed) {
            _changed = changed;
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
                    if (Value is byte unsignedByte) {
                        property.VTUnsignedByte = new VTUnsignedByte(unsignedByte.ToString(CultureInfo.InvariantCulture));
                    } else if (Value is ushort unsignedShort) {
                        property.VTUnsignedShort = new VTUnsignedShort(unsignedShort.ToString(CultureInfo.InvariantCulture));
                    } else if (Value is ulong unsignedLong) {
                        property.VTUnsignedInt64 = new VTUnsignedInt64(unsignedLong.ToString(CultureInfo.InvariantCulture));
                    } else if (Value is uint unsignedInteger) {
                        property.VTUnsignedInt32 = new VTUnsignedInt32(unsignedInteger.ToString(CultureInfo.InvariantCulture));
                    } else {
                        long integer = Convert.ToInt64(Value, CultureInfo.InvariantCulture);
                        if (integer >= int.MinValue && integer <= int.MaxValue) {
                            property.VTInt32 = new VTInt32(integer.ToString(CultureInfo.InvariantCulture));
                        } else {
                            property.VTInt64 = new VTInt64(integer.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                    break;
                case ExcelCustomPropertyType.NumberDouble:
                    property.VTDouble = new VTDouble(Convert.ToDouble(Value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
                    break;
                case ExcelCustomPropertyType.YesNo:
                    property.VTBool = new VTBool(Convert.ToBoolean(Value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture).ToLowerInvariant());
                    break;
                case ExcelCustomPropertyType.Binary:
                    property.VTBlob = new VTBlob(Convert.ToBase64String(GetBinaryValue()));
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

        private static byte[] ParseBinaryProperty(string? text) {
            return string.IsNullOrEmpty(text)
                ? Array.Empty<byte>()
                : Convert.FromBase64String(text);
        }

        private byte[] GetBinaryValue() {
            return Value is byte[] bytes
                ? bytes
                : Convert.FromBase64String(Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        private void MarkChanged() {
            _changed?.Invoke();
        }
    }
}
