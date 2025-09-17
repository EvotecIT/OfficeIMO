using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a custom property value stored in a Word document.
    /// </summary>
    public class WordCustomProperty {
        //public string Name;
        /// <summary>
        /// Gets or sets the raw value of the custom property.
        /// </summary>
        /// <remarks>The actual type is defined by <see cref="PropertyType"/>.</remarks>
        public object Value { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the kind of custom property.
        /// </summary>
        /// <remarks>
        /// This determines how <see cref="Value"/> is interpreted when reading or
        /// writing the property to a document.
        /// </remarks>
        public PropertyTypes PropertyType;

        /// <summary>
        /// Gets the value as a <see cref="DateTime"/> when the property type is a date.
        /// </summary>
        /// <remarks>Returns <see langword="null"/> if the underlying value is not a date.</remarks>
        public DateTime? Date {
            get {
                if ((Value) is DateTime) {
                    return (DateTime)Value;
                }

                return null;
            }
        }
        /// <summary>
        /// Gets the value as an <see cref="int"/> when the property type represents an integer.
        /// </summary>
        /// <remarks>Returns <see langword="null"/> when the value is not an integer.</remarks>
        public int? NumberInteger {
            get {
                if ((Value) is int) {
                    return (int)Value;
                }

                return null;

            }
        }
        /// <summary>
        /// Gets the value as a <see cref="double"/> when the property type represents a floating point number.
        /// </summary>
        /// <remarks>Returns <see langword="null"/> when the value is not a double.</remarks>
        public double? NumberDouble {
            get {
                if ((Value) is double) {
                    return (double)Value;
                }

                return null;
            }
        }
        /// <summary>
        /// Gets the value as text when the property type is textual.
        /// </summary>
        /// <remarks>Returns <see langword="null"/> if the value does not contain text.</remarks>
        public string? Text {
            get {
                if ((Value) is string) {
                    return (string)Value;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the value as a boolean when the property type is <see cref="PropertyTypes.YesNo"/>.
        /// </summary>
        /// <remarks>Returns <see langword="null"/> when the value is not a boolean.</remarks>
        public bool? Bool {
            get {
                if ((Value) is bool) {
                    return (bool)Value;
                }

                return null;
            }
        }

        /// <summary>
        /// Creates a custom property with the specified value and type.
        /// </summary>
        /// <param name="value">Property value.</param>
        /// <param name="propertyType">Type of the property.</param>
        public WordCustomProperty(Object value, PropertyTypes propertyType) {
            this.PropertyType = propertyType;
            this.Value = value;
        }

        /// <summary>
        /// Creates a boolean custom property.
        /// </summary>
        /// <param name="value">Boolean value.</param>
        public WordCustomProperty(bool value) {
            this.PropertyType = PropertyTypes.YesNo;
            this.Value = value;
        }

        /// <summary>
        /// Creates a date/time custom property.
        /// </summary>
        /// <param name="value">Date/time value.</param>
        public WordCustomProperty(DateTime value) {
            this.PropertyType = PropertyTypes.DateTime;
            this.Value = value;
        }

        /// <summary>
        /// Creates a string custom property.
        /// </summary>
        /// <param name="value">Text value.</param>
        public WordCustomProperty(string value) {
            this.PropertyType = PropertyTypes.Text;
            this.Value = value;
        }

        /// <summary>
        /// Creates a double custom property.
        /// </summary>
        /// <param name="value">Numeric value.</param>
        public WordCustomProperty(double value) {
            this.PropertyType = PropertyTypes.NumberDouble;
            this.Value = value;
        }

        /// <summary>
        /// Creates an integer custom property.
        /// </summary>
        /// <param name="value">Integer value.</param>
        public WordCustomProperty(int value) {
            this.PropertyType = PropertyTypes.NumberInteger;
            this.Value = value;
        }

        /// <summary>
        /// Creates an empty custom property.
        /// </summary>
        public WordCustomProperty() { }

        internal WordCustomProperty(CustomDocumentProperty customDocumentProperty) {
            if (customDocumentProperty != null) {
                if (customDocumentProperty.VTInt32 != null) {
                    this.Value = int.Parse(customDocumentProperty.VTInt32.Text);
                    this.PropertyType = PropertyTypes.NumberInteger;
                } else if (customDocumentProperty.VTFileTime != null) {
                    this.Value = DateTime.Parse(customDocumentProperty.VTFileTime.Text).ToUniversalTime();
                    this.PropertyType = PropertyTypes.DateTime;
                } else if (customDocumentProperty.VTFloat != null) {
                    this.Value = double.Parse(customDocumentProperty.VTFloat.Text);
                    this.PropertyType = PropertyTypes.NumberDouble;
                } else if (customDocumentProperty.VTLPWSTR != null) {
                    this.Value = customDocumentProperty.VTLPWSTR.Text;
                    this.PropertyType = PropertyTypes.Text;
                } else if (customDocumentProperty.VTBool != null) {
                    this.Value = bool.Parse(customDocumentProperty.VTBool.Text);
                    this.PropertyType = PropertyTypes.YesNo;
                } else if (customDocumentProperty.VTDouble != null) {
                    this.Value = double.Parse(customDocumentProperty.VTDouble.Text);
                    this.PropertyType = PropertyTypes.NumberDouble;
                } else if (customDocumentProperty.VTInt64 != null) {
                    this.Value = long.Parse(customDocumentProperty.VTInt64.Text);
                    this.PropertyType = PropertyTypes.NumberInteger;
                } else if (customDocumentProperty.VTVector != null) {
                    this.Value = customDocumentProperty.VTVector;
                    this.PropertyType = PropertyTypes.Text;
                } else if (customDocumentProperty.VTEmpty != null) {
                    this.Value = "";
                    this.PropertyType = PropertyTypes.Text;
                } else if (customDocumentProperty.VTDate != null) {
                    this.Value = DateTime.Parse(customDocumentProperty.VTDate.Text).ToUniversalTime();
                    this.PropertyType = PropertyTypes.DateTime;
                } else {
                    Debug.WriteLine("Please add new type handling for customDocumentProperty. ");
                }

            } else {
                Debug.WriteLine("This shouldn't really happen. It means customDocumentProperty is not available.");
            }
        }
    }
}
