using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordCustomProperty {
        //public string Name;
        public Object Value;
        public PropertyTypes PropertyType;

        public DateTime? Date {
            get {
                if ((Value) is DateTime) {
                    return (DateTime) Value;
                }

                return null;
            }
        }
        public int? NumberInteger {
            get {
                if ((Value) is int) {
                    return (int) Value;
                }

                return null;
                
            }
        }
        public double? NumberDouble {
            get {
                if ((Value) is double) {
                    return (double) Value;
                }

                return null;
            }
        }
        public string Text {
            get {
                if ((Value) is string) {
                    return (string)Value;
                }

                return null;
            }
        }

        public WordCustomProperty(Object value, PropertyTypes propertyType) {
            this.PropertyType = propertyType;
            this.Value = value;
        }
        public WordCustomProperty(bool value) {
            this.PropertyType = PropertyTypes.YesNo;
            this.Value = value;
        }
        public WordCustomProperty(DateTime value) {
            this.PropertyType = PropertyTypes.DateTime;
            this.Value = value;
        }
        public WordCustomProperty(string value) {
            this.PropertyType = PropertyTypes.Text;
            this.Value = value;
        }
        public WordCustomProperty(double value) {
            this.PropertyType = PropertyTypes.NumberDouble;
            this.Value = value;
        }
        public WordCustomProperty(int value) {
            this.PropertyType = PropertyTypes.NumberInteger;
            this.Value = value;
        }
        public WordCustomProperty() {}

        public WordCustomProperty(CustomDocumentProperty customDocumentProperty) {
            if (customDocumentProperty != null) {
                if (customDocumentProperty.VTInt32 != null) {
                    this.Value = int.Parse(customDocumentProperty.VTInt32.Text);
                    this.PropertyType = PropertyTypes.NumberInteger;
                } else if (customDocumentProperty.VTFileTime != null) {
                    this.Value = DateTime.Parse(customDocumentProperty.VTFileTime.Text);
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
                } else {
                    throw new InvalidOperationException("Weird?");
                }
            }
        }

    }
}