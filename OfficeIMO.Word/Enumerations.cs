using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Defines custom property types available for Word documents.
    /// </summary>
    public enum PropertyTypes : int {
        /// <summary>
        /// Property type is not defined.
        /// </summary>
        Undefined,

        /// <summary>
        /// Represents a yes/no property type.
        /// </summary>
        YesNo,

        /// <summary>
        /// Represents a text property type.
        /// </summary>
        Text,

        /// <summary>
        /// Represents a date and time property type.
        /// </summary>
        DateTime,

        /// <summary>
        /// Represents an integer number property type.
        /// </summary>
        NumberInteger,

        /// <summary>
        /// Represents a double number property type.
        /// </summary>
        NumberDouble
    }
    /// <summary>
    /// Specifies the capitalization style applied to text.
    /// </summary>
    public enum CapsStyle {
        /// <summary>
        /// No caps, characters as written.
        /// </summary>
        None,

        /// <summary>
        /// All caps, make every character uppercase.
        /// </summary>
        Caps,

        /// <summary>
        /// Small caps, make all characters capital but with a smaller font size.
        /// </summary>
        SmallCaps
    };

    /// <summary>
    /// Shape types supported by <see cref="WordDocument.AddShape(ShapeType,double,double,string,string,double)"/>.
    /// </summary>
    public enum ShapeType {
        /// <summary>
        /// A rectangular shape.
        /// </summary>
        Rectangle,

        /// <summary>
        /// An elliptical shape.
        /// </summary>
        Ellipse,

        /// <summary>
        /// A straight line shape.
        /// </summary>
        Line
    }

}
