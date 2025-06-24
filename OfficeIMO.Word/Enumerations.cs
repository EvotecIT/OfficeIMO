﻿using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Defines the data type of a custom document property.
    /// </summary>
    public enum PropertyTypes : int {
        Undefined,
        YesNo,
        Text,
        DateTime,
        NumberInteger,
        NumberDouble
    }
    /// <summary>
    /// Specifies capitalization styles available for text.
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
        Rectangle,
        Ellipse,
        Line
    }

}
