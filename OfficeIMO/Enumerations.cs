using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public enum PropertyTypes : int {
        Undefined,
        YesNo,
        Text,
        DateTime,
        NumberInteger,
        NumberDouble
    }
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
}