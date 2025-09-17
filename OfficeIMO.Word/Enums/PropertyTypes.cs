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
}

