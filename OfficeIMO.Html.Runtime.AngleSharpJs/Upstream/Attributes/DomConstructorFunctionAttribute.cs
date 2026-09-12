namespace AngleSharp.Js.Attributes
{
    using System;

    /// <summary>
    /// This attribute is used to mark a method to be uses as a
    /// constructor function from scripts.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    public sealed class DomConstructorFunctionAttribute : Attribute
    {
        /// <summary>
        /// Creates a new DomConstructorFunctionAttribute.
        /// </summary>
        /// <param name="officialName">
        /// The official name of the decorated method.
        /// </param>
        public DomConstructorFunctionAttribute(String officialName)
        {
            OfficialName = officialName;
        }

        /// <summary>
        /// Gets the official name of the given class.
        /// </summary>
        public String OfficialName { get; }
    }
}
