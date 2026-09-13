namespace AngleSharp.Dom
{
    using AngleSharp.Attributes;
    using AngleSharp.Text;
    using System;

    /// <summary>
    /// Element extensions exposed to DOM.
    /// </summary>
    [DomExposed("Element")]
    public static class DomElementExtensions
    {
        /// <summary>
        /// Returns the named attribute on the specified element.
        /// </summary>
        /// <param name="element">
        /// Current element
        /// </param>
        /// <param name="name">
        /// The name of the attribute you want to get.
        /// </param>
        /// <returns>
        /// Named attribute or null if it doesn't exist.
        /// </returns>
        [DomName("getAttributeNode")]
        public static IAttr? GetAttributeNode(this IElement element, String name)
        {
            if (element.GivenNamespaceUri.Is(NamespaceNames.HtmlUri))
            {
                name = name.HtmlLower();
            }

            return element.Attributes.GetNamedItem(name);
        }

        /// <summary>
        /// Returns the named attribute on the specified element.
        /// </summary>
        /// <param name="element">
        /// Current element
        /// </param>
        /// <param name="namespaceUri">
        /// A string specifying the namespace of the attribute.
        /// </param>
        /// <param name="localName">
        /// The name of the attribute you want to get.
        /// </param>
        /// <returns>
        /// Named attribute or null if it doesn't exist.
        /// </returns>
        [DomName("getAttributeNodeNS")]
        public static IAttr? GetAttributeNode(this IElement element, String? namespaceUri, String localName)
        {
            if (String.IsNullOrEmpty(namespaceUri))
            {
                namespaceUri = null;
            }

            return element.Attributes.GetNamedItem(namespaceUri, localName);
        }
    }
}
