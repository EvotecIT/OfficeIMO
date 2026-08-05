using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Finds attributes that cannot be represented by a native XLSB writer.</summary>
    internal static class XlsbOpenXmlAttributeValidator {
        private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";

        internal static OpenXmlAttribute? FindUnsupported(
            OpenXmlElement element,
            IEnumerable<string> allowedUnqualifiedNames,
            Func<OpenXmlAttribute, bool>? allowQualified = null) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            if (allowedUnqualifiedNames == null) throw new ArgumentNullException(nameof(allowedUnqualifiedNames));

            var allowed = new HashSet<string>(allowedUnqualifiedNames, StringComparer.Ordinal);
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.NamespaceUri, XmlnsNamespace, StringComparison.Ordinal)) continue;
                if (string.Equals(attribute.NamespaceUri, string.Empty, StringComparison.Ordinal)
                    && allowed.Contains(attribute.LocalName)) {
                    continue;
                }
                if (allowQualified?.Invoke(attribute) == true) continue;
                return attribute;
            }
            return null;
        }
    }
}
