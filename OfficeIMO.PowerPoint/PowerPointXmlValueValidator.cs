using System;
using System.Xml;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointXmlValueValidator {
        internal static void ValidateCharacters(string value,
            string parameterName, string valueDescription) {
            try {
                XmlConvert.VerifyXmlChars(value);
            } catch (XmlException exception) {
                throw new ArgumentException(
                    $"{valueDescription} contains characters that are not valid in Open XML.",
                    parameterName, exception);
            }
        }
    }
}
