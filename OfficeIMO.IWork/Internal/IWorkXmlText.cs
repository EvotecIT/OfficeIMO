namespace OfficeIMO.IWork.Internal;

internal static class IWorkXmlText {
    internal static bool IsRepresentable(string value, bool allowIWorkBreaks = false) {
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (allowIWorkBreaks && current is '\u0004' or '\u0005' or '\u000c') continue;
            if (System.Xml.XmlConvert.IsXmlChar(current)) continue;
            if (char.IsHighSurrogate(current) && index + 1 < value.Length
                && char.IsLowSurrogate(value[index + 1])) {
                index++;
                continue;
            }
            return false;
        }
        return true;
    }
}
