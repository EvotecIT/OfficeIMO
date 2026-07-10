namespace OfficeIMO.OpenDocument;

internal static class OdfTextCodec {
    internal static string Read(XElement element) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        var builder = new StringBuilder();
        AppendValue(element.Nodes(), builder);
        return builder.ToString();
    }

    internal static void Replace(XElement element, string? text) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        element.RemoveNodes();
        Append(element, text);
    }

    internal static void Append(XElement element, string? text) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (string.IsNullOrEmpty(text)) return;

        var plain = new StringBuilder();
        int spaces = 0;
        Action flushPlain = () => {
            if (plain.Length == 0) return;
            element.Add(new XText(plain.ToString()));
            plain.Clear();
        };
        Action flushSpaces = () => {
            if (spaces == 0) return;
            flushPlain();
            var space = new XElement(OdfNamespaces.Text + "s");
            if (spaces != 1) space.SetAttributeValue(OdfNamespaces.Text + "c", spaces);
            element.Add(space);
            spaces = 0;
        };

        foreach (char character in text!) {
            if (character == ' ') {
                spaces++;
                continue;
            }
            flushSpaces();
            if (character == '\t') {
                flushPlain();
                element.Add(new XElement(OdfNamespaces.Text + "tab"));
            } else if (character == '\n') {
                flushPlain();
                element.Add(new XElement(OdfNamespaces.Text + "line-break"));
            } else if (character != '\r') {
                plain.Append(character);
            }
        }
        flushSpaces();
        flushPlain();
    }

    private static void AppendValue(IEnumerable<XNode> nodes, StringBuilder builder) {
        foreach (XNode node in nodes) {
            if (node is XText text) {
                builder.Append(text.Value);
                continue;
            }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s") {
                int count = ParsePositiveCount((string?)element.Attribute(OdfNamespaces.Text + "c"));
                builder.Append(' ', count);
            } else if (element.Name == OdfNamespaces.Text + "tab") {
                builder.Append('\t');
            } else if (element.Name == OdfNamespaces.Text + "line-break") {
                builder.Append('\n');
            } else {
                AppendValue(element.Nodes(), builder);
            }
        }
    }

    private static int ParsePositiveCount(string? value) {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) && count > 0 ? count : 1;
    }
}
