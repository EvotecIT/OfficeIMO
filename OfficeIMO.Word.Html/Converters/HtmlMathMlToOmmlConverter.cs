using AngleSharp.Dom;
using System.Xml.Linq;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Projects the supported structural MathML subset used by HTML import into editable OMML.
    /// Unknown containers retain their supported descendants instead of flattening the whole formula.
    /// </summary>
    internal static class HtmlMathMlToOmmlConverter {
        private const string OmmlNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        private static readonly XNamespace M = OmmlNamespace;

        internal static bool TryConvert(IElement math, out string omml) {
            if (math == null) throw new ArgumentNullException(nameof(math));

            var content = ConvertChildren(math).ToList();
            if (content.Count == 0) {
                omml = string.Empty;
                return false;
            }

            var root = new XElement(
                M + "oMath",
                new XAttribute(XNamespace.Xmlns + "m", OmmlNamespace),
                content);
            omml = root.ToString(SaveOptions.DisableFormatting);
            return true;
        }

        private static IEnumerable<XElement> ConvertElement(IElement element) {
            string name = element.LocalName.ToLowerInvariant();
            switch (name) {
                case "math":
                case "mrow":
                case "mstyle":
                case "mpadded":
                case "mphantom":
                case "maction":
                    return ConvertChildren(element);
                case "semantics":
                    IElement? semanticContent = element.Children.FirstOrDefault(child =>
                        !string.Equals(child.LocalName, "annotation", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(child.LocalName, "annotation-xml", StringComparison.OrdinalIgnoreCase));
                    return semanticContent == null ? Enumerable.Empty<XElement>() : ConvertElement(semanticContent);
                case "mi":
                case "mn":
                case "mo":
                case "mtext":
                case "ms":
                    return CreateTextRun(element.TextContent);
                case "mspace":
                    return CreateTextRun(" ");
                case "mfrac":
                    return Single(CreateFraction(element));
                case "msup":
                    return Single(CreateScript(element, "sSup", "e", "sup"));
                case "msub":
                    return Single(CreateScript(element, "sSub", "e", "sub"));
                case "msubsup":
                    return Single(CreateThreePartScript(element, "sSubSup", "e", "sub", "sup"));
                case "mmultiscripts":
                    return ConvertChildren(element);
                case "msqrt":
                    return Single(CreateSquareRoot(element));
                case "mroot":
                    return Single(CreateRoot(element));
                case "munder":
                    return Single(CreateScript(element, "limLow", "e", "lim"));
                case "mover":
                    return Single(CreateScript(element, "limUpp", "e", "lim"));
                case "munderover":
                    return Single(CreateThreePartScript(element, "sSubSup", "e", "sub", "sup"));
                case "mfenced":
                    return Single(CreateDelimiter(element));
                case "mtable":
                    return Single(CreateMatrix(element));
                case "mtr":
                case "mlabeledtr":
                case "mtd":
                case "menclose":
                    return ConvertChildren(element);
                default:
                    List<XElement> descendants = ConvertChildren(element).ToList();
                    return descendants.Count > 0 ? descendants : CreateTextRun(element.TextContent);
            }
        }

        private static IEnumerable<XElement> ConvertChildren(IElement element) =>
            element.Children.SelectMany(ConvertElement);

        private static IEnumerable<XElement> CreateTextRun(string? text) {
            if (string.IsNullOrEmpty(text)) return Enumerable.Empty<XElement>();
            return Single(new XElement(M + "r", new XElement(M + "t", text)));
        }

        private static XElement CreateFraction(IElement element) {
            IElement[] children = element.Children.ToArray();
            var fraction = new XElement(M + "f");
            string? fractionType = GetFractionType(element);
            if (fractionType != null) {
                fraction.Add(new XElement(
                    M + "fPr",
                    new XElement(M + "type", new XAttribute(M + "val", fractionType))));
            }
            fraction.Add(CreateArgument("num", children.ElementAtOrDefault(0)));
            fraction.Add(CreateArgument("den", children.ElementAtOrDefault(1)));
            return fraction;
        }

        private static string? GetFractionType(IElement element) {
            if (string.Equals(element.GetAttribute("bevelled"), "true", StringComparison.OrdinalIgnoreCase)) {
                return "skw";
            }

            string? lineThickness = element.GetAttribute("linethickness")?.Trim();
            if (lineThickness == "0" || lineThickness == "0px" || lineThickness == "0pt") {
                return "noBar";
            }
            return null;
        }

        private static XElement CreateScript(IElement element, string type, string baseName, string scriptName) {
            IElement[] children = element.Children.ToArray();
            return new XElement(
                M + type,
                CreateArgument(baseName, children.ElementAtOrDefault(0)),
                CreateArgument(scriptName, children.ElementAtOrDefault(1)));
        }

        private static XElement CreateThreePartScript(
            IElement element,
            string type,
            string baseName,
            string lowerName,
            string upperName) {
            IElement[] children = element.Children.ToArray();
            return new XElement(
                M + type,
                CreateArgument(baseName, children.ElementAtOrDefault(0)),
                CreateArgument(lowerName, children.ElementAtOrDefault(1)),
                CreateArgument(upperName, children.ElementAtOrDefault(2)));
        }

        private static XElement CreateSquareRoot(IElement element) =>
            new XElement(
                M + "rad",
                new XElement(M + "radPr", new XElement(M + "degHide", new XAttribute(M + "val", "1"))),
                new XElement(M + "e", ConvertChildren(element)));

        private static XElement CreateRoot(IElement element) {
            IElement[] children = element.Children.ToArray();
            return new XElement(
                M + "rad",
                CreateArgument("deg", children.ElementAtOrDefault(1)),
                CreateArgument("e", children.ElementAtOrDefault(0)));
        }

        private static XElement CreateDelimiter(IElement element) {
            var properties = new XElement(M + "dPr");
            AddCharacterProperty(properties, "begChr", element.GetAttribute("open"));
            AddCharacterProperty(properties, "endChr", element.GetAttribute("close"));
            AddCharacterProperty(properties, "sepChr", element.GetAttribute("separators"));

            var delimiter = new XElement(M + "d");
            if (properties.HasElements) delimiter.Add(properties);
            foreach (IElement child in element.Children) {
                delimiter.Add(CreateArgument("e", child));
            }
            return delimiter;
        }

        private static void AddCharacterProperty(XElement properties, string name, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                properties.Add(new XElement(M + name, new XAttribute(M + "val", value)));
            }
        }

        private static XElement CreateMatrix(IElement element) {
            var matrix = new XElement(M + "m");
            foreach (IElement rowElement in element.Children.Where(child =>
                         string.Equals(child.LocalName, "mtr", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(child.LocalName, "mlabeledtr", StringComparison.OrdinalIgnoreCase))) {
                var row = new XElement(M + "mr");
                foreach (IElement cell in rowElement.Children.Where(child =>
                             string.Equals(child.LocalName, "mtd", StringComparison.OrdinalIgnoreCase))) {
                    row.Add(new XElement(M + "e", ConvertChildren(cell)));
                }
                matrix.Add(row);
            }
            return matrix;
        }

        private static XElement CreateArgument(string name, IElement? element) =>
            new XElement(M + name, element == null ? null : ConvertElement(element));

        private static IEnumerable<XElement> Single(XElement element) {
            yield return element;
        }
    }
}
