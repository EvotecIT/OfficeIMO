using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static bool RewriteMovedXmlStylesheets(XDocument document, string owner, string destination, string oldPath, string newPath, CancellationToken token) {
        bool changed = false;
        foreach (XProcessingInstruction instruction in document.DescendantNodes().OfType<XProcessingInstruction>()) {
            token.ThrowIfCancellationRequested();
            if (instruction.Target != "xml-stylesheet")
                throw new NotSupportedException("Resource renaming cannot inspect an unknown XML processing instruction: " + instruction.Target);
            // Use the platform XML parser to validate pseudo-attribute syntax and decode entities.
            // Replace only href's lexical value, retaining other pseudo-attributes verbatim (including whitespace).
            XElement attributes;
            using (var reader = XmlReader.Create(new StringReader("<style " + instruction.Data + "/>"), new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = (long)instruction.Data.Length + 32
            })) attributes = XElement.Load(reader, LoadOptions.PreserveWhitespace);
            string href = (string?)attributes.Attribute("href") ?? throw new InvalidDataException("XML stylesheet instruction has no href.");
            string replacement = RewriteMovedReference(owner, null, destination, null, href, oldPath, newPath);
            if (replacement == href) continue;
            const string pattern = @"(?<name>[^\s=]+)\s*=\s*(?:'(?<single>[^']*)'|""(?<double>[^""]*)"")";
            Match match = Regex.Matches(instruction.Data, pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1))
                .Cast<Match>().Single(item => item.Groups["name"].Value == "href");
            Group value = match.Groups["single"].Success ? match.Groups["single"] : match.Groups["double"];
            string escaped = replacement.Replace("&", "&amp;").Replace("<", "&lt;").Replace("\"", "&quot;").Replace("'", "&apos;");
            instruction.Data = instruction.Data.Substring(0, value.Index) + escaped + instruction.Data.Substring(value.Index + value.Length);
            changed = true;
        }
        return changed;
    }
}
