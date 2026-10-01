using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

/// <summary>Bounded traversal of unique OPC parts, including shared and cyclic relationships.</summary>
internal static class WordPackageParts {
    internal static IEnumerable<OpenXmlPart> Enumerate(OpenXmlPartContainer container, int maximum,
        Func<string, Exception>? limitException = null) {
        var pending = new Stack<OpenXmlPart>(container.Parts.Select(pair => pair.OpenXmlPart));
        var visited = new HashSet<Uri>();
        while (pending.Count > 0) {
            OpenXmlPart part = pending.Pop();
            if (!visited.Add(part.Uri)) continue;
            if (visited.Count > maximum) {
                string message = "The OPC package contains more than " + maximum + " parts.";
                throw limitException?.Invoke(message) ?? new InvalidDataException(message);
            }
            yield return part;
            foreach (IdPartPair child in part.Parts) pending.Push(child.OpenXmlPart);
        }
    }
}
