using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

/// <summary>Stable private identities distinguish conditional branches and repeated macro invocations without modifying CSL XML.</summary>
internal sealed class CslElementIdentity {
    private CslElementIdentity(int number) { Key = number.ToString("D10", CultureInfo.InvariantCulture); }
    internal string Key { get; }

    internal static void Assign(XElement root, CancellationToken token) {
        int number = 0;
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            if (element.Attribute("disambiguate") != null || element.Name.LocalName == "text" && element.Attribute("macro") != null)
                element.AddAnnotation(new CslElementIdentity(number++));
        }
    }

    /// <summary>Copies evaluation nodes while retaining their immutable identities; XElement's copy constructor omits annotations.</summary>
    internal static XElement Copy(XElement source) {
        var copy = new XElement(source);
        using IEnumerator<XElement> originals = source.DescendantsAndSelf().GetEnumerator();
        foreach (XElement element in copy.DescendantsAndSelf()) {
            originals.MoveNext();
            if (originals.Current.Annotation<CslElementIdentity>() is CslElementIdentity identity) element.AddAnnotation(identity);
        }
        return copy;
    }
}
