using System.Runtime.CompilerServices;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

// Interaction measurement attaches an identity to a cloned element without
// changing its attributes, selector matching, form ownership, or generated content.
internal static class HtmlRenderSourceIdentity {
    private static readonly ConditionalWeakTable<IElement, Identity> Identities = new();

    internal static void Register(IElement element, string value) => Identities.Add(element, new Identity(value));

    internal static bool TryGet(IElement element, out string value) {
        if (Identities.TryGetValue(element, out Identity? identity)) {
            value = identity.Value;
            return true;
        }
        value = string.Empty;
        return false;
    }

    private sealed class Identity(string value) {
        internal string Value { get; } = value;
    }
}
