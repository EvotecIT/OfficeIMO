using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    // Provider metadata describes both retained shorthands and synthesized values.
    // Cache its dependency names for this operation, never across provider contexts.
    internal sealed class ProviderDeclarationQueryCache {
        private readonly Dictionary<IBrowsingContext, Dictionary<string, string[]>> _contexts = new();

        internal ProviderDeclarationQuery For(ICssStyleDeclaration declaration) {
            IBrowsingContext? context = declaration.Parent?.Owner?.Context;
            if (context == null) return default;
            if (!_contexts.TryGetValue(context, out Dictionary<string, string[]>? related)) {
                related = CreateRelatedNames(context.GetFactory<IDeclarationFactory>());
                _contexts.Add(context, related);
            }
            var declaredNames = new HashSet<string>(HtmlCssPropertyNameComparer.Instance);
            foreach (ICssProperty property in declaration) declaredNames.Add(property.Name);
            return new ProviderDeclarationQuery(related, declaredNames);
        }

        private static Dictionary<string, string[]> CreateRelatedNames(IDeclarationFactory factory) {
            var result = new Dictionary<string, string[]>(HtmlCssPropertyNameComparer.Instance);
            foreach (string name in SupportedProperties) {
                var names = new HashSet<string>(HtmlCssPropertyNameComparer.Instance) { name };
                DeclarationInfo info = factory.Create(name);
                // A longhand can be extracted from an explicitly retained shorthand.
                foreach (string shorthand in info.Shorthands) names.Add(shorthand);
                // A synthesized shorthand can consume nested shorthand families.
                var pending = new Stack<string>(info.Longhands);
                var visited = new HashSet<string>(HtmlCssPropertyNameComparer.Instance);
                while (pending.Count != 0) {
                    string child = pending.Pop();
                    if (!visited.Add(child)) continue;
                    names.Add(child);
                    foreach (string longhand in factory.Create(child).Longhands) pending.Push(longhand);
                }
                result.Add(name, names.ToArray());
            }
            return result;
        }
    }

    internal readonly struct ProviderDeclarationQuery {
        private readonly IReadOnlyDictionary<string, string[]>? _related;
        private readonly HashSet<string>? _declaredNames;

        internal ProviderDeclarationQuery(IReadOnlyDictionary<string, string[]> related, HashSet<string> declaredNames) {
            _related = related;
            _declaredNames = declaredNames;
        }

        internal bool CanHaveValue(string propertyName) {
            // Detached declarations lack a context for metadata; keep provider lookup.
            if (_related == null || !_related.TryGetValue(propertyName, out string[]? names)) return true;
            foreach (string name in names) if (_declaredNames!.Contains(name)) return true;
            return false;
        }
    }
}
