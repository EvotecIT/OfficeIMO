using System.Threading;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Computes a generated formatting context without adding synthetic DOM nodes.</summary>
    /// <remarks>The caller owns declaration parsing, matching and input budgets. Parent values
    /// remain inheritance fallbacks; only this context's declarations enter its cascade.</remarks>
    internal static HtmlComputedStyle ComputeGeneratedContext(
        HtmlComputedStyle parent,
        IEnumerable<HtmlGeneratedStyleDeclaration> declarations,
        bool enforceResolutionLimits,
        CancellationToken cancellationToken,
        HtmlRenderingFontDefaults? fontDefaults = null,
        IReadOnlyDictionary<string, CustomPropertyRegistration>? customPropertyRegistrations = null) {
        if (parent == null) throw new ArgumentNullException(nameof(parent));
        if (declarations == null) throw new ArgumentNullException(nameof(declarations));
        cancellationToken.ThrowIfCancellationRequested();
        var properties = new Dictionary<string, CascadedProperty>(HtmlCssPropertyNameComparer.Instance);
        foreach (HtmlGeneratedStyleDeclaration declaration in declarations) {
            cancellationToken.ThrowIfCancellationRequested();
            ApplyDeclaration(properties, parent.Properties, declaration.Name, declaration.Value,
                declaration.IsImportant, new Specificity(0, 0, declaration.Specificity),
                declaration.SourceOrder, declaration.LayerOrder,
                declarationOrder: declaration.DeclarationOrder, customPropertyRegistrations: customPropertyRegistrations,
                enforceResolutionLimits: enforceResolutionLimits);
        }

        Dictionary<string, string> resolvedProperties = ResolveComputedProperties(properties, parent.Properties,
            out HashSet<string> inheritedProperties, out HashSet<string> resetProperties,
            out HashSet<string> originRevertedProperties, out HashSet<string> specifiedProperties,
            out Dictionary<string, HtmlCssCascadePriority> cascadePriorities,
            out Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces,
            includeCascadeTraces: false, customPropertyRegistrations: customPropertyRegistrations, enforceResolutionLimits: enforceResolutionLimits);
        cancellationToken.ThrowIfCancellationRequested();
        MaterializeInitialFontValues(resolvedProperties, resetProperties, originRevertedProperties,
            cascadeTraces, fontDefaults);
        return HtmlComputedStyle.FromOwnedCollections(resolvedProperties, inheritedProperties,
            resetProperties, originRevertedProperties, specifiedProperties, cascadePriorities, cascadeTraces);
    }
}

/// <summary>A matched declaration with its generated-context cascade priority.</summary>
internal readonly struct HtmlGeneratedStyleDeclaration {
    internal HtmlGeneratedStyleDeclaration(string name, string value, bool isImportant,
        int specificity, int sourceOrder, CascadeLayerOrder? layerOrder, int declarationOrder) {
        Name = name;
        Value = value;
        IsImportant = isImportant;
        Specificity = specificity;
        SourceOrder = sourceOrder;
        LayerOrder = layerOrder;
        DeclarationOrder = declarationOrder;
    }

    internal string Name { get; }
    internal string Value { get; }
    internal bool IsImportant { get; }
    internal int Specificity { get; }
    internal int SourceOrder { get; }
    internal CascadeLayerOrder? LayerOrder { get; }
    internal int DeclarationOrder { get; }
}
