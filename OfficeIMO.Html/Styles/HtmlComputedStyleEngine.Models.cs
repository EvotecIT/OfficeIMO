namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private sealed class StyleDeclaration {
        internal StyleDeclaration(string value, bool isImportant) {
            Value = value;
            IsImportant = isImportant;
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
    }

    private sealed class CascadedProperty {
        internal CascadedProperty(string value, bool isImportant, Specificity specificity, int order) {
            Value = value;
            HasValue = true;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        private CascadedProperty(bool isImportant, Specificity specificity, int order) {
            Value = string.Empty;
            HasValue = false;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        internal static CascadedProperty Clear(bool isImportant, Specificity specificity, int order) {
            return new CascadedProperty(isImportant, specificity, order);
        }

        internal string Value { get; }
        internal bool HasValue { get; }
        internal bool IsImportant { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct CssKeywordResolution {
        private CssKeywordResolution(bool hasValue, string value) {
            HasValue = hasValue;
            Value = value;
        }

        internal static CssKeywordResolution Clear => new CssKeywordResolution(false, string.Empty);
        internal static CssKeywordResolution ForValue(string value) => new CssKeywordResolution(true, value);

        internal bool HasValue { get; }
        internal string Value { get; }
    }

    private sealed class Specificity {
        internal Specificity(int ids, int classesAttributesAndPseudoClasses, int elements) {
            Ids = ids;
            ClassesAttributesAndPseudoClasses = classesAttributesAndPseudoClasses;
            Elements = elements;
        }

        internal int Ids { get; }
        internal int ClassesAttributesAndPseudoClasses { get; }
        internal int Elements { get; }
        internal static Specificity Inherited { get; } = new Specificity(-1, -1, -1);
        internal static Specificity PresentationalHint { get; } = new Specificity(0, 0, 0);
        internal static Specificity Inline { get; } = new Specificity(int.MaxValue, int.MaxValue, int.MaxValue);

        internal int CompareTo(Specificity other) {
            if (Ids != other.Ids) {
                return Ids.CompareTo(other.Ids);
            }

            if (ClassesAttributesAndPseudoClasses != other.ClassesAttributesAndPseudoClasses) {
                return ClassesAttributesAndPseudoClasses.CompareTo(other.ClassesAttributesAndPseudoClasses);
            }

            return Elements.CompareTo(other.Elements);
        }
    }

    private sealed class StyleRule {
        internal StyleRule(string selector, Specificity specificity, int order, IDictionary<string, StyleDeclaration> declarations) {
            Selector = selector;
            Specificity = specificity;
            Order = order;
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, StringComparer.OrdinalIgnoreCase);
        }

        internal string Selector { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal IReadOnlyDictionary<string, StyleDeclaration> Declarations { get; }
    }
}
