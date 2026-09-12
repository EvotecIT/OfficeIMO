using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlAccessibilitySemantics {
    /// <summary>Checks an owned element's ARIA role tokens.</summary>
    public static bool HasRole(Dom.HtmlElement? element, string role) => HasRole((IElement)NativeDomBridge.GetNativeOrNull(element)!, role);
    /// <summary>Checks an owned element's EPUB semantic type tokens.</summary>
    public static bool HasEpubType(Dom.HtmlElement? element, string semanticType) => HasEpubType((IElement)NativeDomBridge.GetNativeOrNull(element)!, semanticType);
    /// <summary>Reads the semantic heading level from an owned element.</summary>
    public static bool TryGetHeadingLevel(Dom.HtmlElement? element, out int level) => TryGetHeadingLevel((IElement)NativeDomBridge.GetNativeOrNull(element)!, out level);
    /// <summary>Computes an accessible name, optionally using text as the final fallback.</summary>
    public static string GetAccessibleName(Dom.HtmlElement? element, bool includeTextFallback = false) => GetAccessibleName((IElement)NativeDomBridge.GetNativeOrNull(element)!, includeTextFallback);
    /// <summary>Computes the accessible image name from owned document semantics.</summary>
    public static string GetImageAccessibleName(Dom.HtmlElement? element) => GetImageAccessibleName((IElement)NativeDomBridge.GetNativeOrNull(element)!);
    /// <summary>Checks inherited ARIA hidden state.</summary>
    public static bool IsAriaHidden(Dom.HtmlElement? element) => IsAriaHidden((IElement)NativeDomBridge.GetNativeOrNull(element)!);
}
