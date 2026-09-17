using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

// Attribute mutation is the reliable navigation boundary. The loaded document URL
// may differ from the authored source after a redirect and cannot identify replacement.
internal sealed class RuntimeFrameNavigationObserver(RuntimeFrameRealms realms) : IAttributeObserver {
    public void NotifyChange(IElement host, string name, string? value) {
        if (name is "src" or "srcdoc") realms.BeginFrameNavigation(host);
    }
}
