namespace OfficeIMO.Html;

internal static partial class HtmlConversionInputGuard {
    // Validate attached input before cloning or native projection; detached editing history is not conversion input.
    internal static Dom.HtmlDocument CaptureOwnedTree(Dom.HtmlDocument document, HtmlConversionLimits limits, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        limits.Validate();
        var tracker = HtmlDomLimitTracker.Create(limits.MaxHtmlNodes, limits.MaxHtmlDepth);
        long characters = 0;
        int nodes = 0;
        void Reserve(string? value) {
            characters += value?.Length ?? 0;
            HtmlConversionSourceWriter.ValidateLength(characters, limits);
        }
        foreach (var current in document.AttachedNodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            nodes++;
            if (ReferenceEquals(current.Node, document)) continue;
            if (current.Node is Dom.HtmlElement element) {
                tracker?.RecordElementStart(current.Depth);
                Reserve(element.LocalName);
                Reserve(element.Prefix);
                Reserve(element.FormState?.Value);
                foreach (Dom.HtmlAttribute attribute in element.Attributes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    Reserve(attribute.Name);
                    Reserve(attribute.Value);
                }
            } else {
                tracker?.RecordNode();
                if (current.Node is Dom.HtmlDocumentType type) {
                    Reserve(type.Name);
                    Reserve(type.PublicIdentifier);
                    Reserve(type.SystemIdentifier);
                } else Reserve(current.Node.Data);
            }
        }
        return document.IsReadOnly && nodes == document.RegisteredNodeCount
            ? document
            : document.CloneAttached(cancellationToken).Freeze();
    }

}
