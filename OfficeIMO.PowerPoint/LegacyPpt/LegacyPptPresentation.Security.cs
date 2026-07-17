namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        internal bool HasVbaContent { get; private set; }
        internal bool HasEmbeddedOleContent { get; private set; }
        internal bool HasLinkedOleContent { get; private set; }
        internal bool HasActiveXContent { get; private set; }
        internal bool HasExternalHyperlinkContent { get; private set; }
        internal bool HasExternalMediaContent { get; private set; }
    }
}
