using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private PowerPointLayoutTheme? _layoutTheme;
        private OfficeTheme? _baseTheme;

        /// <summary>The theme a new presentation was created with; the default theme for loaded presentations.</summary>
        internal OfficeTheme BaseTheme => _baseTheme ?? OfficeTheme.Default;

        /// <summary>True when OfficeIMO wrote this presentation's master and layouts, so it may rewrite them.</summary>
        internal bool OwnsLayoutTheme => _baseTheme != null;
        private PowerPointLayoutDeck? _layoutDeck;

        /// <summary>The presentation part the layout system reads and restyles; reaching it opens the SDK.</summary>
        internal OpenXml.PowerPointPartRef LayoutPresentationPartRef => OpenXml.PowerPointPartRef.FromSdk(SdkPresentationPart);

        /// <summary>
        ///     Replaces the child of <paramref name="parent"/> that <paramref name="markup"/> names with the element it
        ///     describes, or adds it where the schema places it. The element keeps the namespace declarations the markup
        ///     carries, as an element built from its outer XML does, so the layout system writes the same XML.
        /// </summary>
        internal static OfficeXmlNode ReplaceWithLayoutMarkup(OfficeXmlNode parent, string markup) {
            OpenXmlElement target = parent.SdkElement as OpenXmlElement
                ?? throw new InvalidOperationException("The layout system writes masters and layouts through the SDK.");
            OpenXmlElement probe = target.CloneNode(false);
            probe.InnerXml = markup;
            OpenXmlElement child = probe.FirstChild ?? throw new InvalidOperationException("The markup holds no element.");
            child.Remove();
            OpenXmlElement? existing = target.ChildElements.FirstOrDefault(element =>
                element.LocalName == child.LocalName && element.NamespaceUri == child.NamespaceUri);
            if (existing != null) target.ReplaceChild(child, existing);
            else if (target is OpenXmlCompositeElement composite) composite.AddChild(child, throwOnError: true);
            else target.AppendChild(child);
            return OfficeXmlNode.FromSdk(child);
        }

        /// <summary>Writes a part the layout system restyled back to its package part.</summary>
        internal static void SaveLayoutPart(OpenXml.PowerPointPartRef part) {
            if (part.Current.SdkPart is OpenXmlPart sdk) sdk.RootElement?.Save();
            else part.MarkEdited();
        }

        /// <summary>Writes the theme part, master text styles, placeholders and layouts of a new presentation.</summary>
        internal void ApplyLayoutTheme(OfficeTheme theme) {
            _baseTheme = theme;
            _layoutTheme = PowerPointLayoutTheme.From(theme);
            PowerPointLayoutMasterWriter.Apply(LayoutPresentationPartRef, _layoutTheme);
        }

        /// <summary>
        ///     Restyles a presentation's theme and master text styles. Presentations created by OfficeIMO also get their
        ///     master placeholders and layouts rewritten; a loaded presentation keeps its shapes, backgrounds and placeholders.
        /// </summary>
        /// <param name="theme">Theme to apply. Explicit shape formatting is retained.</param>
        /// <param name="masterIndex">Zero-based master to restyle, or null for every master and the presentation defaults.</param>
        public void ApplyTheme(OfficeTheme theme, int? masterIndex = null) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(theme);
            SlideMasterPart? selected = masterIndex.HasValue ? GetSlideMasterPart(masterIndex.Value) : null;
            bool allMasters = selected == null || SdkPresentationPart.SlideMasterParts.Count() == 1;
            PowerPointLayoutTheme layout = PowerPointLayoutTheme.From(theme);
            PowerPointLayoutMasterWriter.Apply(LayoutPresentationPartRef, layout, allMasters ? default : OpenXml.PowerPointPartRef.FromSdk(selected),
                keepMasterContent: !OwnsLayoutTheme);
            if (allMasters) SetThemeNameForAllMasters(theme.Name);
            else SetThemeName(theme.Name, masterIndex!.Value);
            if (allMasters && OwnsLayoutTheme) _baseTheme = theme;
            if (allMasters || masterIndex == 0) _layoutTheme = layout;
        }

        /// <summary>Reads a master's theme name, current palette and Latin heading/body fonts without restyling the presentation.</summary>
        /// <param name="masterIndex">Zero-based master position.</param>
        public OfficeTheme GetTheme(int masterIndex = 0) {
            ThrowIfDisposed();
            GetSlideMasterPart(masterIndex);
            return ReadLayoutTheme(masterIndex).Theme;
        }

        private void RefreshLayoutTheme(int masterIndex) {
            if (masterIndex == 0) _layoutTheme = ReadLayoutTheme();
        }

        /// <summary>Rewrites the theme-derived parts for a designer theme mapped onto the presentation's theme.</summary>
        internal void ApplyLayoutTheme(PowerPointLayoutTheme theme) {
            _layoutTheme = theme;
            PowerPointLayoutMasterWriter.Apply(LayoutPresentationPartRef, theme);
        }

        /// <summary>
        ///     Tokens used by the slide layout system. New presentations use the theme they were created with; a
        ///     loaded presentation keeps its own colours and fonts on top of the default type scale and grid.
        /// </summary>
        internal PowerPointLayoutTheme LayoutTheme => _layoutTheme ??= ReadLayoutTheme();

        /// <summary>Deck-level state of the layout system: section titles, numbering and footer text.</summary>
        internal PowerPointLayoutDeck LayoutDeck => _layoutDeck ??= new PowerPointLayoutDeck(this);
    }
}
