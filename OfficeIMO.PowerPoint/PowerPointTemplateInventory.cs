using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>Semantic role inferred for a template placeholder.</summary>
    public enum PowerPointTemplatePlaceholderRole {
        /// <summary>Role could not be inferred.</summary>
        Unknown,
        /// <summary>Primary slide title.</summary>
        Title,
        /// <summary>Subtitle or secondary heading.</summary>
        Subtitle,
        /// <summary>Body or list content.</summary>
        Body,
        /// <summary>Generic content placeholder.</summary>
        Content,
        /// <summary>Picture, screenshot, or logo placeholder.</summary>
        Image,
        /// <summary>Chart placeholder.</summary>
        Chart,
        /// <summary>Table placeholder.</summary>
        Table,
        /// <summary>Footer text placeholder.</summary>
        Footer,
        /// <summary>Date/time placeholder.</summary>
        Date,
        /// <summary>Slide-number placeholder.</summary>
        SlideNumber
    }

    /// <summary>Kind of visual asset found on a master or layout.</summary>
    public enum PowerPointTemplateAssetKind {
        /// <summary>General embedded picture.</summary>
        Picture,
        /// <summary>Picture whose name or description identifies it as a likely logo.</summary>
        Logo
    }

    /// <summary>Template placeholder with its semantic role and authored bounds.</summary>
    public sealed class PowerPointTemplatePlaceholderInfo {
        internal PowerPointTemplatePlaceholderInfo(string name, PlaceholderValues? placeholderType,
            uint? placeholderIndex, PowerPointTemplatePlaceholderRole role, PowerPointLayoutBox? bounds,
            string? defaultText) {
            Name = name ?? string.Empty;
            PlaceholderType = placeholderType;
            PlaceholderIndex = placeholderIndex;
            Role = role;
            Bounds = bounds;
            DefaultText = defaultText;
        }

        /// <summary>Authored placeholder name.</summary>
        public string Name { get; }

        /// <summary>Native PowerPoint placeholder type.</summary>
        public PlaceholderValues? PlaceholderType { get; }

        /// <summary>Native placeholder index.</summary>
        public uint? PlaceholderIndex { get; }

        /// <summary>Inferred semantic role.</summary>
        public PowerPointTemplatePlaceholderRole Role { get; }

        /// <summary>Authored placeholder bounds.</summary>
        public PowerPointLayoutBox? Bounds { get; }

        /// <summary>Default text carried by the layout placeholder, when present.</summary>
        public string? DefaultText { get; }

        /// <inheritdoc />
        public override string ToString() => string.IsNullOrWhiteSpace(Name) ? Role.ToString() : Name;
    }

    /// <summary>Picture asset discovered on a template master or layout.</summary>
    public sealed class PowerPointTemplateAssetInfo {
        internal PowerPointTemplateAssetInfo(PowerPointTemplateAssetKind kind, int masterIndex, int? layoutIndex,
            string name, string? description, string? contentType, PowerPointLayoutBox? bounds) {
            Kind = kind;
            MasterIndex = masterIndex;
            LayoutIndex = layoutIndex;
            Name = name ?? string.Empty;
            Description = description;
            ContentType = contentType;
            Bounds = bounds;
        }

        /// <summary>Inferred asset kind.</summary>
        public PowerPointTemplateAssetKind Kind { get; }

        /// <summary>Owning master index.</summary>
        public int MasterIndex { get; }

        /// <summary>Owning layout index, or null for a master-level asset.</summary>
        public int? LayoutIndex { get; }

        /// <summary>Authored shape name.</summary>
        public string Name { get; }

        /// <summary>Authored alternative-text description.</summary>
        public string? Description { get; }

        /// <summary>Embedded image content type, when resolvable.</summary>
        public string? ContentType { get; }

        /// <summary>Authored asset bounds.</summary>
        public PowerPointLayoutBox? Bounds { get; }
    }

    /// <summary>One slide layout discovered in a template.</summary>
    public sealed class PowerPointTemplateLayoutInfo {
        private readonly ReadOnlyCollection<PowerPointTemplatePlaceholderInfo> _placeholders;

        internal PowerPointTemplateLayoutInfo(int masterIndex, int layoutIndex, string name,
            SlideLayoutValues? type, IList<PowerPointTemplatePlaceholderInfo> placeholders,
            PowerPointLayoutBox safeArea, PowerPointLayoutBox? titleArea) {
            MasterIndex = masterIndex;
            LayoutIndex = layoutIndex;
            Name = name ?? string.Empty;
            Type = type;
            _placeholders = new ReadOnlyCollection<PowerPointTemplatePlaceholderInfo>(
                new List<PowerPointTemplatePlaceholderInfo>(placeholders));
            SafeArea = safeArea;
            TitleArea = titleArea;
        }

        /// <summary>Owning master index.</summary>
        public int MasterIndex { get; }

        /// <summary>Layout index within its master.</summary>
        public int LayoutIndex { get; }

        /// <summary>Authored layout name.</summary>
        public string Name { get; }

        /// <summary>Native layout type, when declared.</summary>
        public SlideLayoutValues? Type { get; }

        /// <summary>Layout placeholders in authored order.</summary>
        public IReadOnlyList<PowerPointTemplatePlaceholderInfo> Placeholders => _placeholders;

        /// <summary>Best content safe area derived from non-footer placeholders or a slide-margin fallback.</summary>
        public PowerPointLayoutBox SafeArea { get; }

        /// <summary>Title safe area, when a title placeholder is present.</summary>
        public PowerPointLayoutBox? TitleArea { get; }

        /// <summary>Resolves one placeholder by semantic role and rejects ambiguous matches.</summary>
        public PowerPointTemplatePlaceholderInfo ResolvePlaceholder(PowerPointTemplatePlaceholderRole role) {
            List<PowerPointTemplatePlaceholderInfo> matches = _placeholders
                .Where(placeholder => placeholder.Role == role).ToList();
            return ResolveUnique(matches, "Template.PlaceholderNotFound", "Template.PlaceholderAmbiguous",
                "placeholder role '" + role + "'");
        }

        /// <summary>Resolves one placeholder by authored or semantic name and rejects ambiguous matches.</summary>
        public PowerPointTemplatePlaceholderInfo ResolvePlaceholder(string semanticName) {
            if (string.IsNullOrWhiteSpace(semanticName)) {
                throw new ArgumentException("Placeholder name cannot be empty.", nameof(semanticName));
            }

            List<PowerPointTemplatePlaceholderInfo> exact = _placeholders.Where(placeholder =>
                string.Equals(placeholder.Name, semanticName, StringComparison.OrdinalIgnoreCase)).ToList();
            if (exact.Count > 0) {
                return ResolveUnique(exact, "Template.PlaceholderNotFound", "Template.PlaceholderAmbiguous",
                    "placeholder '" + semanticName + "'");
            }

            string normalized = NormalizeName(semanticName);
            if (Enum.TryParse(semanticName, true, out PowerPointTemplatePlaceholderRole role)) {
                List<PowerPointTemplatePlaceholderInfo> roleMatches = _placeholders
                    .Where(placeholder => placeholder.Role == role).ToList();
                if (roleMatches.Count > 0) {
                    return ResolveUnique(roleMatches, "Template.PlaceholderNotFound",
                        "Template.PlaceholderAmbiguous", "placeholder role '" + role + "'");
                }
            }

            List<PowerPointTemplatePlaceholderInfo> matches = _placeholders.Where(placeholder =>
                NormalizeName(placeholder.Name).Contains(normalized)).ToList();
            return ResolveUnique(matches, "Template.PlaceholderNotFound", "Template.PlaceholderAmbiguous",
                "placeholder '" + semanticName + "'");
        }

        /// <inheritdoc />
        public override string ToString() => string.IsNullOrWhiteSpace(Name) ? Type?.ToString() ?? "Layout" : Name;

        private static PowerPointTemplatePlaceholderInfo ResolveUnique(
            IList<PowerPointTemplatePlaceholderInfo> matches, string notFoundCode, string ambiguousCode,
            string description) {
            if (matches.Count == 0) {
                throw new PowerPointTemplateResolutionException(notFoundCode,
                    "No " + description + " was found.", Array.Empty<string>());
            }
            if (matches.Count > 1) {
                throw new PowerPointTemplateResolutionException(ambiguousCode,
                    "More than one " + description + " matched. Use an authored placeholder name or index.",
                    matches.Select(match => match.Name).ToArray());
            }
            return matches[0];
        }

        internal static string NormalizeName(string value) =>
            new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    }

    /// <summary>One master and its layouts, theme tokens, and identity content.</summary>
    public sealed class PowerPointTemplateMasterInfo {
        private readonly ReadOnlyCollection<PowerPointTemplateLayoutInfo> _layouts;
        private readonly ReadOnlyDictionary<PowerPointThemeColor, string> _themeColors;

        internal PowerPointTemplateMasterInfo(int masterIndex, string name, string themeName,
            IDictionary<PowerPointThemeColor, string> themeColors, PowerPointThemeFontSet themeFonts,
            IList<PowerPointTemplateLayoutInfo> layouts) {
            MasterIndex = masterIndex;
            Name = name ?? string.Empty;
            ThemeName = themeName ?? string.Empty;
            _themeColors = new ReadOnlyDictionary<PowerPointThemeColor, string>(
                new Dictionary<PowerPointThemeColor, string>(themeColors));
            ThemeFonts = themeFonts;
            _layouts = new ReadOnlyCollection<PowerPointTemplateLayoutInfo>(
                new List<PowerPointTemplateLayoutInfo>(layouts));
        }

        /// <summary>Zero-based master index.</summary>
        public int MasterIndex { get; }

        /// <summary>Authored master name.</summary>
        public string Name { get; }

        /// <summary>Theme name.</summary>
        public string ThemeName { get; }

        /// <summary>Resolved theme color tokens.</summary>
        public IReadOnlyDictionary<PowerPointThemeColor, string> ThemeColors => _themeColors;

        /// <summary>Theme font tokens.</summary>
        public PowerPointThemeFontSet ThemeFonts { get; }

        /// <summary>Layouts owned by the master.</summary>
        public IReadOnlyList<PowerPointTemplateLayoutInfo> Layouts => _layouts;
    }

    /// <summary>Small, immutable inventory of a PowerPoint template.</summary>
    public sealed class PowerPointTemplateInventory {
        private readonly ReadOnlyCollection<PowerPointTemplateMasterInfo> _masters;
        private readonly ReadOnlyCollection<PowerPointTemplateAssetInfo> _assets;
        private readonly ReadOnlyCollection<string> _footerContents;

        internal PowerPointTemplateInventory(string? sourcePath, int sourceSlideCount, PowerPointLayoutBox slideBounds,
            IList<PowerPointTemplateMasterInfo> masters, IList<PowerPointTemplateAssetInfo> assets,
            IList<string> footerContents) {
            SourcePath = sourcePath;
            SourceSlideCount = sourceSlideCount;
            SlideBounds = slideBounds;
            _masters = new ReadOnlyCollection<PowerPointTemplateMasterInfo>(
                new List<PowerPointTemplateMasterInfo>(masters));
            _assets = new ReadOnlyCollection<PowerPointTemplateAssetInfo>(
                new List<PowerPointTemplateAssetInfo>(assets));
            _footerContents = new ReadOnlyCollection<string>(footerContents
                .Where(text => !string.IsNullOrWhiteSpace(text)).Distinct(StringComparer.Ordinal).ToList());
        }

        /// <summary>Template source path, when inventory was read from a file.</summary>
        public string? SourcePath { get; }

        /// <summary>Number of source slides carried by the template.</summary>
        public int SourceSlideCount { get; }

        /// <summary>Template slide canvas bounds.</summary>
        public PowerPointLayoutBox SlideBounds { get; }

        /// <summary>Template masters and layouts.</summary>
        public IReadOnlyList<PowerPointTemplateMasterInfo> Masters => _masters;

        /// <summary>Pictures discovered on masters and layouts.</summary>
        public IReadOnlyList<PowerPointTemplateAssetInfo> Assets => _assets;

        /// <summary>Assets identified as likely logos.</summary>
        public IReadOnlyList<PowerPointTemplateAssetInfo> LikelyLogos =>
            _assets.Where(asset => asset.Kind == PowerPointTemplateAssetKind.Logo).ToList();

        /// <summary>Distinct footer text discovered on template layouts.</summary>
        public IReadOnlyList<string> FooterContents => _footerContents;

        /// <summary>Resolves one layout by authored name, semantic name fragment, or native type name.</summary>
        public PowerPointTemplateLayoutInfo ResolveLayout(string semanticName) {
            if (string.IsNullOrWhiteSpace(semanticName)) {
                throw new ArgumentException("Layout name cannot be empty.", nameof(semanticName));
            }
            List<PowerPointTemplateLayoutInfo> layouts = _masters.SelectMany(master => master.Layouts).ToList();
            List<PowerPointTemplateLayoutInfo> exact = layouts.Where(layout =>
                string.Equals(layout.Name, semanticName, StringComparison.OrdinalIgnoreCase)).ToList();
            if (exact.Count == 1) return exact[0];
            if (exact.Count > 1) throw CreateAmbiguousLayout(semanticName, exact);

            string normalized = PowerPointTemplateLayoutInfo.NormalizeName(semanticName);
            List<PowerPointTemplateLayoutInfo> matches = layouts.Where(layout =>
                PowerPointTemplateLayoutInfo.NormalizeName(layout.Name).Contains(normalized) ||
                PowerPointTemplateLayoutInfo.NormalizeName(layout.Type?.ToString() ?? string.Empty) == normalized)
                .ToList();
            if (matches.Count == 0) {
                throw new PowerPointTemplateResolutionException("Template.LayoutNotFound",
                    "No template layout matched '" + semanticName + "'.",
                    layouts.Select(layout => layout.Name).ToArray());
            }
            if (matches.Count > 1) throw CreateAmbiguousLayout(semanticName, matches);
            return matches[0];
        }

        /// <summary>Creates a reusable designer brief from imported theme colors, fonts, and identity content.</summary>
        public PowerPointDesignBrief CreateDesignBrief(string seed, string? purpose = null) {
            PowerPointTemplateMasterInfo master = _masters.FirstOrDefault()
                ?? throw new InvalidOperationException("Template inventory has no slide masters.");
            string accent = GetColor(master, PowerPointThemeColor.Accent1, "4472C4");
            PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand(accent, seed, purpose)
                .WithIdentity(string.IsNullOrWhiteSpace(master.ThemeName) ? master.Name : master.ThemeName,
                    footerLeft: _footerContents.FirstOrDefault())
                .WithPalette(
                    GetColorOrNull(master, PowerPointThemeColor.Accent2),
                    GetColorOrNull(master, PowerPointThemeColor.Accent3),
                    GetColorOrNull(master, PowerPointThemeColor.Accent4),
                    GetColorOrNull(master, PowerPointThemeColor.Light2),
                    GetColorOrNull(master, PowerPointThemeColor.Accent5))
                .WithFonts(BlankToNull(master.ThemeFonts.MajorLatin), BlankToNull(master.ThemeFonts.MinorLatin));
            return brief;
        }

        /// <summary>Applies imported theme tokens to every master in a target presentation.</summary>
        public void ApplyBrandTo(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            PowerPointTemplateMasterInfo master = _masters.FirstOrDefault()
                ?? throw new InvalidOperationException("Template inventory has no slide masters.");
            if (!string.IsNullOrWhiteSpace(master.ThemeName)) {
                presentation.SetThemeNameForAllMasters(master.ThemeName);
            }
            if (master.ThemeColors.Count > 0) {
                presentation.SetThemeColorsForAllMasters(master.ThemeColors.ToDictionary(pair => pair.Key,
                    pair => pair.Value));
            }
            presentation.SetThemeFontsForAllMasters(new PowerPointThemeFontSet(
                BlankToNull(master.ThemeFonts.MajorLatin),
                BlankToNull(master.ThemeFonts.MinorLatin),
                BlankToNull(master.ThemeFonts.MajorEastAsian),
                BlankToNull(master.ThemeFonts.MinorEastAsian),
                BlankToNull(master.ThemeFonts.MajorComplexScript),
                BlankToNull(master.ThemeFonts.MinorComplexScript)), keepExistingWhenNull: true);
        }

        private static string GetColor(PowerPointTemplateMasterInfo master, PowerPointThemeColor color,
            string fallback) => master.ThemeColors.TryGetValue(color, out string? value) ? value : fallback;

        private static string? GetColorOrNull(PowerPointTemplateMasterInfo master, PowerPointThemeColor color) =>
            master.ThemeColors.TryGetValue(color, out string? value) ? value : null;

        private static string? BlankToNull(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

        private static PowerPointTemplateResolutionException CreateAmbiguousLayout(string semanticName,
            IList<PowerPointTemplateLayoutInfo> matches) =>
            new PowerPointTemplateResolutionException("Template.LayoutAmbiguous",
                "More than one template layout matched '" + semanticName +
                "'. Use the authored name with a specific master index.",
                matches.Select(layout => layout.MasterIndex + ":" + layout.LayoutIndex + " " + layout.Name)
                    .ToArray());
    }

    /// <summary>Raised when semantic template selection is missing or ambiguous.</summary>
    public sealed class PowerPointTemplateResolutionException : InvalidOperationException {
        internal PowerPointTemplateResolutionException(string code, string message, IEnumerable<string> candidates)
            : base(message) {
            Code = code;
            Candidates = new ReadOnlyCollection<string>((candidates ?? Array.Empty<string>()).ToList());
        }

        /// <summary>Stable resolution failure code.</summary>
        public string Code { get; }

        /// <summary>Candidate names useful for correction or UI display.</summary>
        public IReadOnlyList<string> Candidates { get; }
    }
}
