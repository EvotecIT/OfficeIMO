using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes the root default tab width while preserving untouched RTF syntax.
    /// </summary>
    public void SetDefaultTabWidth(int? twips) {
        ValidateNonNegative(twips, nameof(twips));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "deftab", twips, SingleControlName("deftab"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root default language controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetDefaultLanguages(int? defaultLanguageId = null, int? defaultFarEastLanguageId = null, int? defaultAlternateLanguageId = null) {
        ValidateNonNegative(defaultLanguageId, nameof(defaultLanguageId));
        ValidateNonNegative(defaultFarEastLanguageId, nameof(defaultFarEastLanguageId));
        ValidateNonNegative(defaultAlternateLanguageId, nameof(defaultAlternateLanguageId));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "deflang", defaultLanguageId, SingleControlName("deflang"));
        root = SetRootControl(root, "deflangfe", defaultFarEastLanguageId, SingleControlName("deflangfe"));
        root = SetRootControl(root, "adeflang", defaultAlternateLanguageId, SingleControlName("adeflang"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root document view controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetDocumentView(int? kind = null, int? scale = null, int? zoomKind = null, int? backspaceBehavior = null) {
        ValidateNonNegative(kind, nameof(kind));
        ValidateNonNegative(scale, nameof(scale));
        ValidateNonNegative(zoomKind, nameof(zoomKind));
        ValidateNonNegative(backspaceBehavior, nameof(backspaceBehavior));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "viewkind", kind, SingleControlName("viewkind"));
        root = SetRootControl(root, "viewscale", scale, SingleControlName("viewscale"));
        root = SetRootControl(root, "viewzk", zoomKind, SingleControlName("viewzk"));
        root = SetRootControl(root, "viewbksp", backspaceBehavior, SingleControlName("viewbksp"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root document layout option controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetDocumentLayoutOptions(bool? widowOrphanControl = null, bool? facingPages = null, bool? mirrorMargins = null) {
        RtfGroup root = SetRootOptionalToggleControl(_syntaxTree.Root, "widowctrl", widowOrphanControl);
        root = SetRootOptionalToggleControl(root, "facingp", facingPages);
        root = SetRootOptionalToggleControl(root, "margmirror", mirrorMargins);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root document hyphenation controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetDocumentHyphenation(bool? automatic = null, bool? caps = null, int? consecutiveLimit = null, int? zoneTwips = null) {
        ValidateNonNegative(consecutiveLimit, nameof(consecutiveLimit));
        ValidateNonNegative(zoneTwips, nameof(zoneTwips));

        RtfGroup root = SetRootOptionalToggleControl(_syntaxTree.Root, "hyphauto", automatic);
        root = SetRootOptionalToggleControl(root, "hyphcaps", caps);
        root = SetRootControl(root, "hyphconsec", consecutiveLimit, SingleControlName("hyphconsec"));
        root = SetRootControl(root, "hyphhotz", zoneTwips, SingleControlName("hyphhotz"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root document protection controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetDocumentProtection(bool? forms = null, bool? revisions = null, bool? annotations = null, bool? readOnly = null) {
        RtfGroup root = SetRootOptionalToggleControl(_syntaxTree.Root, "formprot", forms);
        root = SetRootOptionalToggleControl(root, "revprot", revisions);
        root = SetRootOptionalToggleControl(root, "annotprot", annotations);
        root = SetRootOptionalToggleControl(root, "readprot", readOnly);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root revision tracking controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetRevisionTracking(bool? enabled = null, int? displayStyle = null, int? barPlacement = null) {
        ValidateNonNegative(displayStyle, nameof(displayStyle));
        ValidateNonNegative(barPlacement, nameof(barPlacement));

        RtfGroup root = SetRootOptionalToggleControl(_syntaxTree.Root, "revisions", enabled);
        root = SetRootControl(root, "revprop", displayStyle, SingleControlName("revprop"));
        root = SetRootControl(root, "revbar", barPlacement, SingleControlName("revbar"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root drawing grid controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetDrawingGrid(
        int? horizontalSpacingTwips = null,
        int? verticalSpacingTwips = null,
        int? horizontalOriginTwips = null,
        int? verticalOriginTwips = null,
        int? horizontalShow = null,
        int? verticalShow = null,
        bool? snapToGrid = null,
        bool? useMargins = null) {
        ValidateNonNegative(horizontalSpacingTwips, nameof(horizontalSpacingTwips));
        ValidateNonNegative(verticalSpacingTwips, nameof(verticalSpacingTwips));
        ValidateNonNegative(horizontalOriginTwips, nameof(horizontalOriginTwips));
        ValidateNonNegative(verticalOriginTwips, nameof(verticalOriginTwips));
        ValidateNonNegative(horizontalShow, nameof(horizontalShow));
        ValidateNonNegative(verticalShow, nameof(verticalShow));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "dghspace", horizontalSpacingTwips, SingleControlName("dghspace"));
        root = SetRootControl(root, "dgvspace", verticalSpacingTwips, SingleControlName("dgvspace"));
        root = SetRootControl(root, "dghorigin", horizontalOriginTwips, SingleControlName("dghorigin"));
        root = SetRootControl(root, "dgvorigin", verticalOriginTwips, SingleControlName("dgvorigin"));
        root = SetRootControl(root, "dghshow", horizontalShow, SingleControlName("dghshow"));
        root = SetRootControl(root, "dgvshow", verticalShow, SingleControlName("dgvshow"));
        root = SetRootOptionalToggleControl(root, "dgsnap", snapToGrid);
        root = SetRootOptionalToggleControl(root, "dgmargin", useMargins);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes the root document direction control while preserving untouched RTF syntax.
    /// </summary>
    public void SetDocumentDirection(RtfTextDirection? direction) {
        string? controlName = direction switch {
            RtfTextDirection.LeftToRight => "ltrdoc",
            RtfTextDirection.RightToLeft => "rtldoc",
            null => null,
            _ => throw new ArgumentOutOfRangeException(nameof(direction), "Unsupported RTF document direction.")
        };

        RtfGroup root = controlName == null
            ? RemoveRootDirectionalControls(_syntaxTree.Root)
            : SetRootControl(_syntaxTree.Root, controlName, null, DirectionControlNames);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetRootOptionalToggleControl(RtfGroup root, string controlName, bool? value) {
        if (!value.HasValue) {
            var children = new List<RtfNode>(root.Children);
            RemoveRootControls(children, SingleControlName(controlName));
            return new RtfGroup(root.Position, children);
        }

        return SetRootControl(root, controlName, value.Value ? null : 0, SingleControlName(controlName));
    }

    private static RtfGroup RemoveRootDirectionalControls(RtfGroup root) {
        var children = new List<RtfNode>(root.Children);
        RemoveRootControls(children, DirectionControlNames);
        return new RtfGroup(root.Position, children);
    }

    private static readonly ISet<string> DirectionControlNames = new HashSet<string>(StringComparer.Ordinal) {
        "ltrdoc",
        "rtldoc"
    };
}
