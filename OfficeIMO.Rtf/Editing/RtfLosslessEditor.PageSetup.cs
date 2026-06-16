using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes root document paper size controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetPageSize(int? widthTwips, int? heightTwips) {
        ValidatePositive(widthTwips, nameof(widthTwips), "Paper width must be greater than zero.");
        ValidatePositive(heightTwips, nameof(heightTwips), "Paper height must be greater than zero.");

        RtfGroup root = SetRootControl(_syntaxTree.Root, "paperw", widthTwips, SingleControlName("paperw"));
        root = SetRootControl(root, "paperh", heightTwips, SingleControlName("paperh"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root printer paper metadata controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetPrinterPaper(int? paperSize = null, int? firstPageSource = null, int? otherPagesSource = null) {
        ValidateNonNegative(paperSize, nameof(paperSize));
        ValidateNonNegative(firstPageSource, nameof(firstPageSource));
        ValidateNonNegative(otherPagesSource, nameof(otherPagesSource));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "psz", paperSize, SingleControlName("psz"));
        root = SetRootControl(root, "binfsxn", firstPageSource, SingleControlName("binfsxn"));
        root = SetRootControl(root, "binsxn", otherPagesSource, SingleControlName("binsxn"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root document margin controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetMargins(int? leftTwips = null, int? rightTwips = null, int? topTwips = null, int? bottomTwips = null) {
        ValidateNonNegative(leftTwips, nameof(leftTwips));
        ValidateNonNegative(rightTwips, nameof(rightTwips));
        ValidateNonNegative(topTwips, nameof(topTwips));
        ValidateNonNegative(bottomTwips, nameof(bottomTwips));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "margl", leftTwips, SingleControlName("margl"));
        root = SetRootControl(root, "margr", rightTwips, SingleControlName("margr"));
        root = SetRootControl(root, "margt", topTwips, SingleControlName("margt"));
        root = SetRootControl(root, "margb", bottomTwips, SingleControlName("margb"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes the root document gutter width while preserving untouched RTF syntax.
    /// </summary>
    public void SetGutterWidth(int? gutterWidthTwips) {
        ValidateNonNegative(gutterWidthTwips, nameof(gutterWidthTwips));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "gutter", gutterWidthTwips, SingleControlName("gutter"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes root header and footer distance controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetHeaderFooterDistance(int? headerDistanceTwips = null, int? footerDistanceTwips = null) {
        ValidateNonNegative(headerDistanceTwips, nameof(headerDistanceTwips));
        ValidateNonNegative(footerDistanceTwips, nameof(footerDistanceTwips));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "headery", headerDistanceTwips, SingleControlName("headery"));
        root = SetRootControl(root, "footery", footerDistanceTwips, SingleControlName("footery"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or removes the root landscape orientation control while preserving untouched RTF syntax.
    /// </summary>
    public void SetLandscape(bool enabled = true) {
        RtfGroup root = SetRootToggleControl(_syntaxTree.Root, "landscape", enabled);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or removes the root different-first-page header/footer control while preserving untouched RTF syntax.
    /// </summary>
    public void SetDifferentFirstPageHeaderFooter(bool enabled = true) {
        RtfGroup root = SetRootToggleControl(_syntaxTree.Root, "titlepg", enabled);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or removes the root right-to-left gutter control while preserving untouched RTF syntax.
    /// </summary>
    public void SetRtlGutter(bool enabled = true) {
        RtfGroup root = SetRootToggleControl(_syntaxTree.Root, "rtlgutter", enabled);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetRootToggleControl(RtfGroup root, string controlName, bool enabled) {
        if (enabled) {
            return SetRootControl(root, controlName, null, SingleControlName(controlName));
        }

        var children = new List<RtfNode>(root.Children);
        RemoveRootControls(children, SingleControlName(controlName));
        return new RtfGroup(root.Position, children);
    }

    private static void ValidatePositive(int? value, string parameterName, string message) {
        if (value.HasValue && value.Value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, message);
        }
    }
}
