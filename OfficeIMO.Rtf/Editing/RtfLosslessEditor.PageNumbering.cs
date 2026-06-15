using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes root page-numbering controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetPageNumbering(
        int? start = null,
        bool? restart = null,
        RtfPageNumberFormat? format = null,
        int? positionXTwips = null,
        int? positionYTwips = null) {
        ValidatePositive(start, nameof(start), "Page number start must be greater than zero.");
        ValidateNonNegative(positionXTwips, nameof(positionXTwips));
        ValidateNonNegative(positionYTwips, nameof(positionYTwips));
        ValidatePageNumberFormat(format);

        RtfGroup root = SetRootControl(_syntaxTree.Root, "pgnstarts", start, SingleControlName("pgnstarts"));
        root = SetRootOptionalParameterlessControl(
            root,
            restart.HasValue ? (restart.Value ? "pgnrestart" : "pgncont") : null,
            PageNumberRestartControlNames);
        root = SetRootControl(root, "pgnx", positionXTwips, SingleControlName("pgnx"));
        root = SetRootControl(root, "pgny", positionYTwips, SingleControlName("pgny"));
        root = SetRootOptionalParameterlessControl(
            root,
            format.HasValue ? GetPageNumberFormatControlName(format.Value) : null,
            PageNumberFormatControlNames);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetRootOptionalParameterlessControl(RtfGroup root, string? controlName, ISet<string> replacedControlNames) {
        if (controlName == null) {
            var children = new List<RtfNode>(root.Children);
            RemoveRootControls(children, replacedControlNames);
            return new RtfGroup(root.Position, children);
        }

        return SetRootControl(root, controlName, null, replacedControlNames);
    }

    private static string GetPageNumberFormatControlName(RtfPageNumberFormat format) {
        switch (format) {
            case RtfPageNumberFormat.UpperRoman:
                return "pgnucrm";
            case RtfPageNumberFormat.LowerRoman:
                return "pgnlcrm";
            case RtfPageNumberFormat.UpperLetter:
                return "pgnucltr";
            case RtfPageNumberFormat.LowerLetter:
                return "pgnlcltr";
            case RtfPageNumberFormat.DoubleByteDecimal:
                return "pgndecd";
            case RtfPageNumberFormat.Decimal:
                return "pgndec";
            default:
                throw new ArgumentOutOfRangeException(nameof(format), "Unsupported RTF page-number format.");
        }
    }

    private static void ValidatePageNumberFormat(RtfPageNumberFormat? format) {
        if (!format.HasValue) {
            return;
        }

        GetPageNumberFormatControlName(format.Value);
    }

    private static readonly ISet<string> PageNumberRestartControlNames = new HashSet<string>(StringComparer.Ordinal) {
        "pgnrestart",
        "pgncont"
    };

    private static readonly ISet<string> PageNumberFormatControlNames = new HashSet<string>(StringComparer.Ordinal) {
        "pgndec",
        "pgnucrm",
        "pgnlcrm",
        "pgnucltr",
        "pgnlcltr",
        "pgndecd"
    };
}
