using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes root page-border option controls while preserving untouched RTF syntax.
    /// </summary>
    public void SetPageBorderOptions(
        bool? includeHeader = null,
        bool? includeFooter = null,
        bool? snapToPageBorder = null,
        RtfPageBorderScope? scope = null,
        bool? displayBehindText = null,
        RtfPageBorderOffset? offsetFrom = null) {
        ValidatePageBorderScope(scope);
        ValidatePageBorderOffset(offsetFrom);

        RtfGroup root = SetRootOptionalParameterlessControl(_syntaxTree.Root, includeHeader == true ? "pgbrdrhead" : null, SingleControlName("pgbrdrhead"));
        root = SetRootOptionalParameterlessControl(root, includeFooter == true ? "pgbrdrfoot" : null, SingleControlName("pgbrdrfoot"));
        root = SetRootControl(root, "pgbrdropt", GetPageBorderDisplayOptionsValue(scope, displayBehindText, offsetFrom), SingleControlName("pgbrdropt"));
        root = SetRootOptionalParameterlessControl(root, snapToPageBorder == true ? "pgbrdrsnap" : null, SingleControlName("pgbrdrsnap"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes one root page-border side while preserving untouched RTF syntax.
    /// </summary>
    public void SetPageBorder(
        RtfPageBorderSide side,
        RtfPageBorderStyle style,
        int? width = null,
        int? space = null,
        int? colorIndex = null,
        bool frame = false) {
        ValidatePageBorderSide(side);
        ValidatePageBorderStyle(style);
        ValidateNonNegative(width, nameof(width));
        ValidateNonNegative(space, nameof(space));
        ValidateNonNegative(colorIndex, nameof(colorIndex));

        RtfGroup root = SetRootPageBorderSide(_syntaxTree.Root, side, style, width, space, colorIndex, frame);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Removes one root page-border side while preserving untouched RTF syntax.
    /// </summary>
    public void RemovePageBorder(RtfPageBorderSide side) {
        SetPageBorder(side, RtfPageBorderStyle.None);
    }

    private static RtfGroup SetRootPageBorderSide(
        RtfGroup root,
        RtfPageBorderSide side,
        RtfPageBorderStyle style,
        int? width,
        int? space,
        int? colorIndex,
        bool frame) {
        string sideControlName = GetPageBorderSideControlName(side);
        var children = new List<RtfNode>(root.Children);
        RemoveRootPageBorderSideRanges(children, sideControlName);

        if (style == RtfPageBorderStyle.None && !width.HasValue && !space.HasValue && !colorIndex.HasValue && !frame) {
            return new RtfGroup(root.Position, children);
        }

        List<RtfNode> borderNodes = CreatePageBorderSideNodes(sideControlName, style, width, space, colorIndex, frame);
        children.InsertRange(GetRootControlInsertIndex(children, sideControlName), borderNodes);
        return new RtfGroup(root.Position, children);
    }

    private static void RemoveRootPageBorderSideRanges(List<RtfNode> children, string sideControlName) {
        int bodyStartIndex = GetRootBodyStartIndex(children);
        for (int index = bodyStartIndex - 1; index >= 0; index--) {
            if (children[index] is not RtfControlWord control || control.Name != sideControlName) {
                continue;
            }

            int endIndex = index + 1;
            while (endIndex < bodyStartIndex &&
                   children[endIndex] is RtfControlWord nextControl &&
                   PageBorderPropertyControlNames.Contains(nextControl.Name)) {
                endIndex++;
            }

            children.RemoveRange(index, endIndex - index);
        }
    }

    private static List<RtfNode> CreatePageBorderSideNodes(
        string sideControlName,
        RtfPageBorderStyle style,
        int? width,
        int? space,
        int? colorIndex,
        bool frame) {
        var nodes = new List<RtfNode> {
            CreateRootControl(sideControlName, null),
            CreateRootControl(GetPageBorderStyleControlName(style), null)
        };

        if (width.HasValue) {
            nodes.Add(CreateRootControl("brdrw", width));
        }

        if (space.HasValue) {
            nodes.Add(CreateRootControl("brsp", space));
        }

        if (colorIndex.HasValue) {
            nodes.Add(CreateRootControl("brdrcf", colorIndex));
        }

        if (frame) {
            nodes.Add(CreateRootControl("brdrframe", null));
        }

        return nodes;
    }

    private static int? GetPageBorderDisplayOptionsValue(
        RtfPageBorderScope? scope,
        bool? displayBehindText,
        RtfPageBorderOffset? offsetFrom) {
        if (!scope.HasValue && !displayBehindText.HasValue && !offsetFrom.HasValue) {
            return null;
        }

        int value = scope switch {
            RtfPageBorderScope.FirstPageInSection => 1,
            RtfPageBorderScope.AllExceptFirstPageInSection => 2,
            RtfPageBorderScope.WholeDocument => 3,
            _ => 0
        };

        if (displayBehindText == true) {
            value |= 8;
        }

        if (offsetFrom == RtfPageBorderOffset.PageEdge) {
            value |= 32;
        }

        return value;
    }

    private static string GetPageBorderSideControlName(RtfPageBorderSide side) {
        switch (side) {
            case RtfPageBorderSide.Bottom:
                return "pgbrdrb";
            case RtfPageBorderSide.Left:
                return "pgbrdrl";
            case RtfPageBorderSide.Right:
                return "pgbrdrr";
            case RtfPageBorderSide.Top:
                return "pgbrdrt";
            default:
                throw new ArgumentOutOfRangeException(nameof(side), "Unsupported RTF page-border side.");
        }
    }

    private static string GetPageBorderStyleControlName(RtfPageBorderStyle style) {
        switch (style) {
            case RtfPageBorderStyle.Single:
                return "brdrs";
            case RtfPageBorderStyle.Double:
                return "brdrdb";
            case RtfPageBorderStyle.Dotted:
                return "brdrdot";
            case RtfPageBorderStyle.Dashed:
                return "brdrdash";
            case RtfPageBorderStyle.Shadow:
                return "brdrsh";
            case RtfPageBorderStyle.None:
                return "brdrnil";
            default:
                throw new ArgumentOutOfRangeException(nameof(style), "Unsupported RTF page-border style.");
        }
    }

    private static void ValidatePageBorderSide(RtfPageBorderSide side) {
        GetPageBorderSideControlName(side);
    }

    private static void ValidatePageBorderStyle(RtfPageBorderStyle style) {
        GetPageBorderStyleControlName(style);
    }

    private static void ValidatePageBorderScope(RtfPageBorderScope? scope) {
        if (!scope.HasValue) {
            return;
        }

        switch (scope.Value) {
            case RtfPageBorderScope.AllPagesInSection:
            case RtfPageBorderScope.FirstPageInSection:
            case RtfPageBorderScope.AllExceptFirstPageInSection:
            case RtfPageBorderScope.WholeDocument:
                return;
            default:
                throw new ArgumentOutOfRangeException(nameof(scope), "Unsupported RTF page-border scope.");
        }
    }

    private static void ValidatePageBorderOffset(RtfPageBorderOffset? offset) {
        if (!offset.HasValue) {
            return;
        }

        switch (offset.Value) {
            case RtfPageBorderOffset.Text:
            case RtfPageBorderOffset.PageEdge:
                return;
            default:
                throw new ArgumentOutOfRangeException(nameof(offset), "Unsupported RTF page-border offset.");
        }
    }

    private static readonly ISet<string> PageBorderPropertyControlNames = new HashSet<string>(StringComparer.Ordinal) {
        "brdrs",
        "brdrdb",
        "brdrdot",
        "brdrdash",
        "brdrsh",
        "brdrnone",
        "brdrnil",
        "brdrw",
        "brsp",
        "brdrcf",
        "brdrframe"
    };
}
