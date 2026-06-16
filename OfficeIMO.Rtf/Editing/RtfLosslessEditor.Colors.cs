using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds or replaces a one-based color-table entry while preserving the rest of the RTF stream.
    /// </summary>
    public void SetColor(int index, RtfColor color) {
        if (index < 1) {
            throw new ArgumentOutOfRangeException(nameof(index), "RTF color table indexes are one-based.");
        }

        if (color == null) {
            throw new ArgumentNullException(nameof(color));
        }

        RtfGroup root = SetColor(_syntaxTree.Root, index, color);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or replaces a one-based RGB color-table entry while preserving the rest of the RTF stream.
    /// </summary>
    public void SetColor(int index, byte red, byte green, byte blue) {
        SetColor(index, new RtfColor(red, green, blue));
    }

    /// <summary>
    /// Removes the one-based color-table entry when it exists while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveColor(int index) {
        if (index < 1) {
            throw new ArgumentOutOfRangeException(nameof(index), "RTF color table indexes are one-based.");
        }

        RtfGroup root = RemoveColor(_syntaxTree.Root, index);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetColor(RtfGroup root, int index, RtfColor color) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "colortbl");

        if (tableIndex >= 0) {
            children[tableIndex] = SetColorInTable((RtfGroup)children[tableIndex], index, color);
            return new RtfGroup(root.Position, children);
        }

        if (index != 1) {
            throw new ArgumentOutOfRangeException(nameof(index), "Cannot create a sparse RTF color table.");
        }

        children.Insert(GetColorTableInsertIndex(children), CreateColorTable(color));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveColor(RtfGroup root, int index) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "colortbl");
        if (tableIndex < 0) {
            return root;
        }

        RtfGroup? updatedTable = RemoveColorFromTable((RtfGroup)children[tableIndex], index);
        if (updatedTable == null) {
            children.RemoveAt(tableIndex);
        } else {
            children[tableIndex] = updatedTable;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetColorInTable(RtfGroup colorTable, int index, RtfColor color) {
        var children = new List<RtfNode>(colorTable.Children);
        List<ColorTableEntryRange> ranges = GetColorTableEntryRanges(children);
        int rangeIndex = ranges.FindIndex(item => item.Index == index);

        if (rangeIndex >= 0) {
            ColorTableEntryRange value = ranges[rangeIndex];
            children.RemoveRange(value.StartIndex, value.Length);
            children.InsertRange(value.StartIndex, CreateColorEntryNodes(color));
            return new RtfGroup(colorTable.Position, children);
        }

        if (index != ranges.Count + 1) {
            throw new ArgumentOutOfRangeException(nameof(index), "Cannot create a sparse RTF color table.");
        }

        EnsureColorTableAutoDelimiter(children);
        children.AddRange(CreateColorEntryNodes(color));
        return new RtfGroup(colorTable.Position, children);
    }

    private static RtfGroup? RemoveColorFromTable(RtfGroup colorTable, int index) {
        var children = new List<RtfNode>(colorTable.Children);
        List<ColorTableEntryRange> ranges = GetColorTableEntryRanges(children);
        int rangeIndex = ranges.FindIndex(item => item.Index == index);
        if (rangeIndex < 0) {
            return colorTable;
        }

        ColorTableEntryRange value = ranges[rangeIndex];
        children.RemoveRange(value.StartIndex, value.Length);

        return GetColorTableEntryRanges(children).Count > 0
            ? new RtfGroup(colorTable.Position, children)
            : null;
    }

    private static List<ColorTableEntryRange> GetColorTableEntryRanges(IReadOnlyList<RtfNode> children) {
        var ranges = new List<ColorTableEntryRange>();
        int startIndex = -1;
        int colorIndex = 0;
        bool hasValue = false;

        for (int index = 0; index < children.Count; index++) {
            RtfNode child = children[index];
            if (child is RtfControlWord control && control.Name == "colortbl") {
                continue;
            }

            if (startIndex < 0) {
                startIndex = index;
            }

            hasValue |= IsColorEntryContent(child);
            if (IsColorEntryDelimiter(child)) {
                if (hasValue) {
                    ranges.Add(new ColorTableEntryRange(++colorIndex, startIndex, index - startIndex + 1));
                }

                startIndex = index + 1;
                hasValue = false;
            }
        }

        if (hasValue && startIndex >= 0) {
            ranges.Add(new ColorTableEntryRange(++colorIndex, startIndex, children.Count - startIndex));
        }

        return ranges;
    }

    private static bool IsColorEntryContent(RtfNode node) {
        return node is RtfControlWord control && IsColorEntryControl(control.Name);
    }

    private static bool IsColorEntryControl(string controlName) {
        switch (controlName) {
            case "red":
            case "green":
            case "blue":
            case "ctint":
            case "cshade":
            case "cmaindarkone":
            case "cmainlightone":
            case "cmaindarktwo":
            case "cmainlighttwo":
            case "caccentone":
            case "caccenttwo":
            case "caccentthree":
            case "caccentfour":
            case "caccentfive":
            case "caccentsix":
            case "chyperlink":
            case "cfollowedhyperlink":
            case "cbackgroundone":
            case "ctextone":
            case "cbackgroundtwo":
            case "ctexttwo":
                return true;
            default:
                return false;
        }
    }

    private static bool IsColorEntryDelimiter(RtfNode node) {
        return node is RtfText text && text.Text.Contains(";");
    }

    private static void EnsureColorTableAutoDelimiter(List<RtfNode> children) {
        foreach (RtfNode child in children) {
            if (child is RtfControlWord control && control.Name == "colortbl") {
                continue;
            }

            if (IsColorEntryDelimiter(child)) {
                return;
            }

            if (IsColorEntryContent(child)) {
                break;
            }
        }

        int insertIndex = children.FindLastIndex(node => node is RtfControlWord control && control.Name == "colortbl") + 1;
        children.Insert(insertIndex, new RtfText(0, ";", ";"));
    }

    private static RtfGroup CreateColorTable(RtfColor color) {
        var children = new List<RtfNode> {
            new RtfControlWord(0, "colortbl", null, hasParameter: false, rawText: @"\colortbl"),
            new RtfText(0, ";", ";")
        };
        children.AddRange(CreateColorEntryNodes(color));
        return new RtfGroup(0, children);
    }

    private static IReadOnlyList<RtfNode> CreateColorEntryNodes(RtfColor color) {
        var children = new List<RtfNode> {
            CreateColorNumberControl("red", color.Red),
            CreateColorNumberControl("green", color.Green),
            CreateColorNumberControl("blue", color.Blue)
        };

        string? themeControl = GetThemeColorControlName(color.ThemeColor);
        if (themeControl != null) {
            children.Add(new RtfControlWord(0, themeControl, null, hasParameter: false, rawText: "\\" + themeControl));
        }

        AddOptionalColorNumber(children, "ctint", color.Tint);
        AddOptionalColorNumber(children, "cshade", color.Shade);
        children.Add(new RtfText(0, ";", ";"));
        return children;
    }

    private static RtfControlWord CreateColorNumberControl(string controlName, int value) {
        return new RtfControlWord(
            0,
            controlName,
            value,
            hasParameter: true,
            rawText: "\\" + controlName + value.ToString(CultureInfo.InvariantCulture));
    }

    private static void AddOptionalColorNumber(List<RtfNode> children, string controlName, int? value) {
        if (!value.HasValue) {
            return;
        }

        children.Add(CreateColorNumberControl(controlName, value.Value));
    }

    private static string? GetThemeColorControlName(RtfThemeColor? themeColor) {
        if (!themeColor.HasValue) {
            return null;
        }

        switch (themeColor.Value) {
            case RtfThemeColor.MainDarkOne:
                return "cmaindarkone";
            case RtfThemeColor.MainLightOne:
                return "cmainlightone";
            case RtfThemeColor.MainDarkTwo:
                return "cmaindarktwo";
            case RtfThemeColor.MainLightTwo:
                return "cmainlighttwo";
            case RtfThemeColor.AccentOne:
                return "caccentone";
            case RtfThemeColor.AccentTwo:
                return "caccenttwo";
            case RtfThemeColor.AccentThree:
                return "caccentthree";
            case RtfThemeColor.AccentFour:
                return "caccentfour";
            case RtfThemeColor.AccentFive:
                return "caccentfive";
            case RtfThemeColor.AccentSix:
                return "caccentsix";
            case RtfThemeColor.Hyperlink:
                return "chyperlink";
            case RtfThemeColor.FollowedHyperlink:
                return "cfollowedhyperlink";
            case RtfThemeColor.BackgroundOne:
                return "cbackgroundone";
            case RtfThemeColor.TextOne:
                return "ctextone";
            case RtfThemeColor.BackgroundTwo:
                return "cbackgroundtwo";
            default:
                return "ctexttwo";
        }
    }

    private static int GetColorTableInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "colortbl" || group.Destination == "xmlnstbl" ||
                 group.Destination == "filetbl" || group.Destination == "fonttbl" ||
                 group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }

    private readonly struct ColorTableEntryRange {
        internal ColorTableEntryRange(int index, int startIndex, int length) {
            Index = index;
            StartIndex = startIndex;
            Length = length;
        }

        internal int Index { get; }

        internal int StartIndex { get; }

        internal int Length { get; }
    }
}
