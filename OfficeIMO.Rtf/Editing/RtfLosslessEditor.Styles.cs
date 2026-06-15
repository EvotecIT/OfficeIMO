using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds a basic stylesheet entry or renames an existing entry while preserving the existing style formatting controls.
    /// </summary>
    public void SetStyleName(int id, string name, RtfStyleKind kind = RtfStyleKind.Paragraph) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "RTF style identifiers cannot be negative.");
        }

        if (name == null) {
            throw new ArgumentNullException(nameof(name));
        }

        RtfGroup root = SetStyleName(_syntaxTree.Root, id, name, kind);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Removes all stylesheet entries matching the supplied id and kind while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveStyle(int id, RtfStyleKind kind = RtfStyleKind.Paragraph) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "RTF style identifiers cannot be negative.");
        }

        RtfGroup root = RemoveStyle(_syntaxTree.Root, id, kind);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetStyleName(RtfGroup root, int id, string name, RtfStyleKind kind) {
        var children = new List<RtfNode>(root.Children);
        int stylesheetIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "stylesheet");

        if (stylesheetIndex >= 0) {
            children[stylesheetIndex] = SetStyleNameInStylesheet((RtfGroup)children[stylesheetIndex], id, name, kind);
            return new RtfGroup(root.Position, children);
        }

        children.Insert(GetStylesheetInsertIndex(children), CreateStylesheet(id, name, kind));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveStyle(RtfGroup root, int id, RtfStyleKind kind) {
        var children = new List<RtfNode>(root.Children);
        int stylesheetIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "stylesheet");
        if (stylesheetIndex < 0) {
            return root;
        }

        RtfGroup? updatedStylesheet = RemoveStyleFromStylesheet((RtfGroup)children[stylesheetIndex], id, kind);
        if (updatedStylesheet == null) {
            children.RemoveAt(stylesheetIndex);
        } else {
            children[stylesheetIndex] = updatedStylesheet;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetStyleNameInStylesheet(RtfGroup stylesheet, int id, string name, RtfStyleKind kind) {
        var children = new List<RtfNode>(stylesheet.Children);
        var matchingIndexes = new List<int>();

        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup styleGroup && StyleMatches(styleGroup, id, kind)) {
                matchingIndexes.Add(index);
            }
        }

        if (matchingIndexes.Count > 0) {
            children[matchingIndexes[0]] = RenameStyleGroup((RtfGroup)children[matchingIndexes[0]], name);
            for (int index = matchingIndexes.Count - 1; index > 0; index--) {
                children.RemoveAt(matchingIndexes[index]);
            }

            return new RtfGroup(stylesheet.Position, children);
        }

        children.Add(CreateStyleGroup(id, name, kind));
        return new RtfGroup(stylesheet.Position, children);
    }

    private static RtfGroup? RemoveStyleFromStylesheet(RtfGroup stylesheet, int id, RtfStyleKind kind) {
        var children = new List<RtfNode>(stylesheet.Children);

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is RtfGroup styleGroup && StyleMatches(styleGroup, id, kind)) {
                children.RemoveAt(index);
            }
        }

        return children.Any(node => node is RtfGroup group && TryGetStyleIdentity(group, out _, out _))
            ? new RtfGroup(stylesheet.Position, children)
            : null;
    }

    private static bool StyleMatches(RtfGroup styleGroup, int id, RtfStyleKind kind) {
        return TryGetStyleIdentity(styleGroup, out int foundId, out RtfStyleKind foundKind) &&
               foundId == id &&
               foundKind == kind;
    }

    private static bool TryGetStyleIdentity(RtfGroup styleGroup, out int id, out RtfStyleKind kind) {
        foreach (RtfControlWord control in styleGroup.Children.OfType<RtfControlWord>()) {
            switch (control.Name) {
                case "s":
                    id = control.Parameter ?? -1;
                    kind = RtfStyleKind.Paragraph;
                    return id >= 0;
                case "cs":
                    id = control.Parameter ?? -1;
                    kind = RtfStyleKind.Character;
                    return id >= 0;
                case "ts":
                    id = control.Parameter ?? -1;
                    kind = RtfStyleKind.Table;
                    return id >= 0;
            }
        }

        id = -1;
        kind = RtfStyleKind.Paragraph;
        return false;
    }

    private static RtfGroup RenameStyleGroup(RtfGroup styleGroup, string name) {
        var children = new List<RtfNode>(styleGroup.Children);
        int nameIndex = FindStyleNameTextIndex(children);
        if (nameIndex >= 0) {
            children[nameIndex] = CreateStyleNameText(name, StyleNameHasLeadingSeparator((RtfText)children[nameIndex]));
        } else {
            children.Add(CreateStyleNameText(name));
        }

        return new RtfGroup(styleGroup.Position, children);
    }

    private static int FindStyleNameTextIndex(IReadOnlyList<RtfNode> children) {
        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is RtfText text && text.Text.Contains(";")) {
                return index;
            }
        }

        return -1;
    }

    private static bool StyleNameHasLeadingSeparator(RtfText text) {
        return text.RawText.Length > 0 && char.IsWhiteSpace(text.RawText[0]);
    }

    private static RtfGroup CreateStylesheet(int id, string name, RtfStyleKind kind) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, "stylesheet", null, hasParameter: false, rawText: @"\stylesheet"),
            CreateStyleGroup(id, name, kind)
        });
    }

    private static RtfGroup CreateStyleGroup(int id, string name, RtfStyleKind kind) {
        var children = new List<RtfNode>();
        if (kind == RtfStyleKind.Character || kind == RtfStyleKind.Table) {
            children.Add(new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"));
        }

        string controlName = kind == RtfStyleKind.Character ? "cs" : kind == RtfStyleKind.Table ? "ts" : "s";
        children.Add(new RtfControlWord(
            0,
            controlName,
            id,
            hasParameter: true,
            rawText: "\\" + controlName + id.ToString(CultureInfo.InvariantCulture)));
        children.Add(CreateStyleNameText(name));
        return new RtfGroup(0, children);
    }

    private static RtfText CreateStyleNameText(string name, bool includeLeadingSeparator = true) {
        string separator = includeLeadingSeparator ? " " : string.Empty;
        string text = separator + name + ";";
        return new RtfText(0, text, separator + RtfTextEncoding.EncodeText(name) + ";");
    }

    private static int GetStylesheetInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "stylesheet" || group.Destination == "colortbl" ||
                 group.Destination == "xmlnstbl" || group.Destination == "filetbl" ||
                 group.Destination == "fonttbl" || group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }
}
