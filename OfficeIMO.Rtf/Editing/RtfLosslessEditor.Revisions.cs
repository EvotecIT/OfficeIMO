using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes the root revision save id in the root <c>rsidtbl</c> destination while preserving the rest of the RTF stream.
    /// </summary>
    public void SetRevisionRootSaveId(int? id) {
        if (id.HasValue && id.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "Revision root save id cannot be negative.");
        }

        RtfGroup root = SetRevisionRootSaveId(_syntaxTree.Root, id);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds a revision save id to the root <c>rsidtbl</c> destination, replacing duplicate entries with a single occurrence.
    /// </summary>
    public void AddRevisionSaveId(int id) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "Revision save id cannot be negative.");
        }

        RtfGroup root = AddRevisionSaveId(_syntaxTree.Root, id);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Removes all matching revision save ids from the root <c>rsidtbl</c> destination while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveRevisionSaveId(int id) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "Revision save id cannot be negative.");
        }

        RtfGroup root = RemoveRevisionSaveId(_syntaxTree.Root, id);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetRevisionRootSaveId(RtfGroup root, int? id) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "rsidtbl");

        if (tableIndex >= 0) {
            RtfGroup? updatedTable = SetRevisionRootSaveIdInTable((RtfGroup)children[tableIndex], id);
            if (updatedTable == null) {
                children.RemoveAt(tableIndex);
            } else {
                children[tableIndex] = updatedTable;
            }

            return new RtfGroup(root.Position, children);
        }

        if (!id.HasValue) {
            return root;
        }

        children.Insert(GetRevisionSaveIdTableInsertIndex(children), CreateRevisionSaveIdTable(CreateRevisionRootSaveIdControl(id.Value)));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup AddRevisionSaveId(RtfGroup root, int id) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "rsidtbl");

        if (tableIndex >= 0) {
            children[tableIndex] = AddRevisionSaveIdToTable((RtfGroup)children[tableIndex], id);
            return new RtfGroup(root.Position, children);
        }

        children.Insert(GetRevisionSaveIdTableInsertIndex(children), CreateRevisionSaveIdTable(CreateRevisionSaveIdControl(id)));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveRevisionSaveId(RtfGroup root, int id) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "rsidtbl");
        if (tableIndex < 0) {
            return root;
        }

        RtfGroup? updatedTable = RemoveRevisionSaveIdFromTable((RtfGroup)children[tableIndex], id);
        if (updatedTable == null) {
            children.RemoveAt(tableIndex);
        } else {
            children[tableIndex] = updatedTable;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup? SetRevisionRootSaveIdInTable(RtfGroup table, int? id) {
        var children = new List<RtfNode>(table.Children);
        bool replaced = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfControlWord control || control.Name != "rsidroot") {
                continue;
            }

            if (!id.HasValue || replaced) {
                children.RemoveAt(index);
            } else {
                children[index] = CreateRevisionRootSaveIdControl(id.Value);
                replaced = true;
            }
        }

        if (!replaced && id.HasValue) {
            children.Insert(GetRevisionRootInsertIndex(children), CreateRevisionRootSaveIdControl(id.Value));
        }

        return HasRevisionSaveIdTableContent(children) ? new RtfGroup(table.Position, children) : null;
    }

    private static RtfGroup AddRevisionSaveIdToTable(RtfGroup table, int id) {
        var children = new List<RtfNode>(table.Children);
        bool exists = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfControlWord control || control.Name != "rsid" || control.Parameter != id) {
                continue;
            }

            if (exists) {
                children.RemoveAt(index);
            } else {
                exists = true;
            }
        }

        if (!exists) {
            children.Add(CreateRevisionSaveIdControl(id));
        }

        return new RtfGroup(table.Position, children);
    }

    private static RtfGroup? RemoveRevisionSaveIdFromTable(RtfGroup table, int id) {
        var children = new List<RtfNode>(table.Children);

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is RtfControlWord control && control.Name == "rsid" && control.Parameter == id) {
                children.RemoveAt(index);
            }
        }

        return HasRevisionSaveIdTableContent(children) ? new RtfGroup(table.Position, children) : null;
    }

    private static bool HasRevisionSaveIdTableContent(IEnumerable<RtfNode> children) {
        return children.Any(node => node is RtfControlWord control && (control.Name == "rsidroot" || control.Name == "rsid"));
    }

    private static int GetRevisionRootInsertIndex(IReadOnlyList<RtfNode> children) {
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfControlWord control && control.Name == "rsid") {
                return index;
            }
        }

        return children.Count;
    }

    private static RtfGroup CreateRevisionSaveIdTable(RtfNode child) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "rsidtbl", null, hasParameter: false, rawText: @"\rsidtbl"),
            child
        });
    }

    private static RtfControlWord CreateRevisionRootSaveIdControl(int id) {
        return new RtfControlWord(
            0,
            "rsidroot",
            id,
            hasParameter: true,
            rawText: @"\rsidroot" + id.ToString(CultureInfo.InvariantCulture));
    }

    private static RtfControlWord CreateRevisionSaveIdControl(int id) {
        return new RtfControlWord(
            0,
            "rsid",
            id,
            hasParameter: true,
            rawText: @"\rsid" + id.ToString(CultureInfo.InvariantCulture));
    }

    private static int GetRevisionSaveIdTableInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "rsidtbl" || group.Destination == "revtbl" ||
                 group.Destination == "listoverridetable" || group.Destination == "listtable" ||
                 group.Destination == "stylesheet" || group.Destination == "colortbl" ||
                 group.Destination == "xmlnstbl" || group.Destination == "filetbl" ||
                 group.Destination == "fonttbl" || group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }
}
