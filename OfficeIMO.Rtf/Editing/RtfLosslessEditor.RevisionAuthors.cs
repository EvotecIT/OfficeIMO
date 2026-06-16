using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds or replaces a zero-based revision author table entry while preserving the rest of the RTF stream.
    /// </summary>
    public void SetRevisionAuthor(int index, string name) {
        if (index < 0) {
            throw new ArgumentOutOfRangeException(nameof(index), "Revision author indexes cannot be negative.");
        }

        if (name == null) {
            throw new ArgumentNullException(nameof(name));
        }

        RtfGroup root = SetRevisionAuthor(_syntaxTree.Root, index, name);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Removes the zero-based revision author table entry when it exists while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveRevisionAuthor(int index) {
        if (index < 0) {
            throw new ArgumentOutOfRangeException(nameof(index), "Revision author indexes cannot be negative.");
        }

        RtfGroup root = RemoveRevisionAuthor(_syntaxTree.Root, index);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetRevisionAuthor(RtfGroup root, int index, string name) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "revtbl");

        if (tableIndex >= 0) {
            children[tableIndex] = SetRevisionAuthorInTable((RtfGroup)children[tableIndex], index, name);
            return new RtfGroup(root.Position, children);
        }

        if (index != 0) {
            throw new ArgumentOutOfRangeException(nameof(index), "Cannot create a sparse RTF revision author table.");
        }

        children.Insert(GetRevisionAuthorTableInsertIndex(children), CreateRevisionAuthorTable(name));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveRevisionAuthor(RtfGroup root, int index) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "revtbl");
        if (tableIndex < 0) {
            return root;
        }

        RtfGroup? updatedTable = RemoveRevisionAuthorFromTable((RtfGroup)children[tableIndex], index);
        if (updatedTable == null) {
            children.RemoveAt(tableIndex);
        } else {
            children[tableIndex] = updatedTable;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetRevisionAuthorInTable(RtfGroup revisionTable, int index, string name) {
        var children = new List<RtfNode>(revisionTable.Children);
        List<int> authorIndexes = GetRevisionAuthorChildIndexes(children);

        if (index < authorIndexes.Count) {
            children[authorIndexes[index]] = CreateRevisionAuthorGroup(name);
            return new RtfGroup(revisionTable.Position, children);
        }

        if (index != authorIndexes.Count) {
            throw new ArgumentOutOfRangeException(nameof(index), "Cannot create a sparse RTF revision author table.");
        }

        children.Add(CreateRevisionAuthorGroup(name));
        return new RtfGroup(revisionTable.Position, children);
    }

    private static RtfGroup? RemoveRevisionAuthorFromTable(RtfGroup revisionTable, int index) {
        var children = new List<RtfNode>(revisionTable.Children);
        List<int> authorIndexes = GetRevisionAuthorChildIndexes(children);
        if (index >= authorIndexes.Count) {
            return revisionTable;
        }

        children.RemoveAt(authorIndexes[index]);

        return GetRevisionAuthorChildIndexes(children).Count > 0
            ? new RtfGroup(revisionTable.Position, children)
            : null;
    }

    private static List<int> GetRevisionAuthorChildIndexes(IReadOnlyList<RtfNode> children) {
        var indexes = new List<int>();
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup) {
                indexes.Add(index);
            }
        }

        return indexes;
    }

    private static RtfGroup CreateRevisionAuthorTable(string name) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "revtbl", null, hasParameter: false, rawText: @"\revtbl"),
            CreateRevisionAuthorGroup(name)
        });
    }

    private static RtfGroup CreateRevisionAuthorGroup(string name) {
        string value = name + ";";
        return new RtfGroup(0, new RtfNode[] {
            new RtfText(0, value, RtfTextEncoding.EncodeText(value))
        });
    }

    private static int GetRevisionAuthorTableInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "revtbl" || group.Destination == "listoverridetable" ||
                 group.Destination == "listtable" || group.Destination == "stylesheet" ||
                 group.Destination == "colortbl" || group.Destination == "xmlnstbl" ||
                 group.Destination == "filetbl" || group.Destination == "fonttbl" ||
                 group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }
}
