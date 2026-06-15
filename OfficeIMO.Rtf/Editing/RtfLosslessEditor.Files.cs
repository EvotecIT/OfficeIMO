using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds or replaces a file-table reference in the root <c>filetbl</c> destination while preserving the rest of the RTF stream.
    /// </summary>
    public void SetFileReference(RtfFileReference fileReference) {
        if (fileReference == null) {
            throw new ArgumentNullException(nameof(fileReference));
        }

        if (fileReference.Id < 0) {
            throw new ArgumentOutOfRangeException(nameof(fileReference), "File reference id cannot be negative.");
        }

        RtfGroup root = SetFileReference(_syntaxTree.Root, fileReference);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or replaces a file-table reference with the supplied identifier and path.
    /// </summary>
    public void SetFileReference(int id, string path) {
        SetFileReference(new RtfFileReference(id, path));
    }

    /// <summary>
    /// Removes all file-table references with the supplied identifier while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveFileReference(int id) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "File reference id cannot be negative.");
        }

        RtfGroup root = RemoveFileReference(_syntaxTree.Root, id);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetFileReference(RtfGroup root, RtfFileReference fileReference) {
        var children = new List<RtfNode>(root.Children);
        int fileTableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "filetbl");

        if (fileTableIndex >= 0) {
            children[fileTableIndex] = SetFileReferenceInTable((RtfGroup)children[fileTableIndex], fileReference);
            return new RtfGroup(root.Position, children);
        }

        children.Insert(GetFileTableInsertIndex(children), CreateFileTable(fileReference));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveFileReference(RtfGroup root, int id) {
        var children = new List<RtfNode>(root.Children);
        int fileTableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "filetbl");
        if (fileTableIndex < 0) {
            return root;
        }

        RtfGroup? updatedTable = RemoveFileReferenceFromTable((RtfGroup)children[fileTableIndex], id);
        if (updatedTable == null) {
            children.RemoveAt(fileTableIndex);
        } else {
            children[fileTableIndex] = updatedTable;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetFileReferenceInTable(RtfGroup fileTable, RtfFileReference fileReference) {
        var children = new List<RtfNode>(fileTable.Children);
        bool replaced = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfGroup fileGroup || !FileReferenceIdMatches(fileGroup, fileReference.Id)) {
                continue;
            }

            if (replaced) {
                children.RemoveAt(index);
            } else {
                children[index] = CreateFileReferenceGroup(fileReference);
                replaced = true;
            }
        }

        if (!replaced) {
            children.Add(CreateFileReferenceGroup(fileReference));
        }

        return new RtfGroup(fileTable.Position, children);
    }

    private static RtfGroup? RemoveFileReferenceFromTable(RtfGroup fileTable, int id) {
        var children = new List<RtfNode>(fileTable.Children);

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is RtfGroup fileGroup && FileReferenceIdMatches(fileGroup, id)) {
                children.RemoveAt(index);
            }
        }

        return children.Any(node => node is RtfGroup group && group.Destination == "file")
            ? new RtfGroup(fileTable.Position, children)
            : null;
    }

    private static bool FileReferenceIdMatches(RtfGroup fileGroup, int id) {
        return fileGroup.Destination == "file" &&
               fileGroup.Children.OfType<RtfControlWord>().Any(control => control.Name == "fid" && control.Parameter == id);
    }

    private static RtfGroup CreateFileTable(RtfFileReference fileReference) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "filetbl", null, hasParameter: false, rawText: @"\filetbl"),
            CreateFileReferenceGroup(fileReference)
        });
    }

    private static RtfGroup CreateFileReferenceGroup(RtfFileReference fileReference) {
        var children = new List<RtfNode> {
            new RtfControlWord(0, "file", null, hasParameter: false, rawText: @"\file"),
            new RtfControlWord(
                0,
                "fid",
                fileReference.Id,
                hasParameter: true,
                rawText: @"\fid" + fileReference.Id.ToString(CultureInfo.InvariantCulture))
        };

        AddOptionalFileNumber(children, "frelative", fileReference.RelativePathStart);
        AddOptionalFileNumber(children, "fosnum", fileReference.OperatingSystemNumber);
        AddFileSourceControls(children, fileReference.Sources);
        children.Add(new RtfText(0, " " + fileReference.Path, " " + RtfTextEncoding.EncodeText(fileReference.Path)));
        return new RtfGroup(0, children);
    }

    private static void AddOptionalFileNumber(List<RtfNode> children, string controlName, int? value) {
        if (!value.HasValue) {
            return;
        }

        children.Add(new RtfControlWord(
            0,
            controlName,
            value.Value,
            hasParameter: true,
            rawText: "\\" + controlName + value.Value.ToString(CultureInfo.InvariantCulture)));
    }

    private static void AddFileSourceControls(List<RtfNode> children, RtfFileSource sources) {
        AddFileSourceControl(children, sources, RtfFileSource.Mac, "fvalidmac");
        AddFileSourceControl(children, sources, RtfFileSource.Dos, "fvaliddos");
        AddFileSourceControl(children, sources, RtfFileSource.Ntfs, "fvalidntfs");
        AddFileSourceControl(children, sources, RtfFileSource.Hpfs, "fvalidhpfs");
        AddFileSourceControl(children, sources, RtfFileSource.Network, "fnetwork");
    }

    private static void AddFileSourceControl(List<RtfNode> children, RtfFileSource sources, RtfFileSource flag, string controlName) {
        if ((sources & flag) != flag) {
            return;
        }

        children.Add(new RtfControlWord(0, controlName, null, hasParameter: false, rawText: "\\" + controlName));
    }

    private static int GetFileTableInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "filetbl" || group.Destination == "fonttbl" || group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }
}
