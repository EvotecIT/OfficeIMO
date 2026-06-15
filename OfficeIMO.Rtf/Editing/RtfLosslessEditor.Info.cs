using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes the root generator metadata group while preserving the rest of the RTF stream.
    /// </summary>
    public void SetGenerator(string? value) {
        RtfGroup root = SetGenerator(_syntaxTree.Root, value);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes a document information field while preserving the rest of the RTF stream.
    /// </summary>
    public void SetInfo(RtfDocumentInfoField field, string? value) {
        string destination = GetInfoDestination(field);
        RtfGroup root = SetInfo(_syntaxTree.Root, destination, value);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes a document information timestamp while preserving the rest of the RTF stream.
    /// </summary>
    public void SetInfoTimestamp(RtfDocumentInfoTimestampField field, DateTime? value) {
        string destination = GetInfoTimestampDestination(field);
        RtfGroup root = SetInfoTimestamp(_syntaxTree.Root, destination, value);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes a numeric document information field while preserving the rest of the RTF stream.
    /// </summary>
    public void SetInfoNumber(RtfDocumentInfoNumberField field, int? value) {
        string controlName = GetInfoNumberControlName(field);
        RtfGroup root = SetInfoNumber(_syntaxTree.Root, controlName, value);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetGenerator(RtfGroup root, string? value) {
        var children = new List<RtfNode>(root.Children);
        bool replaced = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfGroup group || group.Destination != "generator") {
                continue;
            }

            if (string.IsNullOrEmpty(value) || replaced) {
                children.RemoveAt(index);
            } else {
                children[index] = CreateGeneratorGroup(value!);
                replaced = true;
            }
        }

        if (!replaced && !string.IsNullOrEmpty(value)) {
            children.Insert(GetInfoInsertIndex(children), CreateGeneratorGroup(value!));
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetInfo(RtfGroup root, string destination, string? value) {
        var children = new List<RtfNode>(root.Children);
        int infoIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "info");

        if (infoIndex >= 0) {
            RtfGroup infoGroup = (RtfGroup)children[infoIndex];
            RtfGroup? updatedInfo = SetInfoChild(infoGroup, destination, value);
            if (updatedInfo == null) {
                children.RemoveAt(infoIndex);
            } else {
                children[infoIndex] = updatedInfo;
            }

            return new RtfGroup(root.Position, children);
        }

        if (string.IsNullOrEmpty(value)) {
            return root;
        }

        children.Insert(GetInfoInsertIndex(children), CreateInfoGroup(destination, value!));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetInfoTimestamp(RtfGroup root, string destination, DateTime? value) {
        var children = new List<RtfNode>(root.Children);
        int infoIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "info");

        if (infoIndex >= 0) {
            RtfGroup infoGroup = (RtfGroup)children[infoIndex];
            RtfGroup? updatedInfo = SetInfoTimestampChild(infoGroup, destination, value);
            if (updatedInfo == null) {
                children.RemoveAt(infoIndex);
            } else {
                children[infoIndex] = updatedInfo;
            }

            return new RtfGroup(root.Position, children);
        }

        if (!value.HasValue) {
            return root;
        }

        children.Insert(GetInfoInsertIndex(children), CreateInfoGroup(CreateInfoTimestampGroup(destination, value.Value)));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetInfoNumber(RtfGroup root, string controlName, int? value) {
        var children = new List<RtfNode>(root.Children);
        int infoIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "info");

        if (infoIndex >= 0) {
            RtfGroup infoGroup = (RtfGroup)children[infoIndex];
            RtfGroup? updatedInfo = SetInfoNumberChild(infoGroup, controlName, value);
            if (updatedInfo == null) {
                children.RemoveAt(infoIndex);
            } else {
                children[infoIndex] = updatedInfo;
            }

            return new RtfGroup(root.Position, children);
        }

        if (!value.HasValue) {
            return root;
        }

        children.Insert(GetInfoInsertIndex(children), CreateInfoGroup(CreateInfoNumberControl(controlName, value.Value)));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup? SetInfoChild(RtfGroup infoGroup, string destination, string? value) {
        var children = new List<RtfNode>(infoGroup.Children);
        int childIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == destination);

        if (childIndex >= 0) {
            if (string.IsNullOrEmpty(value)) {
                children.RemoveAt(childIndex);
            } else {
                children[childIndex] = CreateInfoFieldGroup(destination, value!);
            }
        } else if (!string.IsNullOrEmpty(value)) {
            children.Add(CreateInfoFieldGroup(destination, value!));
        }

        return HasInfoContent(children) ? new RtfGroup(infoGroup.Position, children) : null;
    }

    private static RtfGroup? SetInfoTimestampChild(RtfGroup infoGroup, string destination, DateTime? value) {
        var children = new List<RtfNode>(infoGroup.Children);
        int childIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == destination);

        if (childIndex >= 0) {
            if (!value.HasValue) {
                children.RemoveAt(childIndex);
            } else {
                children[childIndex] = CreateInfoTimestampGroup(destination, value.Value);
            }
        } else if (value.HasValue) {
            children.Add(CreateInfoTimestampGroup(destination, value.Value));
        }

        return HasInfoContent(children) ? new RtfGroup(infoGroup.Position, children) : null;
    }

    private static RtfGroup? SetInfoNumberChild(RtfGroup infoGroup, string controlName, int? value) {
        var children = new List<RtfNode>(infoGroup.Children);
        int childIndex = children.FindIndex(node => node is RtfControlWord control && control.Name == controlName);

        if (childIndex >= 0) {
            if (!value.HasValue) {
                children.RemoveAt(childIndex);
            } else {
                children[childIndex] = CreateInfoNumberControl(controlName, value.Value);
            }
        } else if (value.HasValue) {
            children.Add(CreateInfoNumberControl(controlName, value.Value));
        }

        return HasInfoContent(children) ? new RtfGroup(infoGroup.Position, children) : null;
    }

    private static RtfGroup CreateGeneratorGroup(string value) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "generator", null, hasParameter: false, rawText: @"\generator "),
            new RtfText(0, value + ";", RtfTextEncoding.EncodeText(value + ";"))
        });
    }

    private static RtfGroup CreateInfoGroup(string destination, string value) {
        return CreateInfoGroup(CreateInfoFieldGroup(destination, value));
    }

    private static RtfGroup CreateInfoGroup(RtfNode child) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, "info", null, hasParameter: false, rawText: @"\info"),
            child
        });
    }

    private static RtfGroup CreateInfoFieldGroup(string destination, string value) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, destination, null, hasParameter: false, rawText: "\\" + destination + " "),
            new RtfText(0, value, RtfTextEncoding.EncodeText(value))
        });
    }

    private static RtfGroup CreateInfoTimestampGroup(string destination, DateTime value) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, destination, null, hasParameter: false, rawText: "\\" + destination),
            CreateInfoNumberControl("yr", value.Year),
            CreateInfoNumberControl("mo", value.Month),
            CreateInfoNumberControl("dy", value.Day),
            CreateInfoNumberControl("hr", value.Hour),
            CreateInfoNumberControl("min", value.Minute),
            CreateInfoNumberControl("sec", value.Second)
        });
    }

    private static RtfControlWord CreateInfoNumberControl(string controlName, int value) {
        return new RtfControlWord(
            0,
            controlName,
            value,
            hasParameter: true,
            rawText: "\\" + controlName + value.ToString(CultureInfo.InvariantCulture));
    }

    private static bool HasInfoContent(IEnumerable<RtfNode> children) {
        return children.Any(node => node is RtfGroup || node is RtfControlWord control && control.Name != "info");
    }

    private static string GetInfoDestination(RtfDocumentInfoField field) {
        return field switch {
            RtfDocumentInfoField.Title => "title",
            RtfDocumentInfoField.Subject => "subject",
            RtfDocumentInfoField.Author => "author",
            RtfDocumentInfoField.Manager => "manager",
            RtfDocumentInfoField.Company => "company",
            RtfDocumentInfoField.Operator => "operator",
            RtfDocumentInfoField.Category => "category",
            RtfDocumentInfoField.Keywords => "keywords",
            RtfDocumentInfoField.Comments => "comment",
            RtfDocumentInfoField.HyperlinkBase => "hlinkbase",
            _ => throw new ArgumentOutOfRangeException(nameof(field), field, "Unsupported RTF document information field.")
        };
    }

    private static string GetInfoTimestampDestination(RtfDocumentInfoTimestampField field) {
        return field switch {
            RtfDocumentInfoTimestampField.Created => "creatim",
            RtfDocumentInfoTimestampField.Revised => "revtim",
            RtfDocumentInfoTimestampField.Printed => "printim",
            RtfDocumentInfoTimestampField.BackedUp => "buptim",
            _ => throw new ArgumentOutOfRangeException(nameof(field), field, "Unsupported RTF document information timestamp field.")
        };
    }

    private static string GetInfoNumberControlName(RtfDocumentInfoNumberField field) {
        return field switch {
            RtfDocumentInfoNumberField.EditingMinutes => "edmins",
            RtfDocumentInfoNumberField.NumberOfPages => "nofpages",
            RtfDocumentInfoNumberField.NumberOfWords => "nofwords",
            RtfDocumentInfoNumberField.NumberOfCharacters => "nofchars",
            RtfDocumentInfoNumberField.NumberOfCharactersWithSpaces => "nofcharsws",
            RtfDocumentInfoNumberField.InternalVersion => "vern",
            _ => throw new ArgumentOutOfRangeException(nameof(field), field, "Unsupported RTF document information number field.")
        };
    }
}
