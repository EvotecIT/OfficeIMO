using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <summary>
/// Applies targeted edits to an RTF syntax tree while preserving untouched raw RTF.
/// </summary>
public sealed class RtfLosslessEditor {
    private RtfSyntaxTree _syntaxTree;

    /// <summary>
    /// Creates a lossless editor from a read result.
    /// </summary>
    public RtfLosslessEditor(RtfReadResult readResult)
        : this((readResult ?? throw new ArgumentNullException(nameof(readResult))).SyntaxTree) {
    }

    /// <summary>
    /// Creates a lossless editor from an RTF syntax tree.
    /// </summary>
    public RtfLosslessEditor(RtfSyntaxTree syntaxTree) {
        _syntaxTree = syntaxTree ?? throw new ArgumentNullException(nameof(syntaxTree));
    }

    /// <summary>Current edited syntax tree.</summary>
    public RtfSyntaxTree SyntaxTree => _syntaxTree;

    /// <summary>
    /// Replaces literal visible text contained inside individual text nodes.
    /// </summary>
    public int ReplaceText(string oldText, string newText, StringComparison comparison = StringComparison.Ordinal) {
        if (oldText == null) throw new ArgumentNullException(nameof(oldText));
        if (newText == null) throw new ArgumentNullException(nameof(newText));
        if (oldText.Length == 0) throw new ArgumentException("Text to replace cannot be empty.", nameof(oldText));

        int replacements = 0;
        RtfGroup root = RewriteGroup(_syntaxTree.Root, skipTextReplacement: false, node => {
            if (node is not RtfText text || text.Text.IndexOf(oldText, comparison) < 0) {
                return node;
            }

            string replaced = Replace(text.Text, oldText, newText, comparison, ref replacements);
            return new RtfText(text.Position, replaced, RtfTextEncoding.EncodeText(replaced));
        });

        if (replacements > 0) {
            _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
        }

        return replacements;
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
    /// Appends a plain RTF paragraph at the end of the root document group while preserving existing syntax.
    /// </summary>
    public void AppendParagraph(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));

        var children = new List<RtfNode>(_syntaxTree.Root.Children) {
            new RtfControlWord(0, "pard", null, hasParameter: false, rawText: @"\pard "),
            new RtfText(0, text, RtfTextEncoding.EncodeText(text)),
            new RtfControlWord(0, "par", null, hasParameter: false, rawText: @"\par")
        };

        _syntaxTree = new RtfSyntaxTree(new RtfGroup(_syntaxTree.Root.Position, children), _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Serializes the edited syntax tree without semantic normalization.
    /// </summary>
    public string ToRtf() => _syntaxTree.ToRtf();

    /// <summary>
    /// Reads the edited syntax tree into a fresh semantic model.
    /// </summary>
    public RtfReadResult ToReadResult(RtfReadOptions? options = null) => RtfDocument.Read(ToRtf(), options);

    private static RtfGroup RewriteGroup(RtfGroup group, bool skipTextReplacement, Func<RtfNode, RtfNode> rewriteNode) {
        bool childSkip = skipTextReplacement || ShouldSkipTextReplacement(group);
        var children = new List<RtfNode>(group.Children.Count);
        bool changed = false;

        foreach (RtfNode child in group.Children) {
            RtfNode rewritten = child;
            if (child is RtfGroup childGroup) {
                rewritten = RewriteGroup(childGroup, childSkip, rewriteNode);
            } else if (!childSkip) {
                rewritten = rewriteNode(child);
            }

            changed |= !ReferenceEquals(child, rewritten);
            children.Add(rewritten);
        }

        return changed ? new RtfGroup(group.Position, children) : group;
    }

    private static bool ShouldSkipTextReplacement(RtfGroup group) {
        if (RtfDestinationRegistry.IsIgnorableDestinationGroup(group)) {
            return true;
        }

        string? destination = group.Destination;
        return RtfDestinationRegistry.ShouldSkipTextReplacement(destination);
    }

    private static string Replace(string text, string oldText, string newText, StringComparison comparison, ref int replacements) {
        var builder = new StringBuilder(text.Length);
        int current = 0;
        while (current < text.Length) {
            int match = text.IndexOf(oldText, current, comparison);
            if (match < 0) {
                builder.Append(text, current, text.Length - current);
                break;
            }

            builder.Append(text, current, match - current);
            builder.Append(newText);
            current = match + oldText.Length;
            replacements++;
        }

        return builder.ToString();
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

        return children.Count == 0 ? null : new RtfGroup(infoGroup.Position, children);
    }

    private static RtfGroup CreateInfoGroup(string destination, string value) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, "info", null, hasParameter: false, rawText: @"\info"),
            CreateInfoFieldGroup(destination, value)
        });
    }

    private static RtfGroup CreateInfoFieldGroup(string destination, string value) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, destination, null, hasParameter: false, rawText: "\\" + destination + " "),
            new RtfText(0, value, RtfTextEncoding.EncodeText(value))
        });
    }

    private static int GetInfoInsertIndex(List<RtfNode> children) {
        int index = 0;
        while (index < children.Count &&
               children[index] is RtfControlWord control &&
               RtfDestinationRegistry.IsHeaderControlBeforeInfo(control.Name)) {
            index++;
        }

        return index;
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
            _ => throw new ArgumentOutOfRangeException(nameof(field), field, "Unsupported RTF document information field.")
        };
    }
}
