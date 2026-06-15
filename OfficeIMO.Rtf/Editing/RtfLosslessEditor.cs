using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <summary>
/// Applies targeted edits to an RTF syntax tree while preserving untouched raw RTF.
/// </summary>
public sealed partial class RtfLosslessEditor {
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
    /// Adds, replaces, or removes a root document variable while preserving the rest of the RTF stream.
    /// Pass <c>null</c> for <paramref name="value"/> to remove all variables with the supplied name.
    /// </summary>
    public void SetDocumentVariable(string name, string? value) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Document variable name cannot be empty.", nameof(name));
        }

        RtfGroup root = SetDocumentVariable(_syntaxTree.Root, name, value);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or replaces a custom document property in the root <c>userprops</c> destination while preserving the rest of the RTF stream.
    /// </summary>
    public void SetUserProperty(RtfUserProperty property) {
        if (property == null) {
            throw new ArgumentNullException(nameof(property));
        }

        RtfGroup root = SetUserPropertyInRoot(_syntaxTree.Root, property);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or replaces a text custom document property in the root <c>userprops</c> destination.
    /// </summary>
    public void SetUserProperty(string name, string value) {
        SetUserProperty(RtfUserProperty.Text(name, value));
    }

    /// <summary>
    /// Removes all custom document properties with the supplied name while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveUserProperty(string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Custom property name cannot be empty.", nameof(name));
        }

        RtfGroup root = RemoveUserPropertyFromRoot(_syntaxTree.Root, name);
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

    private static RtfGroup SetUserPropertyInRoot(RtfGroup root, RtfUserProperty property) {
        var children = new List<RtfNode>(root.Children);
        int userPropertiesIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "userprops");
        if (userPropertiesIndex >= 0) {
            children[userPropertiesIndex] = SetUserProperty((RtfGroup)children[userPropertiesIndex], property);
        } else {
            children.Insert(GetUserPropertiesInsertIndex(children), CreateUserPropertiesGroup(property));
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveUserPropertyFromRoot(RtfGroup root, string name) {
        var children = new List<RtfNode>(root.Children);
        int userPropertiesIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "userprops");
        if (userPropertiesIndex < 0) {
            return root;
        }

        RtfGroup? updatedUserProperties = RemoveUserProperty((RtfGroup)children[userPropertiesIndex], name);
        if (updatedUserProperties == null) {
            children.RemoveAt(userPropertiesIndex);
        } else {
            children[userPropertiesIndex] = updatedUserProperties;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetUserProperty(RtfGroup userPropertiesGroup, RtfUserProperty property) {
        var children = new List<RtfNode>(userPropertiesGroup.Children);
        List<UserPropertyRange> ranges = GetUserPropertyRanges(children);
        bool replaced = false;

        for (int index = ranges.Count - 1; index >= 0; index--) {
            UserPropertyRange range = ranges[index];
            if (!string.Equals(range.Name, property.Name, StringComparison.Ordinal)) {
                continue;
            }

            if (replaced) {
                children.RemoveRange(range.StartIndex, range.Length);
            } else {
                children.RemoveRange(range.StartIndex, range.Length);
                children.InsertRange(range.StartIndex, CreateUserPropertyNodes(property));
                replaced = true;
            }
        }

        if (!replaced) {
            children.AddRange(CreateUserPropertyNodes(property));
        }

        return new RtfGroup(userPropertiesGroup.Position, children);
    }

    private static RtfGroup? RemoveUserProperty(RtfGroup userPropertiesGroup, string name) {
        var children = new List<RtfNode>(userPropertiesGroup.Children);
        List<UserPropertyRange> ranges = GetUserPropertyRanges(children);

        for (int index = ranges.Count - 1; index >= 0; index--) {
            UserPropertyRange range = ranges[index];
            if (string.Equals(range.Name, name, StringComparison.Ordinal)) {
                children.RemoveRange(range.StartIndex, range.Length);
            }
        }

        return children.Any(node => node is RtfGroup group && group.Destination == "propname")
            ? new RtfGroup(userPropertiesGroup.Position, children)
            : null;
    }

    private static RtfGroup SetDocumentVariable(RtfGroup root, string name, string? value) {
        var children = new List<RtfNode>(root.Children);
        bool replaced = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfGroup group || group.Destination != "docvar") {
                continue;
            }

            if (!DocumentVariableNameMatches(group, name)) {
                continue;
            }

            if (value == null || replaced) {
                children.RemoveAt(index);
            } else {
                children[index] = CreateDocumentVariableGroup(name, value);
                replaced = true;
            }
        }

        if (!replaced && value != null) {
            children.Insert(GetDocumentVariableInsertIndex(children), CreateDocumentVariableGroup(name, value));
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup CreateDocumentVariableGroup(string name, string value) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "docvar", null, hasParameter: false, rawText: @"\docvar "),
            CreatePlainGroup(name),
            CreatePlainGroup(value)
        });
    }

    private static RtfGroup CreatePlainGroup(string text) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfText(0, text, RtfTextEncoding.EncodeText(text))
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

    private static int GetDocumentVariableInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "docvar" || group.Destination == "userprops" || group.Destination == "info")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }

    private static int GetUserPropertiesInsertIndex(List<RtfNode> children) {
        int infoIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "info");
        if (infoIndex >= 0) {
            return infoIndex + 1;
        }

        int firstDocVarIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "docvar");
        return firstDocVarIndex >= 0 ? firstDocVarIndex : GetInfoInsertIndex(children);
    }

    private static RtfGroup CreateUserPropertiesGroup(RtfUserProperty property) {
        var children = new List<RtfNode> {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "userprops", null, hasParameter: false, rawText: @"\userprops")
        };
        children.AddRange(CreateUserPropertyNodes(property));
        return new RtfGroup(0, children);
    }

    private static IReadOnlyList<RtfNode> CreateUserPropertyNodes(RtfUserProperty property) {
        var nodes = new List<RtfNode> {
            new RtfGroup(0, new RtfNode[] {
                new RtfControlWord(0, "propname", null, hasParameter: false, rawText: @"\propname "),
                new RtfText(0, property.Name, RtfTextEncoding.EncodeText(property.Name))
            })
        };

        if (property.TypeCode.HasValue) {
            nodes.Add(new RtfControlWord(
                0,
                "proptype",
                property.TypeCode.Value,
                hasParameter: true,
                rawText: @"\proptype" + property.TypeCode.Value.ToString(CultureInfo.InvariantCulture)));
        }

        AddUserPropertyValue(nodes, "staticval", property.StaticValue);
        AddUserPropertyValue(nodes, "linkval", property.LinkedValue);
        return nodes;
    }

    private static void AddUserPropertyValue(List<RtfNode> nodes, string destination, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        nodes.Add(new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, destination, null, hasParameter: false, rawText: "\\" + destination + " "),
            new RtfText(0, value!, RtfTextEncoding.EncodeText(value!))
        }));
    }

    private static List<UserPropertyRange> GetUserPropertyRanges(IReadOnlyList<RtfNode> children) {
        var ranges = new List<UserPropertyRange>();
        int startIndex = -1;
        string? name = null;

        for (int index = 0; index < children.Count; index++) {
            if (children[index] is not RtfGroup group || group.Destination != "propname") {
                continue;
            }

            if (startIndex >= 0 && name != null) {
                ranges.Add(new UserPropertyRange(startIndex, index - startIndex, name));
            }

            startIndex = index;
            name = CollectPlainText(group).Trim();
        }

        if (startIndex >= 0 && name != null) {
            ranges.Add(new UserPropertyRange(startIndex, children.Count - startIndex, name));
        }

        return ranges;
    }

    private static bool DocumentVariableNameMatches(RtfGroup group, string name) {
        RtfGroup? nameGroup = group.Children
            .OfType<RtfGroup>()
            .FirstOrDefault();

        return nameGroup != null && string.Equals(CollectPlainText(nameGroup).Trim(), name, StringComparison.Ordinal);
    }

    private static string CollectPlainText(RtfGroup group) {
        var builder = new StringBuilder();
        var state = new DocumentVariableTextState();
        AppendPlainText(group, builder, state);
        return builder.ToString();
    }

    private static void AppendPlainText(RtfGroup group, StringBuilder builder, DocumentVariableTextState state) {
        foreach (RtfNode child in group.Children) {
            switch (child) {
                case RtfText text:
                    AppendWithSkip(text.Text, builder, state);
                    break;
                case RtfControlSymbol symbol:
                    AppendControlSymbolText(symbol, builder, state);
                    break;
                case RtfControlWord control when control.Name == "uc" && control.Parameter.HasValue && control.Parameter.Value >= 0:
                    state.UnicodeSkipCount = control.Parameter.Value;
                    break;
                case RtfControlWord control when control.Name == "u" && control.Parameter.HasValue:
                    AppendUnicodeValue(control.Parameter.Value, builder, state);
                    state.SkipCharacters = state.UnicodeSkipCount;
                    break;
                case RtfControlWord control when control.Name == "tab":
                    AppendWithSkip("\t", builder, state);
                    break;
                case RtfControlWord control when control.Name == "line" || control.Name == "par":
                    AppendWithSkip(Environment.NewLine, builder, state);
                    break;
                case RtfGroup childGroup:
                    AppendPlainText(childGroup, builder, state);
                    break;
            }
        }
    }

    private static void AppendControlSymbolText(RtfControlSymbol symbol, StringBuilder builder, DocumentVariableTextState state) {
        if (symbol.Symbol == '\'' && symbol.Parameter.HasValue) {
            AppendWithSkip(((char)symbol.Parameter.Value).ToString(), builder, state);
        } else if (symbol.Symbol == '\\' || symbol.Symbol == '{' || symbol.Symbol == '}') {
            AppendWithSkip(symbol.Symbol.ToString(), builder, state);
        } else if (symbol.Symbol == '~') {
            AppendWithSkip("\u00A0", builder, state);
        } else if (symbol.Symbol == '_') {
            AppendWithSkip("\u2011", builder, state);
        } else if (symbol.Symbol == '-') {
            AppendWithSkip("\u00AD", builder, state);
        }
    }

    private static void AppendUnicodeValue(int value, StringBuilder builder, DocumentVariableTextState state) {
        int codePoint = value < 0 ? value + 65536 : value;
        AppendWithSkip(char.ConvertFromUtf32(codePoint), builder, state);
    }

    private static void AppendWithSkip(string text, StringBuilder builder, DocumentVariableTextState state) {
        if (state.SkipCharacters <= 0) {
            builder.Append(text);
            return;
        }

        if (state.SkipCharacters >= text.Length) {
            state.SkipCharacters -= text.Length;
            return;
        }

        builder.Append(text, state.SkipCharacters, text.Length - state.SkipCharacters);
        state.SkipCharacters = 0;
    }

    private sealed class DocumentVariableTextState {
        internal int UnicodeSkipCount { get; set; } = 1;
        internal int SkipCharacters { get; set; }
    }

    private readonly struct UserPropertyRange {
        internal UserPropertyRange(int startIndex, int length, string name) {
            StartIndex = startIndex;
            Length = length;
            Name = name;
        }

        internal int StartIndex { get; }

        internal int Length { get; }

        internal string Name { get; }
    }

}
