using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds or replaces a font-table entry by its RTF font identifier while preserving the rest of the RTF stream.
    /// </summary>
    public void SetFont(RtfFont font) {
        if (font == null) {
            throw new ArgumentNullException(nameof(font));
        }

        if (font.Id < 0) {
            throw new ArgumentOutOfRangeException(nameof(font), "RTF font identifiers cannot be negative.");
        }

        RtfGroup root = SetFont(_syntaxTree.Root, font);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds or replaces a font-table entry by its RTF font identifier and display name.
    /// </summary>
    public void SetFont(int id, string name) {
        SetFont(new RtfFont(id, name));
    }

    /// <summary>
    /// Removes all font-table entries with the supplied RTF font identifier while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveFont(int id) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "RTF font identifiers cannot be negative.");
        }

        RtfGroup root = RemoveFont(_syntaxTree.Root, id);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetFont(RtfGroup root, RtfFont font) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "fonttbl");

        if (tableIndex >= 0) {
            children[tableIndex] = SetFontInTable((RtfGroup)children[tableIndex], font);
            return new RtfGroup(root.Position, children);
        }

        children.Insert(GetFontTableInsertIndex(children), CreateFontTable(font));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup RemoveFont(RtfGroup root, int id) {
        var children = new List<RtfNode>(root.Children);
        int tableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "fonttbl");
        if (tableIndex < 0) {
            return root;
        }

        RtfGroup? updatedTable = RemoveFontFromTable((RtfGroup)children[tableIndex], id);
        if (updatedTable == null) {
            children.RemoveAt(tableIndex);
        } else {
            children[tableIndex] = updatedTable;
        }

        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup SetFontInTable(RtfGroup fontTable, RtfFont font) {
        var children = new List<RtfNode>(fontTable.Children);
        bool replaced = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfGroup fontGroup || !FontIdMatches(fontGroup, font.Id)) {
                continue;
            }

            if (replaced) {
                children.RemoveAt(index);
            } else {
                children[index] = CreateFontGroup(font);
                replaced = true;
            }
        }

        if (!replaced) {
            children.Add(CreateFontGroup(font));
        }

        return new RtfGroup(fontTable.Position, children);
    }

    private static RtfGroup? RemoveFontFromTable(RtfGroup fontTable, int id) {
        var children = new List<RtfNode>(fontTable.Children);

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is RtfGroup fontGroup && FontIdMatches(fontGroup, id)) {
                children.RemoveAt(index);
            }
        }

        return children.Any(node => node is RtfGroup group && group.Children.OfType<RtfControlWord>().Any(control => control.Name == "f"))
            ? new RtfGroup(fontTable.Position, children)
            : null;
    }

    private static bool FontIdMatches(RtfGroup fontGroup, int id) {
        return fontGroup.Children.OfType<RtfControlWord>().Any(control => control.Name == "f" && control.Parameter == id);
    }

    private static RtfGroup CreateFontTable(RtfFont font) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(0, "fonttbl", null, hasParameter: false, rawText: @"\fonttbl"),
            CreateFontGroup(font)
        });
    }

    private static RtfGroup CreateFontGroup(RtfFont font) {
        var children = new List<RtfNode> {
            new RtfControlWord(0, "f", font.Id, hasParameter: true, rawText: @"\f" + font.Id.ToString(CultureInfo.InvariantCulture))
        };

        AddFontFamily(children, font.Family);
        AddOptionalFontNumber(children, "fcharset", font.Charset);
        AddOptionalFontNumber(children, "fprq", font.Pitch);
        AddOptionalFontNumber(children, "cpg", font.CodePage);
        AddOptionalFontNumber(children, "fbias", font.Bias);
        AddFontDestination(children, "panose", font.Panose);
        AddFontDestination(children, "fname", font.NonTaggedName);
        AddFontEmbedding(children, font.Embedding);
        children.Add(new RtfText(0, " " + font.Name, " " + RtfTextEncoding.EncodeText(font.Name)));
        AddFontDestination(children, "falt", font.AlternateName);
        children.Add(new RtfText(0, ";", ";"));
        return new RtfGroup(0, children);
    }

    private static void AddFontFamily(List<RtfNode> children, RtfFontFamily? family) {
        string? controlName = GetFontFamilyControlName(family);
        if (controlName == null) {
            return;
        }

        children.Add(new RtfControlWord(0, controlName, null, hasParameter: false, rawText: "\\" + controlName));
    }

    private static string? GetFontFamilyControlName(RtfFontFamily? family) {
        if (!family.HasValue) {
            return null;
        }

        switch (family.Value) {
            case RtfFontFamily.Roman:
                return "froman";
            case RtfFontFamily.Swiss:
                return "fswiss";
            case RtfFontFamily.Modern:
                return "fmodern";
            case RtfFontFamily.Script:
                return "fscript";
            case RtfFontFamily.Decorative:
                return "fdecor";
            case RtfFontFamily.Technical:
                return "ftech";
            case RtfFontFamily.Bidirectional:
                return "fbidi";
            default:
                return "fnil";
        }
    }

    private static void AddOptionalFontNumber(List<RtfNode> children, string controlName, int? value) {
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

    private static void AddFontDestination(List<RtfNode> children, string destination, string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }

        children.Add(new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, destination, null, hasParameter: false, rawText: "\\" + destination + " "),
            new RtfText(0, value!.Trim(), RtfTextEncoding.EncodeText(value!.Trim()))
        }));
    }

    private static void AddFontEmbedding(List<RtfNode> children, RtfFontEmbedding? embedding) {
        if (embedding == null) {
            return;
        }

        var embeddingChildren = new List<RtfNode> {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "fontemb", null, hasParameter: false, rawText: @"\fontemb"),
            new RtfControlWord(
                0,
                embedding.Type == RtfEmbeddedFontType.TrueType ? "fttruetype" : "ftnil",
                null,
                hasParameter: false,
                rawText: embedding.Type == RtfEmbeddedFontType.TrueType ? @"\fttruetype" : @"\ftnil")
        };

        if (!string.IsNullOrWhiteSpace(embedding.FileName) || embedding.FileCodePage.HasValue) {
            var fileChildren = new List<RtfNode> {
                new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
                new RtfControlWord(0, "fontfile", null, hasParameter: false, rawText: @"\fontfile")
            };
            AddOptionalFontNumber(fileChildren, "cpg", embedding.FileCodePage);
            if (!string.IsNullOrWhiteSpace(embedding.FileName)) {
                string trimmedFileName = embedding.FileName!.Trim();
                fileChildren.Add(new RtfText(0, " " + trimmedFileName, " " + RtfTextEncoding.EncodeText(trimmedFileName)));
            }

            embeddingChildren.Add(new RtfGroup(0, fileChildren));
        }

        if (embedding.Data.Length > 0) {
            embeddingChildren.Add(new RtfText(0, " " + FormatHexBytes(embedding.Data), " " + FormatHexBytes(embedding.Data)));
        }

        children.Add(new RtfGroup(0, embeddingChildren));
    }

    private static string FormatHexBytes(byte[] data) {
        var builder = new StringBuilder(data.Length * 2);
        foreach (byte value in data) {
            builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static int GetFontTableInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group && (group.Destination == "fonttbl" || group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }
}
