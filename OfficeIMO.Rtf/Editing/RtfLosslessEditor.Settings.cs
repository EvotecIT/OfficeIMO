using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Sets the root document character set declaration and optionally replaces the root ANSI code page.
    /// </summary>
    public void SetCharacterSet(RtfDocumentCharacterSet characterSet, int? ansiCodePage = null) {
        ValidateCharacterSet(characterSet);
        ValidateNonNegative(ansiCodePage, nameof(ansiCodePage));

        RtfGroup root = SetRootControl(_syntaxTree.Root, GetCharacterSetControlName(characterSet), null, CharacterSetControlNames);
        root = SetRootControl(root, "ansicpg", ansiCodePage, SingleControlName("ansicpg"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes the root ANSI code page control while preserving the rest of the RTF stream.
    /// </summary>
    public void SetAnsiCodePage(int? codePage) {
        ValidateNonNegative(codePage, nameof(codePage));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "ansicpg", codePage, SingleControlName("ansicpg"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes the root default font id control while preserving the rest of the RTF stream.
    /// </summary>
    public void SetDefaultFont(int? fontId) {
        ValidateNonNegative(fontId, nameof(fontId));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "deff", fontId, SingleControlName("deff"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Adds, replaces, or removes the root Unicode fallback skip count while preserving the rest of the RTF stream.
    /// </summary>
    public void SetUnicodeSkipCount(int? count) {
        ValidateNonNegative(count, nameof(count));

        RtfGroup root = SetRootControl(_syntaxTree.Root, "uc", count, SingleControlName("uc"));
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    private static RtfGroup SetRootControl(RtfGroup root, string controlName, int? parameter, ISet<string> replacedControlNames) {
        var children = new List<RtfNode>(root.Children);
        RemoveRootControls(children, replacedControlNames);

        if (ShouldInsertRootControl(controlName, parameter)) {
            children.Insert(GetRootControlInsertIndex(children, controlName), CreateRootControl(controlName, parameter));
        }

        return new RtfGroup(root.Position, children);
    }

    private static void RemoveRootControls(List<RtfNode> children, ISet<string> controlNames) {
        int bodyStartIndex = GetRootBodyStartIndex(children);
        for (int index = bodyStartIndex - 1; index >= 0; index--) {
            if (children[index] is RtfControlWord control && controlNames.Contains(control.Name)) {
                children.RemoveAt(index);
            }
        }
    }

    private static bool ShouldInsertRootControl(string controlName, int? parameter) {
        return controlName == "ansi" || controlName == "mac" || controlName == "pc" || controlName == "pca" || IsRootParameterlessControl(controlName) || parameter.HasValue;
    }

    private static bool IsRootParameterlessControl(string controlName) {
        switch (controlName) {
            case "rtlgutter":
            case "pgnrestart":
            case "pgncont":
            case "pgndec":
            case "pgnucrm":
            case "pgnlcrm":
            case "pgnucltr":
            case "pgnlcltr":
            case "pgndecd":
            case "landscape":
            case "titlepg":
                return true;
            default:
                return false;
        }
    }

    private static int GetRootControlInsertIndex(IReadOnlyList<RtfNode> children, string controlName) {
        int bodyStartIndex = GetRootBodyStartIndex(children);
        int order = GetRootControlOrder(controlName);
        int insertIndex = 0;

        for (int index = 0; index < bodyStartIndex; index++) {
            if (children[index] is not RtfControlWord control) {
                continue;
            }

            if (GetRootControlOrder(control.Name) <= order) {
                insertIndex = index + 1;
            }
        }

        return insertIndex;
    }

    private static int GetRootBodyStartIndex(IReadOnlyList<RtfNode> children) {
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfControlWord control && IsRootBodyStartControl(control.Name)) {
                return index;
            }
        }

        return children.Count;
    }

    private static bool IsRootBodyStartControl(string controlName) {
        return controlName == "pard" || controlName == "sectd" || controlName == "sect";
    }

    private static RtfControlWord CreateRootControl(string controlName, int? parameter) {
        string rawText = parameter.HasValue
            ? "\\" + controlName + parameter.Value.ToString(CultureInfo.InvariantCulture)
            : "\\" + controlName;
        return new RtfControlWord(0, controlName, parameter, parameter.HasValue, rawText);
    }

    private static int GetRootControlOrder(string controlName) {
        switch (controlName) {
            case "rtf":
                return 0;
            case "ansi":
            case "mac":
            case "pc":
            case "pca":
                return 1;
            case "ansicpg":
                return 2;
            case "deff":
                return 3;
            case "uc":
                return 4;
            case "paperw":
                return 10;
            case "paperh":
                return 11;
            case "psz":
                return 12;
            case "binfsxn":
                return 13;
            case "binsxn":
                return 14;
            case "margl":
                return 20;
            case "margr":
                return 21;
            case "margt":
                return 22;
            case "margb":
                return 23;
            case "gutter":
                return 24;
            case "headery":
                return 25;
            case "footery":
                return 26;
            case "rtlgutter":
                return 27;
            case "pgnstarts":
                return 30;
            case "pgnrestart":
            case "pgncont":
                return 31;
            case "pgnx":
                return 32;
            case "pgny":
                return 33;
            case "pgndec":
            case "pgnucrm":
            case "pgnlcrm":
            case "pgnucltr":
            case "pgnlcltr":
            case "pgndecd":
                return 34;
            case "landscape":
                return 40;
            case "titlepg":
                return 41;
            default:
                return 100;
        }
    }

    private static string GetCharacterSetControlName(RtfDocumentCharacterSet characterSet) {
        switch (characterSet) {
            case RtfDocumentCharacterSet.Ansi:
                return "ansi";
            case RtfDocumentCharacterSet.Mac:
                return "mac";
            case RtfDocumentCharacterSet.Pc:
                return "pc";
            case RtfDocumentCharacterSet.Pca:
                return "pca";
            default:
                throw new ArgumentOutOfRangeException(nameof(characterSet), "Unsupported RTF document character set.");
        }
    }

    private static ISet<string> SingleControlName(string controlName) {
        return new HashSet<string>(StringComparer.Ordinal) {
            controlName
        };
    }

    private static void ValidateNonNegative(int? value, string parameterName) {
        if (value.HasValue && value.Value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "RTF document setting cannot be negative.");
        }
    }

    private static void ValidateCharacterSet(RtfDocumentCharacterSet characterSet) {
        if (characterSet != RtfDocumentCharacterSet.Ansi &&
            characterSet != RtfDocumentCharacterSet.Mac &&
            characterSet != RtfDocumentCharacterSet.Pc &&
            characterSet != RtfDocumentCharacterSet.Pca) {
            throw new ArgumentOutOfRangeException(nameof(characterSet), "Unsupported RTF document character set.");
        }
    }

    private static readonly ISet<string> CharacterSetControlNames = new HashSet<string>(StringComparer.Ordinal) {
        "ansi",
        "mac",
        "pc",
        "pca"
    };
}
