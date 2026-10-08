using System;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Strict VBA code-page conversion and source-file normalization.</summary>
internal static class OfficeVbaText {
    // Windows-1252 is the usual Office project encoding and has no in-box provider on modern .NET.
    private const string Windows1252High = "\u20ac\u0081\u201a\u0192\u201e\u2026\u2020\u2021\u02c6\u2030\u0160\u2039\u0152\u008d\u017d\u008f\u0090\u2018\u2019\u201c\u201d\u2022\u2013\u2014\u02dc\u2122\u0161\u203a\u0153\u009d\u017e\u0178";

    internal static string Decode(byte[] bytes, int codePage) {
        if (codePage != 1252) return GetEncoding(codePage).GetString(bytes);
        var chars = new char[bytes.Length];
        for (int index = 0; index < bytes.Length; index++) {
            byte value = bytes[index];
            chars[index] = value >= 128 && value < 160 ? Windows1252High[value - 128] : (char)value;
        }
        return new string(chars);
    }

    internal static byte[] Encode(string text, int codePage) {
        if (codePage != 1252) return GetEncoding(codePage).GetBytes(text);
        var bytes = new byte[text.Length];
        for (int index = 0; index < text.Length; index++) {
            char value = text[index];
            int special = Windows1252High.IndexOf(value);
            if (special >= 0) bytes[index] = (byte)(128 + special);
            else if (value < 128 || value >= 160 && value <= 255) bytes[index] = (byte)value;
            else throw new EncoderFallbackException("The VBA source contains a character outside the project's Windows-1252 code page.");
        }
        return bytes;
    }

    private static Encoding GetEncoding(int codePage) {
        try {
            return Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        } catch (ArgumentException exception) {
            throw new NotSupportedException($"VBA code page {codePage} is not available in this runtime. Register an encoding provider in the application when required.", exception);
        }
    }

    internal static string NormalizeSource(string source, string name, OfficeVbaModuleKind kind, string? existingSource = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (source.IndexOf('\0') >= 0) throw new ArgumentException("VBA source cannot contain NUL characters reserved for compression padding.", nameof(source));
        string[] lines = source.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        ValidateClassAttributes(lines);
        // VBE .cls exports have a designer preamble; the compound module holds only attributes and code.
        if (lines.Length > 0 && lines[0].TrimStart().StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase)) {
            int firstAttribute = Array.FindIndex(lines, line => HasAttributeName(line, "VB_Name"));
            if (firstAttribute < 0) throw new InvalidDataException("An exported class module must contain Attribute VB_Name.");
            lines = lines.Skip(firstAttribute).ToArray();
        }
        string expected = "Attribute VB_Name = \"" + name + "\"";
        lines = new[] { expected }.Concat(lines.Where(line => !HasAttributeName(line, "VB_Name"))).ToArray();
        if (existingSource != null) {
            if (kind == OfficeVbaModuleKind.Document || kind == OfficeVbaModuleKind.Designer) {
                string? persistedBase = GetBaseIdentity(existingSource);
                string? suppliedBase = GetBaseIdentity(string.Join("\r\n", lines));
                if (persistedBase != null && suppliedBase != null && !persistedBase.Equals(suppliedBase, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidDataException("Source replacement cannot change a host document or form designer's base identity.");
                }
            }
            var additions = existingSource.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
                .Where(line => GetClassAttributeName(line) is string attribute && !lines.Any(candidate => HasAttributeName(candidate, attribute)))
                .ToArray();
            lines = new[] { lines[0] }.Concat(additions).Concat(lines.Skip(1)).ToArray();
        }
        if (kind == OfficeVbaModuleKind.Class) {
            lines = AddMissingAttributes(lines, "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = False");
        }
        if (kind == OfficeVbaModuleKind.Class) {
            // Office's native generic class identity; VBE export omits this persisted attribute.
            lines = AddMissingAttributes(lines, "Attribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"");
        }
        if (kind == OfficeVbaModuleKind.Class || kind == OfficeVbaModuleKind.Document) {
            foreach (string attribute in new[] { "Attribute VB_TemplateDerived = False", "Attribute VB_Customizable = " + (kind == OfficeVbaModuleKind.Document ? "True" : "False") }) {
                lines = AddMissingAttributes(lines, attribute);
            }
        }
        ValidateClassAttributes(lines);
        return string.Join("\r\n", lines);
    }

    private static bool HasAttributeName(string line, string name) =>
        string.Equals(GetClassAttributeName(line), name, StringComparison.OrdinalIgnoreCase);

    private static string[] AddMissingAttributes(string[] lines, params string[] attributes) =>
        new[] { lines[0] }.Concat(attributes.Where(attribute => !lines.Any(line => HasAttributeName(line, GetClassAttributeName(attribute)!))))
            .Concat(lines.Skip(1)).ToArray();

    internal static void ValidateIdentifier(string name, int maximumLength = 31) {
        if (string.IsNullOrEmpty(name) || name.Length > maximumLength || !IsLetter(name[0])
            || name.Any(character => !IsLetter(character) && !(character >= '0' && character <= '9') && character != '_')) {
            throw new ArgumentException($"A VBA identifier must start with an ASCII letter and contain at most {maximumLength} letters, digits, or underscores.", nameof(name));
        }
    }

    internal static string? GetBaseIdentity(string source) {
        string[] lines = source.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        ValidateClassAttributes(lines);
        foreach (string line in lines) {
            if (!string.Equals(GetClassAttributeName(line), "VB_Base", StringComparison.OrdinalIgnoreCase)) continue;
            string value = line.Substring(line.IndexOf('=') + 1).Trim();
            if (value.Length < 3 || value[0] != '"' || value.IndexOf('"', 1) != value.Length - 1) {
                throw new InvalidDataException("VBA source has an invalid VB_Base string attribute.");
            }
            return value.Substring(1, value.Length - 2);
        }
        return null;
    }

    private static void ValidateClassAttributes(string[] lines) {
        var attributes = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string line in lines) {
            string? name = GetClassAttributeName(line);
            if (name != null && !attributes.Add(name)) throw new InvalidDataException("VBA source repeats a class attribute: " + name + ".");
        }
    }

    private static string? GetClassAttributeName(string line) {
        string text = line.TrimStart();
        if (!text.StartsWith("Attribute", StringComparison.OrdinalIgnoreCase) || text.Length <= 9 || !char.IsWhiteSpace(text[9])) return null;
        int equals = text.IndexOf('=');
        if (equals < 10) return null;
        string name = text.Substring(9, equals - 9).Trim();
        return name.StartsWith("VB_", StringComparison.OrdinalIgnoreCase) && name.IndexOf('.') < 0 ? name : null;
    }

    private static bool IsLetter(char value) => value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z';
}
