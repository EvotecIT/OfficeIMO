using System;
using System.Collections.Generic;
using System.Globalization;
using System.Numerics;
using System.Text;

namespace OfficeIMO.Patching;

/// <summary>
/// RFC 6901 JSON Pointer syntax, shared by the document engines that apply RFC 6902 patches: parsing a pointer into its
/// reference tokens and reading a token as an array index.
/// </summary>
internal static class JsonPointer {
    /// <summary>
    /// Splits a pointer into its unescaped reference tokens. The empty pointer is the whole document and has no tokens.
    /// </summary>
    /// <returns><see langword="null"/> when the pointer is malformed, with the reason in <paramref name="error"/>.</returns>
    public static string[]? Parse(string pointer, out string? error) {
        error = null;
        if (pointer.Length == 0) return Array.Empty<string>();
        if (pointer[0] != '/') {
            error = "a JSON Pointer is empty or starts with '/'";
            return null;
        }
        var tokens = new List<string>();
        var token = new StringBuilder();
        for (int i = 1; i <= pointer.Length; i++) {
            if (i == pointer.Length || pointer[i] == '/') {
                tokens.Add(token.ToString());
                token.Clear();
                continue;
            }
            char c = pointer[i];
            if (c != '~') {
                token.Append(c);
                continue;
            }
            char next = i + 1 < pointer.Length ? pointer[i + 1] : '\0';
            if (next == '0') token.Append('~');
            else if (next == '1') token.Append('/');
            else {
                error = "'~' is followed by '0' or '1' in a JSON Pointer";
                return null;
            }
            i++;
        }
        return tokens.ToArray();
    }

    /// <summary>
    /// Reads a token as an array index: <c>0</c> or digits without a leading zero, below <paramref name="limit"/>. With
    /// <paramref name="allowEnd"/>, <c>-</c> stands for the index after the last item.
    /// </summary>
    public static bool TryGetIndex(string token, int count, int limit, bool allowEnd, out int index) {
        index = -1;
        if (allowEnd && token == "-") {
            index = count;
            return true;
        }
        if (token.Length == 0 || token.Length > 10 || (token.Length > 1 && token[0] == '0')) return false;
        foreach (char c in token) {
            if (c < '0' || c > '9') return false;
        }
        if (!int.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out int value) || value >= limit) return false;
        index = value;
        return true;
    }

    /// <summary>Returns whether <paramref name="prefix"/> is a proper prefix of <paramref name="tokens"/>.</summary>
    public static bool IsProperPrefix(string[] prefix, string[] tokens) {
        if (prefix.Length >= tokens.Length) return false;
        for (int i = 0; i < prefix.Length; i++) {
            if (!string.Equals(prefix[i], tokens[i], StringComparison.Ordinal)) return false;
        }
        return true;
    }

    /// <summary>Returns whether two token lists address the same location.</summary>
    public static bool AreEqual(string[] left, string[] right) {
        if (left.Length != right.Length) return false;
        for (int i = 0; i < left.Length; i++) {
            if (!string.Equals(left[i], right[i], StringComparison.Ordinal)) return false;
        }
        return true;
    }
}

/// <summary>The six RFC 6902 operations.</summary>
internal enum JsonPatchOperationKind : byte {
    Add,
    Remove,
    Replace,
    Move,
    Copy,
    Test
}

/// <summary>Reading the <c>op</c> member of an RFC 6902 operation.</summary>
internal static class JsonPatchOperations {
    public static bool TryParse(string name, out JsonPatchOperationKind kind) {
        switch (name) {
            case "add": kind = JsonPatchOperationKind.Add; return true;
            case "remove": kind = JsonPatchOperationKind.Remove; return true;
            case "replace": kind = JsonPatchOperationKind.Replace; return true;
            case "move": kind = JsonPatchOperationKind.Move; return true;
            case "copy": kind = JsonPatchOperationKind.Copy; return true;
            case "test": kind = JsonPatchOperationKind.Test; return true;
            default: kind = default; return false;
        }
    }

    public static bool NeedsFrom(JsonPatchOperationKind kind) => kind is JsonPatchOperationKind.Move or JsonPatchOperationKind.Copy;

    public static bool NeedsValue(JsonPatchOperationKind kind) => kind is JsonPatchOperationKind.Add or JsonPatchOperationKind.Replace or JsonPatchOperationKind.Test;
}

/// <summary>JSON number equality across integer and floating-point values, without rounding either.</summary>
internal static class JsonNumbers {
    /// <summary>
    /// Returns whether an integer equals a floating-point value exactly: the float is finite, whole and the same number.
    /// Converting the integer to <see cref="double"/> instead would round integers above 2^53, so 9007199254740993
    /// would equal 9007199254740992.0.
    /// </summary>
    public static bool IntegerEquals(BigInteger integer, double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value) && Math.Floor(value) == value && new BigInteger(value) == integer;
}
