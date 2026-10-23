using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

public static partial class OfficeArabicTextShaper {
    /// <summary>Resolves joining features in logical source order without replacing Unicode characters.</summary>
    /// <remarks>Two linear passes retain transparent marks and joining controls without rescanning their source spans.</remarks>
    internal static string?[]? GetOpenTypeJoiningForms(string source, Action? charge = null) {
        var result = new string?[source.Length];
        bool hasPrevious = false;
        ArabicForms previous = default;
        for (int index = 0; index < source.Length; index++) {
            charge?.Invoke();
            char current = LogicalForms.TryGetValue(source[index], out char logical) ? logical : source[index];
            int scalar = current;
            UnicodeCategory category = char.GetUnicodeCategory(current);
            if (char.IsHighSurrogate(current) && index + 1 < source.Length && char.IsLowSurrogate(source[index + 1])) {
                scalar = char.ConvertToUtf32(current, source[index + 1]);
                category = CharUnicodeInfo.GetUnicodeCategory(source, index);
            }
            bool mapped = Forms.TryGetValue(current, out ArabicForms forms);
            if (OfficeTextElements.IsJoiningScriptScalar(scalar) && IsLetter(category) && scalar != 0x0640 && !mapped)
                return null;
            if (mapped) {
                result[index] = forms.CanJoinPrevious && hasPrevious && previous.CanJoinNext ? "fina" : "isol";
                previous = forms; hasPrevious = true;
            } else if (current == '\u200C') hasPrevious = false;
            else if (IsTransparent(current)) { }
            else if (current == '\u0640') { previous = ArabicForms.JoiningCarrier; hasPrevious = true; }
            else hasPrevious = false;
        }
        bool hasNext = false;
        ArabicForms next = default;
        for (int index = source.Length - 1; index >= 0; index--) {
            charge?.Invoke();
            char current = LogicalForms.TryGetValue(source[index], out char logical) ? logical : source[index];
            if (Forms.TryGetValue(current, out ArabicForms forms)) {
                bool before = result[index] == "fina";
                bool after = forms.CanJoinNext && hasNext && next.CanJoinPrevious;
                result[index] = before ? after ? "medi" : "fina" : after ? "init" : "isol";
                next = forms; hasNext = true;
            } else if (current == '\u200C') hasNext = false;
            else if (IsTransparent(current)) { }
            else if (current == '\u0640') { next = ArabicForms.JoiningCarrier; hasNext = true; }
            else hasNext = false;
        }
        return result;
    }
}
