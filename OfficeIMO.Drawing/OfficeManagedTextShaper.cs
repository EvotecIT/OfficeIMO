using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

internal static class OfficeManagedTextShaper {
    internal static OfficeManagedTextFallback Resolve(
        string text,
        OfficeTrueTypeFont font,
        System.Threading.CancellationToken cancellationToken = default) {
        if (string.IsNullOrEmpty(text) || !RequiresComplexLayout(text)) {
            return new OfficeManagedTextFallback(text ?? string.Empty, used: false, incomplete: false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        bool incomplete =
            OfficeTextElements.ContainsShapingRequiredScript(text) ||
            OfficeTextElements.ContainsBidiControl(text) ||
            (OfficeTextElements.ContainsJoiningScript(text) &&
             !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(text));
        string contextual = OfficeArabicTextShaper.Shape(text);
        cancellationToken.ThrowIfCancellationRequested();
        string visual = ToVisualOrder(contextual, cancellationToken);
        if (font.HasGlyphs(visual)) {
            return new OfficeManagedTextFallback(visual, used: true, incomplete);
        }

        cancellationToken.ThrowIfCancellationRequested();
        string reordered = ToVisualOrder(
            OfficeArabicTextShaper.ToLogicalText(text),
            cancellationToken);
        if (font.HasGlyphs(reordered)) {
            return new OfficeManagedTextFallback(reordered, used: true, incomplete: true);
        }

        return new OfficeManagedTextFallback(text, used: false, incomplete: true);
    }

    internal static bool RequiresComplexLayout(string? text) =>
        OfficeTextElements.ContainsRightToLeft(text) ||
        OfficeTextElements.ContainsJoiningScript(text) ||
        OfficeTextElements.ContainsShapingRequiredScript(text) ||
        OfficeTextElements.ContainsBidiControl(text);

    internal static string ToVisualOrder(
        string? value,
        System.Threading.CancellationToken cancellationToken = default) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        string withoutControls = RemoveBidiControls(value!, cancellationToken);
        var groups = new List<DirectionalGroup>();
        var current = new StringBuilder();
        TextElementDirection? direction = null;
        OfficeTextDirection baseDirection = OfficeTextElements.ResolveBaseDirection(withoutControls);
        TextElementDirection neutralDefault =
            baseDirection == OfficeTextDirection.RightToLeft
                ? TextElementDirection.RightToLeft
                : TextElementDirection.LeftToRight;
        int elementIndex = 0;
        foreach (string element in OfficeTextElements.Enumerate(withoutControls)) {
            if ((elementIndex++ & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            TextElementDirection resolved = ResolveDirection(element);
            if (resolved == TextElementDirection.Neutral) {
                resolved = direction ?? neutralDefault;
            }
            if (direction.HasValue && direction.Value != resolved) {
                groups.Add(new DirectionalGroup(current.ToString(), direction.Value));
                current.Clear();
            }
            direction = resolved;
            current.Append(element);
        }
        if (current.Length > 0) {
            groups.Add(new DirectionalGroup(current.ToString(), direction ?? neutralDefault));
        }

        if (baseDirection == OfficeTextDirection.RightToLeft) groups.Reverse();
        var visual = new StringBuilder(withoutControls.Length);
        foreach (DirectionalGroup group in groups) {
            cancellationToken.ThrowIfCancellationRequested();
            if (group.Direction == TextElementDirection.RightToLeft) {
                var elements = new List<string>(OfficeTextElements.Enumerate(group.Text));
                for (int index = elements.Count - 1; index >= 0; index--) visual.Append(elements[index]);
            } else {
                visual.Append(group.Text);
            }
        }
        return visual.ToString();
    }

    private static string RemoveBidiControls(
        string value,
        System.Threading.CancellationToken cancellationToken) {
        if (!OfficeTextElements.ContainsBidiControl(value)) return value;
        var result = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            char character = value[index];
            if (character == '\u061C' || character == '\u200E' || character == '\u200F' ||
                character >= '\u202A' && character <= '\u202E' ||
                character >= '\u2066' && character <= '\u2069') {
                continue;
            }
            result.Append(character);
        }
        return result.ToString();
    }

    private static TextElementDirection ResolveDirection(string element) {
        bool hasDigit = false;
        for (int index = 0; index < element.Length;) {
            int scalarIndex = index;
            int scalar = ReadScalar(element, ref index);
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(element, scalarIndex);
            if (category == UnicodeCategory.DecimalDigitNumber) {
                hasDigit = true;
                continue;
            }
            if (OfficeTextElements.IsRightToLeftScalar(scalar)) return TextElementDirection.RightToLeft;
            if (category is UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or
                UnicodeCategory.TitlecaseLetter or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter) {
                return TextElementDirection.LeftToRight;
            }
        }
        return hasDigit ? TextElementDirection.LeftToRight : TextElementDirection.Neutral;
    }

    private static int ReadScalar(string text, ref int index) {
        char first = text[index++];
        return char.IsHighSurrogate(first) &&
               index < text.Length &&
               char.IsLowSurrogate(text[index])
            ? char.ConvertToUtf32(first, text[index++])
            : first;
    }

    private enum TextElementDirection {
        Neutral,
        LeftToRight,
        RightToLeft
    }

    private readonly struct DirectionalGroup {
        internal DirectionalGroup(string text, TextElementDirection direction) {
            Text = text;
            Direction = direction;
        }

        internal string Text { get; }
        internal TextElementDirection Direction { get; }
    }
}

internal readonly struct OfficeManagedTextFallback {
    internal OfficeManagedTextFallback(string text, bool used, bool incomplete) {
        Text = text;
        Used = used;
        Incomplete = incomplete;
    }

    internal string Text { get; }
    internal bool Used { get; }
    internal bool Incomplete { get; }
}
