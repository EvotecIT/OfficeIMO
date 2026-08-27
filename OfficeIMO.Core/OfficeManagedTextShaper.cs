using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

internal static class OfficeManagedTextShaper {
    internal static OfficeManagedTextFallback Resolve(
        string text,
        IOfficeFontProgram font,
        System.Threading.CancellationToken cancellationToken = default) {
        if (string.IsNullOrEmpty(text) || !RequiresComplexLayout(text)) {
            return new OfficeManagedTextFallback(text ?? string.Empty, used: false, incomplete: false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        bool incomplete =
            OfficeTextElements.ContainsShapingRequiredScript(text) ||
            OfficeTextElements.ContainsVariationSelector(text) ||
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
        OfficeTextElements.ContainsBidiControl(text) ||
        OfficeTextElements.ContainsVariationSelector(text);

    internal static string ToVisualOrder(
        string? value,
        System.Threading.CancellationToken cancellationToken = default) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        return OfficeBidiTextResolver.ToVisualOrder(value, OfficeTextDirection.Auto, cancellationToken);
    }

    internal static IReadOnlyList<T> ToVisualOrder<T>(
        IReadOnlyList<T> elements,
        Func<T, string> textSelector,
        System.Threading.CancellationToken cancellationToken = default) {
        return ToVisualOrder(elements, textSelector, OfficeTextDirection.Auto, cancellationToken);
    }

    internal static IReadOnlyList<T> ToVisualOrder<T>(
        IReadOnlyList<T> elements,
        Func<T, string> textSelector,
        OfficeTextDirection requestedDirection,
        System.Threading.CancellationToken cancellationToken = default) {
        if (elements.Count == 0) return Array.Empty<T>();

        var completeText = new StringBuilder();
        foreach (T element in elements) completeText.Append(textSelector(element));
        var groups = new List<DirectionalElementGroup<T>>();
        var current = new List<T>();
        TextElementDirection? direction = null;
        OfficeTextDirection baseDirection = requestedDirection == OfficeTextDirection.Auto
            ? OfficeTextElements.ResolveBaseDirection(completeText.ToString())
            : requestedDirection;
        TextElementDirection neutralDefault =
            baseDirection == OfficeTextDirection.RightToLeft
                ? TextElementDirection.RightToLeft
                : TextElementDirection.LeftToRight;
        int elementIndex = 0;
        foreach (T element in elements) {
            if ((elementIndex++ & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            TextElementDirection resolved = ResolveDirection(textSelector(element));
            if (resolved == TextElementDirection.Neutral) {
                resolved = direction ?? neutralDefault;
            }
            if (direction.HasValue && direction.Value != resolved) {
                groups.Add(new DirectionalElementGroup<T>(current.ToArray(), direction.Value));
                current.Clear();
            }
            direction = resolved;
            current.Add(element);
        }
        if (current.Count > 0) {
            groups.Add(new DirectionalElementGroup<T>(current.ToArray(), direction ?? neutralDefault));
        }

        if (baseDirection == OfficeTextDirection.RightToLeft) groups.Reverse();
        var visual = new List<T>(elements.Count);
        foreach (DirectionalElementGroup<T> group in groups) {
            cancellationToken.ThrowIfCancellationRequested();
            if (group.Direction == TextElementDirection.RightToLeft) {
                for (int index = group.Elements.Count - 1; index >= 0; index--) visual.Add(group.Elements[index]);
            } else {
                visual.AddRange(group.Elements);
            }
        }
        return visual;
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

    private readonly struct DirectionalElementGroup<T> {
        internal DirectionalElementGroup(IReadOnlyList<T> elements, TextElementDirection direction) {
            Elements = elements;
            Direction = direction;
        }

        internal IReadOnlyList<T> Elements { get; }
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
