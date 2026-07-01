using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes reusable first-line and continuation-line offsets inside a text frame.
/// </summary>
public readonly struct OfficeTextParagraphIndent : IEquatable<OfficeTextParagraphIndent> {
    /// <summary>Represents text without paragraph indentation.</summary>
    public static OfficeTextParagraphIndent Empty { get; } = new OfficeTextParagraphIndent(0D, 0D);

    /// <summary>
    /// Creates a paragraph indentation descriptor.
    /// </summary>
    /// <param name="firstLineOffset">Additional offset applied to the first visual line of each paragraph.</param>
    /// <param name="continuationLineOffset">Additional offset applied to wrapped continuation lines.</param>
    public OfficeTextParagraphIndent(double firstLineOffset, double continuationLineOffset) {
        ValidateOffset(firstLineOffset, nameof(firstLineOffset));
        ValidateOffset(continuationLineOffset, nameof(continuationLineOffset));
        FirstLineOffset = firstLineOffset;
        ContinuationLineOffset = continuationLineOffset;
    }

    /// <summary>Additional offset applied to the first visual line of each paragraph.</summary>
    public double FirstLineOffset { get; }

    /// <summary>Additional offset applied to wrapped continuation lines.</summary>
    public double ContinuationLineOffset { get; }

    /// <summary>Largest line offset used by this paragraph indentation.</summary>
    public double MaximumOffset => Math.Max(FirstLineOffset, ContinuationLineOffset);

    /// <summary>Whether the indentation has no visual effect.</summary>
    public bool IsEmpty => FirstLineOffset == 0D && ContinuationLineOffset == 0D;

    /// <summary>Creates a first-line indentation descriptor.</summary>
    public static OfficeTextParagraphIndent FirstLine(double offset) => new OfficeTextParagraphIndent(offset, 0D);

    /// <summary>Creates a hanging indentation descriptor.</summary>
    public static OfficeTextParagraphIndent Hanging(double offset) => new OfficeTextParagraphIndent(0D, offset);

    /// <summary>Scales indentation offsets by a positive rendering factor.</summary>
    public OfficeTextParagraphIndent Scale(double scale) {
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Text paragraph indent scale must be a finite positive number.");
        }

        return IsEmpty ? Empty : new OfficeTextParagraphIndent(FirstLineOffset * scale, ContinuationLineOffset * scale);
    }

    /// <inheritdoc />
    public bool Equals(OfficeTextParagraphIndent other) =>
        FirstLineOffset.Equals(other.FirstLineOffset) && ContinuationLineOffset.Equals(other.ContinuationLineOffset);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeTextParagraphIndent other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + FirstLineOffset.GetHashCode();
            hash = (hash * 31) + ContinuationLineOffset.GetHashCode();
            return hash;
        }
    }

    /// <summary>Compares two indentation descriptors for equality.</summary>
    public static bool operator ==(OfficeTextParagraphIndent left, OfficeTextParagraphIndent right) => left.Equals(right);

    /// <summary>Compares two indentation descriptors for inequality.</summary>
    public static bool operator !=(OfficeTextParagraphIndent left, OfficeTextParagraphIndent right) => !left.Equals(right);

    private static void ValidateOffset(double value, string paramName) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Text paragraph indent offsets must be finite non-negative numbers.");
        }
    }
}
