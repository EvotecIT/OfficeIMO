using System;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Foundation;

/// <summary>A live slide or a one-based slide position in a document.</summary>
public sealed class PowerPointSlideSelector {
    private PowerPointSlideSelector(PowerPointSlide? slide = null, int? position = null) { Slide = slide; Position = position; }
    /// <summary>Gets the supplied native slide.</summary>
    public PowerPointSlide? Slide { get; }
    /// <summary>Gets the one-based slide position.</summary>
    public int? Position { get; }

    /// <summary>Creates a selector without converting text or fractional numbers to slide positions.</summary>
    public static PowerPointSlideSelector From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw OfficeArgumentValue.Invalid("A slide is required.");
        if (value is PowerPointSlideSelector selector) return selector;
        if (value is PowerPointSlide slide) return new(slide: slide);
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try {
                int position = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                if (position >= 1) return new(position: position);
            } catch (OverflowException) { }
            throw OfficeArgumentValue.Invalid("Slide positions start at 1 and must fit an Int32.");
        }
        throw OfficeArgumentValue.Invalid("Expected a native slide object or a one-based integer slide position.");
    }

    /// <summary>Resolves the slide, checking membership when a document is supplied.</summary>
    internal PowerPointSlide Resolve(PowerPointDocument? document) {
        if (Slide != null && (document == null || document.Slides.Any(item => ReferenceEquals(item, Slide)))) return Slide;
        if (Position.HasValue && document != null && Position.Value <= document.Slides.Count) return document.Slides[Position.Value - 1];
        throw new ArgumentException(document == null ? "Supply -Document for a numeric slide position." : "The selected slide does not belong to this document.");
    }
}

/// <summary>Binds native slides and one-based integer positions.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class PowerPointSlideTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value is System.Collections.IEnumerable selections && value is not string)
            return selections.Cast<object>().Select(PowerPointSlideSelector.From).ToArray();
        return PowerPointSlideSelector.From(inputData);
    }
}
