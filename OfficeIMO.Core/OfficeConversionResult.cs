using System;

namespace OfficeIMO;

/// <summary>A converted value and its report. Format-owned results derive from it and narrow <see cref="Report"/>.</summary>
/// <typeparam name="T">The target document, text, artifact or other converted value.</typeparam>
public class OfficeConversionResult<T> {
    /// <summary>Creates a result containing a non-null converted value and its report.</summary>
    public OfficeConversionResult(T value, OfficeConversionReport report) {
        if (value is null) throw new ArgumentNullException(nameof(value));
        Value = value;
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Gets the converted value.</summary>
    public T Value { get; }

    /// <summary>Gets the conversion report.</summary>
    public OfficeConversionReport Report { get; }

    /// <summary>Gets whether the report describes possible fidelity loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Throws the report's <see cref="OfficeConversionException"/> when it describes possible fidelity loss.</summary>
    public void ThrowIfLoss() => Report.ThrowIfLoss();

    /// <summary>Deconstructs the result into its converted value and report.</summary>
    public void Deconstruct(out T value, out OfficeConversionReport report) {
        value = Value;
        report = Report;
    }
}

/// <summary>A converted value with its format-owned report, for operations that return typed evidence without a dedicated result type.</summary>
/// <typeparam name="T">The converted value.</typeparam>
/// <typeparam name="TReport">The format-owned report.</typeparam>
public sealed class OfficeConversionResult<T, TReport> : OfficeConversionResult<T> where TReport : OfficeConversionReport {
    /// <summary>Creates a result containing a non-null converted value and its typed report.</summary>
    public OfficeConversionResult(T value, TReport report) : base(value, report) {
        Report = report;
    }

    /// <summary>Gets the typed conversion report.</summary>
    public new TReport Report { get; }
}

/// <summary>Describes whether an OfficeIMO operation completed successfully.</summary>
public interface IOfficeResult {
    /// <summary>Gets whether the operation completed successfully.</summary>
    bool Succeeded { get; }
}

/// <summary>Describes an OfficeIMO operation that can produce a reference-type value.</summary>
/// <typeparam name="TValue">Value produced by the operation.</typeparam>
public interface IOfficeResult<out TValue> : IOfficeResult where TValue : class {
    /// <summary>Gets the produced value, or <see langword="null"/> when the operation did not complete.</summary>
    TValue? Value { get; }

    /// <summary>Returns the produced value or throws when the operation did not complete.</summary>
    TValue RequireValue();
}

/// <summary>Raised when a conversion reports loss that the caller does not accept, or produces no usable value.</summary>
/// <remarks>Format-owned exceptions derive from this type and carry their typed failure evidence.</remarks>
public class OfficeConversionException : InvalidOperationException {
    /// <summary>Creates an exception carrying the completed report.</summary>
    public OfficeConversionException(string message, OfficeConversionReport report, Exception? innerException = null)
        : base(message, innerException) {
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Gets the report produced before the operation stopped.</summary>
    public OfficeConversionReport Report { get; }
}
