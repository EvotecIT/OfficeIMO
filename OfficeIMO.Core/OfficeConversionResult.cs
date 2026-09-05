using System;

namespace OfficeIMO;

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

/// <summary>
/// Describes the common value-and-report surface returned by OfficeIMO document conversions.
/// </summary>
/// <typeparam name="TValue">Native target document, text, artifact, or reference type.</typeparam>
/// <typeparam name="TReport">Typed fidelity report produced by the conversion.</typeparam>
public interface IOfficeConversionResult<out TValue, out TReport> : IOfficeResult<TValue>
    where TValue : class
    where TReport : class, IOfficeConversionReport {
    /// <summary>Gets the typed conversion report.</summary>
    TReport Report { get; }

    /// <summary>Gets whether the conversion reported possible fidelity loss.</summary>
    bool HasLoss { get; }

    /// <summary>Returns the produced value only when conversion completed without reported loss.</summary>
    TValue RequireNoLoss();
}

/// <summary>
/// Common implementation for conversions that return a native value together with a typed fidelity report.
/// </summary>
/// <typeparam name="TValue">Native target document, text, artifact, or reference type.</typeparam>
/// <typeparam name="TReport">Typed fidelity report produced by the conversion.</typeparam>
public class OfficeConversionResult<TValue, TReport> : IOfficeConversionResult<TValue, TReport>
    where TValue : class
    where TReport : class, IOfficeConversionReport {
    /// <summary>Creates a conversion result.</summary>
    public OfficeConversionResult(TValue value, TReport report)
    {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <inheritdoc />
    public TValue Value { get; }

    /// <inheritdoc />
    public TReport Report { get; }

    /// <inheritdoc />
    public virtual bool Succeeded => true;

    /// <inheritdoc />
    public virtual bool HasLoss => Report.HasLoss;

    /// <inheritdoc />
    public virtual TValue RequireValue() {
        if (!Succeeded) {
            throw new OfficeConversionException("The conversion did not produce a usable value.", Report);
        }

        return Value;
    }

    /// <inheritdoc />
    public virtual TValue RequireNoLoss() {
        TValue value = RequireValue();
        Report.RequireNoLoss();
        return value;
    }
}

/// <summary>Raised when a structured OfficeIMO conversion result has no usable value.</summary>
public sealed class OfficeConversionException : InvalidOperationException {
    /// <summary>Creates an exception for a failed conversion report.</summary>
    public OfficeConversionException(string message, IOfficeConversionReport report, Exception? innerException = null)
        : base(message, innerException) {
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Gets the report produced before conversion failed.</summary>
    public IOfficeConversionReport Report { get; }
}
