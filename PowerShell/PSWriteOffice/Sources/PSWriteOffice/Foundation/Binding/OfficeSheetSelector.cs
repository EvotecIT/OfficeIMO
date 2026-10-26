using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

/// <summary>A sheet selected by name, one-based position, or the sheet object itself.</summary>
public sealed class OfficeSheetSelector {
    private OfficeSheetSelector(string? name, int? position, object? sheet) { Name = name; Position = position; Sheet = sheet; }
    /// <summary>Gets the literal name. A numeric string remains a name.</summary>
    public string? Name { get; }
    /// <summary>Gets the one-based position, when selected numerically.</summary>
    public int? Position { get; }
    /// <summary>Gets an existing sheet object, when supplied.</summary>
    public object? Sheet { get; }

    /// <summary>Creates a selector without converting numeric strings into positions.</summary>
    public static OfficeSheetSelector From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw OfficeArgumentValue.Invalid("A sheet is required.");
        if (value is OfficeSheetSelector selector) return selector;
        if (value is string name) {
            if (string.IsNullOrWhiteSpace(name)) throw OfficeArgumentValue.Invalid("A sheet name cannot be empty.");
            return new(name, null, null);
        }
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try {
                int position = checked(Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture));
                if (position >= 1) return new(null, position, null);
            } catch (OverflowException) { }
            throw OfficeArgumentValue.Invalid("Sheet positions start at 1 and must fit an Int32.");
        }
        // Restrict object binding to actual worksheet model objects, without depending on an engine-specific selector.
        if (value is OfficeIMO.Excel.ExcelSheet) return new(null, null, value);
        throw OfficeArgumentValue.Invalid("Expected a sheet name, a one-based integer or an ExcelSheet.");
    }

    /// <summary>Resolves against the format's sheet collection, checking membership and position.</summary>
    public T Resolve<T>(IReadOnlyList<T> sheets, Func<T, string> getName) where T : class {
        if (sheets == null) throw new ArgumentNullException(nameof(sheets));
        if (getName == null) throw new ArgumentNullException(nameof(getName));
        if (Position.HasValue) {
            if (Position.Value <= sheets.Count) return sheets[Position.Value - 1];
        } else {
            foreach (T item in sheets) {
                if (Sheet != null ? ReferenceEquals(item, Sheet) : string.Equals(getName(item), Name, StringComparison.OrdinalIgnoreCase)) return item;
            }
        }
        throw new ArgumentException("The selected sheet does not belong to this document.");
    }
}

/// <summary>Binds sheet names, integer positions and existing sheet objects.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeSheetTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => OfficeSheetSelector.From(inputData);
}
