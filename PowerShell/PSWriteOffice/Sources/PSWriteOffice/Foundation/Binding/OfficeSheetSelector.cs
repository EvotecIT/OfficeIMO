using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Foundation;

/// <summary>Binds sheet names, integer positions and existing sheet objects to the native immutable selector.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeSheetTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        try {
            object? input = OfficeArgumentValue.Unwrap(inputData);
            if (input is Array items) {
                var selectors = new OfficeSheetSelector[items.Length];
                for (int index = 0; index < items.Length; index++)
                    selectors[index] = OfficeSheetSelector.From(OfficeArgumentValue.Unwrap(items.GetValue(index))!);
                return selectors;
            }
            return OfficeSheetSelector.From(input!);
        } catch (ArgumentException exception) {
            throw OfficeArgumentValue.Invalid(exception.Message);
        }
    }
}
