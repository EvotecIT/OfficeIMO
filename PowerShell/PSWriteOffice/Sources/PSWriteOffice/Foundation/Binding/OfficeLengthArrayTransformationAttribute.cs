using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>Binds physical length lists with the same units and bare-point convention as scalar lengths.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeLengthArrayTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        var scalar = new OfficeLengthTransformationAttribute();
        if (value is IEnumerable values && value is not string) {
            var result = new List<OfficeLength>();
            foreach (object item in values) result.Add((OfficeLength)scalar.Transform(engineIntrinsics, item));
            return result.ToArray();
        }
        return new[] { (OfficeLength)scalar.Transform(engineIntrinsics, value!) };
    }
}
