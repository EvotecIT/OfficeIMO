using System;
using System.Collections;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>Binds inspection policy hashtables or the shared typed Reader option families.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeInspectionOptionsTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? input = OfficeArgumentValue.Unwrap(inputData);
        object[] values = input is object[] array ? array : new[] { input! };
        return values.Select(item => {
            object? value = OfficeArgumentValue.Unwrap(item);
            if (value is IDictionary dictionary && dictionary.Keys.Cast<object>().All(key => key is string name &&
                typeof(OfficeInspectOptions).GetProperty(name, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance) != null))
                return new OfficeOptionsTransformationAttribute(typeof(OfficeInspectOptions)).Transform(engineIntrinsics, value);
            return ((object[])new OfficeReadOptionsTransformationAttribute().Transform(engineIntrinsics, new[] { value! }))[0];
        }).ToArray();
    }
}
