using System;
using System.Collections;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace PSWriteOffice.Foundation;

/// <summary>Binds typed read options, common option hashtables, or a format-keyed ReaderAllOptions hashtable.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeReadOptionsTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? input = OfficeArgumentValue.Unwrap(inputData);
        object[] values = input is object[] array ? array : new[] { input! };
        return values.Select(item => {
            object? value = OfficeArgumentValue.Unwrap(item);
            if (value is not IDictionary dictionary) return value!;
            bool common = dictionary.Keys.Cast<object>().All(key => key is string name &&
                typeof(ReaderOptions).GetProperty(name, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance) != null);
            return new OfficeOptionsTransformationAttribute(common ? typeof(ReaderOptions) : typeof(ReaderAllOptions)).Transform(engineIntrinsics, value);
        }).ToArray();
    }
}
