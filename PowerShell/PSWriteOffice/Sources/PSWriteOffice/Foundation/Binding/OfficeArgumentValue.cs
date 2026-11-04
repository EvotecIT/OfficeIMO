using System;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

internal static class OfficeArgumentValue {
    internal static object? Unwrap(object? value) {
        while (value is PSObject wrapped) value = wrapped.BaseObject;
        return value;
    }

    internal static ArgumentTransformationMetadataException Invalid(string message) => new(message);
}
