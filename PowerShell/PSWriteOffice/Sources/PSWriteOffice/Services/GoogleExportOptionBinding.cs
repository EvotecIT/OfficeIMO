using System;
using System.Collections;
using System.Management.Automation;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Captures native Google export options before applying shell property overrides.</summary>
internal static class GoogleExportOptionBinding {
    internal static T Capture<T>(EngineIntrinsics engine, object[]? values, Func<T> create, Func<T, T> clone) where T : class {
        T result = create();
        bool typed = false;
        foreach (object? item in values ?? Array.Empty<object>()) {
            object? value = Unwrap(item);
            if (value is IDictionary) continue;
            if (value is not T options) throw new PSArgumentException("Expected " + typeof(T).Name + " or a property hashtable.");
            if (typed) throw new PSArgumentException("Provide at most one typed Google export options object.");
            typed = true;
            result = clone(options);
        }
        var binder = new OfficeOptionsTransformationAttribute(typeof(T));
        foreach (object? item in values ?? Array.Empty<object>())
            if (Unwrap(item) is IDictionary table) result = clone((T)binder.ApplyTo(engine, result, table));
        return result;
    }

    private static object? Unwrap(object? value) {
        while (value is PSObject wrapper && !ReferenceEquals(value, wrapper.BaseObject)) value = wrapper.BaseObject;
        return value;
    }
}
