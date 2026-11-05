using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Project;

/// <summary>Selects a project task by stable UID, literal name or native object.</summary>
public sealed class OfficeProjectTaskSelector {
    private readonly object _value;
    private OfficeProjectTaskSelector(object value) => _value = value;
    /// <summary>Creates a selector; numeric strings remain literal names.</summary>
    public static OfficeProjectTaskSelector From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw new PSArgumentException("A task is required.");
        if (value is OfficeProjectTaskSelector selector) return selector;
        if (value is ProjectTask || value is string text && !string.IsNullOrWhiteSpace(text)) return new(value);
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try { return new(Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture)); }
            catch (OverflowException) { throw new PSArgumentException("The UID must fit an Int32."); }
        }
        throw new PSArgumentException("Use a native ProjectTask, an integer UID or a literal name.");
    }
    internal ProjectTask Resolve(ProjectDocument document) {
        if (_value is int uid) return document.Tasks.GetByUid(uid);
        var matches = document.AllTasks.Where(item => _value is ProjectTask selected
            ? ReferenceEquals(item, selected) : string.Equals(item.Name, (string)_value, StringComparison.OrdinalIgnoreCase)).Take(2).ToArray();
        if (matches.Length == 1) return matches[0];
        throw new PSArgumentException(matches.Length == 0 ? "The task does not belong to this project." : "The name is ambiguous; use a stable UID or native object.");
    }
}

/// <summary>Binds project task selectors.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeProjectTaskTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => OfficeProjectTaskSelector.From(inputData);
}
