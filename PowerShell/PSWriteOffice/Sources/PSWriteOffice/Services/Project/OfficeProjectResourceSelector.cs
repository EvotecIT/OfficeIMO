using System.Management.Automation;
using OfficeIMO.Project;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Project;

/// <summary>Selects a project resource by stable UID, literal name or native object.</summary>
public sealed class OfficeProjectResourceSelector {
    private readonly object _value;
    private OfficeProjectResourceSelector(object value) => _value = value;
    /// <summary>Creates a selector; numeric strings remain literal names.</summary>
    public static OfficeProjectResourceSelector From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw new PSArgumentException("A resource is required.");
        if (value is OfficeProjectResourceSelector selector) return selector;
        if (value is ProjectResource || value is string text && !string.IsNullOrWhiteSpace(text)) return new(value);
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try { return new(Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture)); }
            catch (OverflowException) { throw new PSArgumentException("The UID must fit an Int32."); }
        }
        throw new PSArgumentException("Use a native ProjectResource, an integer UID or a literal name.");
    }
    internal ProjectResource Resolve(ProjectDocument document) {
        if (_value is int uid) return document.Resources.GetByUid(uid);
        var matches = document.Resources.Where(item => _value is ProjectResource selected
            ? ReferenceEquals(item, selected) : string.Equals(item.Name, (string)_value, StringComparison.OrdinalIgnoreCase)).Take(2).ToArray();
        if (matches.Length == 1) return matches[0];
        throw new PSArgumentException(matches.Length == 0 ? "The resource does not belong to this project." : "The name is ambiguous; use a stable UID or native object.");
    }
}

/// <summary>Binds project resource selectors.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeProjectResourceTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => OfficeProjectResourceSelector.From(inputData);
}
