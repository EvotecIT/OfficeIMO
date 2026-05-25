namespace System.Management.Automation;

internal sealed class PSObject {
    public PSObject(params PSPropertyInfo[] properties) {
        Properties = properties;
    }

    public IReadOnlyList<PSPropertyInfo> Properties { get; }
}

internal sealed class PSPropertyInfo {
    public PSPropertyInfo(string name, object? value, bool isGettable = true) {
        Name = name;
        Value = value;
        IsGettable = isGettable;
    }

    public string Name { get; }

    public object? Value { get; }

    public bool IsGettable { get; }
}
