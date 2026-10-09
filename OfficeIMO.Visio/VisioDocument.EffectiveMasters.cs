namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Resolves the master the page writer will use without attaching it to the source shape.</summary>
    internal VisioMaster? ResolveEffectiveMaster(VisioShape shape) {
        if (shape.Master != null) return shape.Master;
        if (!UseMastersByDefault) return null;
        string name = shape.NameU?.Trim() ?? string.Empty;
        return name.Length > 0 && TryEnsureBuiltinMaster(name, out VisioMaster? master) ? master : null;
    }

    /// <summary>Includes the dynamic-master identity preserved when a connector was loaded.</summary>
    internal VisioMaster? ResolveEffectiveMaster(VisioConnector connector) =>
        connector.Kind == ConnectorKind.Dynamic && (UseMastersByDefault || connector.PreserveDynamicConnectorMaster)
            ? EnsureBuiltinMaster("Dynamic connector") : null;
}
