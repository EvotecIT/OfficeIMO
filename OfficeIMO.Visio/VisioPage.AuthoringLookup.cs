namespace OfficeIMO.Visio;

public partial class VisioPage {
    /// <summary>Gets an existing shape by identifier, throwing when it is missing.</summary>
    public VisioShape GetShape(string id) => this.FindShapeById(id) ?? throw new ArgumentException($"Unknown shape id '{id}'.", nameof(id));

    /// <summary>Connects existing shapes selected by their identifiers.</summary>
    public VisioConnector AddConnector(string fromId, string toId, ConnectorKind kind = ConnectorKind.Dynamic,
        VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) =>
        AddConnector(GetShape(fromId), GetShape(toId), kind, fromSide, toSide);

    /// <summary>Adds a container around existing shapes selected by identifier.</summary>
    public VisioShape AddContainer(string id, string text, IEnumerable<string> memberIds, VisioContainerOptions? options = null) {
        if (memberIds == null) throw new ArgumentNullException(nameof(memberIds));
        return AddContainer(id, text, memberIds.Select(GetShape), options);
    }

    /// <summary>Configures and adds a container around existing shapes selected by identifier.</summary>
    public VisioShape AddContainer(string id, string text, IEnumerable<string> memberIds, Action<VisioContainerOptions> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        var options = new VisioContainerOptions();
        configure(options);
        return AddContainer(id, text, memberIds, options);
    }
}
