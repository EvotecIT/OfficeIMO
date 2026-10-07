namespace OfficeIMO.Xps;

// Native accessibility metadata is separate from painted Unicode and inferred captions.
internal sealed class XpsGraphicDescription {
    private XpsGraphicDescription(string? name, string? help) {
        Name = name; HelpText = help;
    }
    internal string? Name { get; }
    internal string? HelpText { get; }
    internal string AlternativeText => string.Join("\n", new[] { Name, HelpText }.Where(value => !string.IsNullOrWhiteSpace(value)));
    internal int CharacterCount => (Name?.Length ?? 0) + (HelpText?.Length ?? 0);

    internal static XpsGraphicDescription? Read(XElement element) {
        if (element.Name.LocalName is not "Path" and not "Canvas") return null;
        string? name = (string?)element.Attribute("AutomationProperties.Name");
        string? help = (string?)element.Attribute("AutomationProperties.HelpText");
        return string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(help) ? null : new(name, help);
    }
}
