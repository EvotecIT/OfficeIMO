namespace OfficeIMO.Visio;

public partial class VisioPage {
    /// <summary>Configures this page and checks identifiers before and after the callback.</summary>
    public VisioPage Configure(Action<VisioPage> configure) {
        ArgumentNullException.ThrowIfNull(configure); ValidateShapeIdentifiers(); configure(this); ValidateShapeIdentifiers(); return this;
    }

    internal void ValidateShapeIdentifiers() {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        void Visit(VisioShape shape) {
            if (string.IsNullOrWhiteSpace(shape.Id) || !ids.Add(shape.Id))
                throw new ArgumentException("Shape identifiers must be nonempty and unique on the page.");
            foreach (var child in shape.Children) Visit(child);
        }
        foreach (var shape in Shapes) Visit(shape);
    }
}
