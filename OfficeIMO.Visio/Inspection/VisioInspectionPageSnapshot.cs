using System.Globalization;
using System.Text;

namespace OfficeIMO.Visio {
/// <summary>
    /// Snapshot of a Visio page.
    /// </summary>
    public sealed class VisioInspectionPageSnapshot {
        internal VisioInspectionPageSnapshot(
            int id,
            string name,
            string? nameU,
            double width,
            double height,
            IReadOnlyList<string> layers,
            IReadOnlyList<VisioInspectionShapeSnapshot> shapes,
            IReadOnlyList<VisioInspectionConnectorSnapshot> connectors) {
            Id = id;
            Name = name;
            NameU = nameU;
            Width = width;
            Height = height;
            Layers = layers;
            Shapes = shapes;
            Connectors = connectors;
        }

        /// <summary>Page identifier.</summary>
        public int Id { get; }

        /// <summary>Page display name.</summary>
        public string Name { get; }

        /// <summary>Page universal name.</summary>
        public string? NameU { get; }

        /// <summary>Page width in inches.</summary>
        public double Width { get; }

        /// <summary>Page height in inches.</summary>
        public double Height { get; }

        /// <summary>Layer names used on the page.</summary>
        public IReadOnlyList<string> Layers { get; }

        /// <summary>Shape snapshots on the page, including group children.</summary>
        public IReadOnlyList<VisioInspectionShapeSnapshot> Shapes { get; }

        /// <summary>Connector snapshots on the page.</summary>
        public IReadOnlyList<VisioInspectionConnectorSnapshot> Connectors { get; }

        internal void AppendText(StringBuilder builder) {
            string prefix = "page[" + Id.ToString(CultureInfo.InvariantCulture) + ":" + Escape + "]";
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".id", Id);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".nameU", NameU);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".width", Width);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".height", Height);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".layers", string.Join(",", Layers));
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".shapeCount", Shapes.Count);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".connectorCount", Connectors.Count);

            foreach (VisioInspectionShapeSnapshot shape in Shapes) {
                shape.AppendText(builder, prefix);
            }

            foreach (VisioInspectionConnectorSnapshot connector in Connectors) {
                connector.AppendText(builder, prefix);
            }
        }

        private string Escape => VisioInspectionSnapshot.EscapeKey(Name);
    }
}
