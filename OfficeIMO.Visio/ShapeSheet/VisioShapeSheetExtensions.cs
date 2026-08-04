using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>Typed access to source-preserved Visio ShapeSheet sections.</summary>
    public static class VisioShapeSheetExtensions {
        /// <summary>Returns unmodeled ShapeSheet sections retained on a shape.</summary>
        public static IReadOnlyList<VisioShapeSheetSection> GetShapeSheetSections(
            this VisioShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            return ReadSections(shape.PreservedNonGeometrySections);
        }

        /// <summary>Returns unmodeled ShapeSheet sections retained on a connector.</summary>
        public static IReadOnlyList<VisioShapeSheetSection> GetShapeSheetSections(
            this VisioConnector connector) {
            if (connector == null) throw new ArgumentNullException(nameof(connector));
            return ReadSections(connector.PreservedNonGeometrySections);
        }

        /// <summary>Sets a typed ShapeSheet section without disturbing other preserved sections.</summary>
        public static VisioShape SetShapeSheetSection(this VisioShape shape,
            VisioShapeSheetSection section) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            SetSection(shape.PreservedNonGeometrySections, section);
            return shape;
        }

        /// <summary>Sets a typed ShapeSheet section without disturbing other preserved sections.</summary>
        public static VisioConnector SetShapeSheetSection(this VisioConnector connector,
            VisioShapeSheetSection section) {
            if (connector == null) throw new ArgumentNullException(nameof(connector));
            SetSection(connector.PreservedNonGeometrySections, section);
            return connector;
        }

        /// <summary>Removes a named unmodeled ShapeSheet section from a shape.</summary>
        public static bool RemoveShapeSheetSection(this VisioShape shape, string name) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            return RemoveSection(shape.PreservedNonGeometrySections, name);
        }

        /// <summary>Removes a named unmodeled ShapeSheet section from a connector.</summary>
        public static bool RemoveShapeSheetSection(this VisioConnector connector, string name) {
            if (connector == null) throw new ArgumentNullException(nameof(connector));
            return RemoveSection(connector.PreservedNonGeometrySections, name);
        }

        private static IReadOnlyList<VisioShapeSheetSection> ReadSections(
            IEnumerable<XElement> sections) => sections
            .Where(element => string.Equals(element.Name.LocalName, "Section",
                StringComparison.OrdinalIgnoreCase))
            .Select(element => new VisioShapeSheetSection(element)).ToList();

        private static void SetSection(IList<XElement> sections,
            VisioShapeSheetSection section) {
            if (section == null) throw new ArgumentNullException(nameof(section));
            string? name = section.Name;
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Section name cannot be empty.", nameof(section));
            for (int index = 0; index < sections.Count; index++) {
                if (string.Equals((string?)sections[index].Attribute("N"), name,
                        StringComparison.OrdinalIgnoreCase)) {
                    sections[index] = section.ToXElement();
                    return;
                }
            }
            sections.Add(section.ToXElement());
        }

        private static bool RemoveSection(IList<XElement> sections, string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Section name cannot be empty.", nameof(name));
            for (int index = sections.Count - 1; index >= 0; index--) {
                if (string.Equals((string?)sections[index].Attribute("N"), name,
                        StringComparison.OrdinalIgnoreCase)) {
                    sections.RemoveAt(index);
                    return true;
                }
            }
            return false;
        }
    }
}
