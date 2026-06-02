using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private static PackageRelationship GetRequiredSingleRelationship(IEnumerable<PackageRelationship> relationships, string description) {
            List<PackageRelationship> matches = relationships.ToList();
            if (matches.Count != 1) {
                throw new InvalidDataException($"Expected exactly one {description} relationship, but found {matches.Count}.");
            }

            return matches[0];
        }

        private static VisioConnectionPoint? ResolveConnectionPoint(VisioShape shape, string? connectionCell) {
            if (string.IsNullOrWhiteSpace(connectionCell) || string.Equals(connectionCell, "PinX", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            const string prefix = "Connections.X";
            string resolvedConnectionCell = connectionCell!;
            if (!resolvedConnectionCell.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string suffix = resolvedConnectionCell.Substring(prefix.Length);
            if (!int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)) {
                return null;
            }

            int sectionIndex = index - 1;
            foreach (VisioConnectionPoint point in shape.ConnectionPoints) {
                if (point.SectionIndex == sectionIndex) {
                    return point;
                }
            }

            bool hasExplicitSectionIndices = shape.ConnectionPoints.Any(point => point.SectionIndex.HasValue);
            if (hasExplicitSectionIndices) {
                return null;
            }

            return sectionIndex >= 0 && sectionIndex < shape.ConnectionPoints.Count ? shape.ConnectionPoints[sectionIndex] : null;
        }

        private static void CaptureConnectAttributes(XElement connectElement, IList<XAttribute> preservedAttributes, IList<XName> attributeOrder) {
            preservedAttributes.Clear();
            attributeOrder.Clear();
            foreach (XAttribute attribute in connectElement.Attributes()) {
                if (attribute.IsNamespaceDeclaration) {
                    continue;
                }

                attributeOrder.Add(attribute.Name);

                if (string.Equals(attribute.Name.LocalName, "FromSheet", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "FromCell", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "ToSheet", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "ToCell", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                preservedAttributes.Add(new XAttribute(attribute));
            }
        }

        private static void CopyPreservedAttributes(IEnumerable<XAttribute> source, IList<XAttribute> destination) {
            destination.Clear();
            foreach (XAttribute attribute in source) {
                destination.Add(new XAttribute(attribute));
            }
        }

        private static void CopyPreservedAttributeOrder(IEnumerable<XName> source, IList<XName> destination) {
            destination.Clear();
            foreach (XName attributeName in source) {
                destination.Add(attributeName);
            }
        }
    }
}
