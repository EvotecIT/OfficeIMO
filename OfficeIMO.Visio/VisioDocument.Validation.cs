using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Authoring-time validation helpers for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>
        /// Validates the in-memory document model and returns authoring-time issues.
        /// </summary>
        public IReadOnlyList<string> Validate() {
            List<string> issues = new();
            HashSet<int> pageIds = new();

            foreach (VisioPage page in _pages) {
                if (!pageIds.Add(page.Id)) {
                    issues.Add($"Duplicate page id '{page.Id}' found for page '{page.Name}'.");
                }

                if (page.Width <= 0) {
                    issues.Add($"Page '{page.Name}' must have a positive width.");
                }

                if (page.Height <= 0) {
                    issues.Add($"Page '{page.Name}' must have a positive height.");
                }

                HashSet<string> ids = new(StringComparer.Ordinal);
                HashSet<VisioShape> pageShapes = new();

                void ReserveId(string? id, string kind) {
                    if (string.IsNullOrWhiteSpace(id)) {
                        issues.Add($"{kind} id cannot be null or whitespace on page '{page.Name}'.");
                        return;
                    }

                    string idValue = id!;
                    if (!ids.Add(idValue)) {
                        issues.Add($"Duplicate {kind.ToLowerInvariant()} id '{idValue}' found on page '{page.Name}'.");
                    }
                }

                void VisitShape(VisioShape shape) {
                    ReserveId(shape.Id, "Shape");
                    pageShapes.Add(shape);

                    if (shape.Width < 0) {
                        issues.Add($"Shape '{shape.Id}' on page '{page.Name}' cannot have a negative width.");
                    }

                    if (shape.Height < 0) {
                        issues.Add($"Shape '{shape.Id}' on page '{page.Name}' cannot have a negative height.");
                    }

                    foreach (VisioShape child in shape.Children) {
                        if (!ReferenceEquals(child.Parent, shape)) {
                            issues.Add($"Child shape '{child.Id}' on page '{page.Name}' has an inconsistent parent reference.");
                        }

                        VisitShape(child);
                    }
                }

                foreach (VisioShape shape in page.Shapes) {
                    VisitShape(shape);
                }

                foreach (VisioConnector connector in page.Connectors) {
                    ReserveId(connector.Id, "Connector");

                    if (!pageShapes.Contains(connector.From)) {
                        issues.Add($"Connector '{connector.Id}' on page '{page.Name}' references a source shape that is not part of the page.");
                    }

                    if (!pageShapes.Contains(connector.To)) {
                        issues.Add($"Connector '{connector.Id}' on page '{page.Name}' references a target shape that is not part of the page.");
                    }

                    if (connector.FromConnectionPoint != null && !connector.From.ConnectionPoints.Contains(connector.FromConnectionPoint)) {
                        issues.Add($"Connector '{connector.Id}' on page '{page.Name}' references a source connection point that does not belong to shape '{connector.From.Id}'.");
                    }

                    if (connector.ToConnectionPoint != null && !connector.To.ConnectionPoints.Contains(connector.ToConnectionPoint)) {
                        issues.Add($"Connector '{connector.Id}' on page '{page.Name}' references a target connection point that does not belong to shape '{connector.To.Id}'.");
                    }
                }
            }

            return issues;
        }
    }
}
