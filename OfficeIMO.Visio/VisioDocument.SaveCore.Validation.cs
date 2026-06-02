using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private static void ValidatePagesForSave(IEnumerable<VisioPage> pages) {
            foreach (VisioPage page in pages) {
                HashSet<string> ids = new(StringComparer.Ordinal);

                void Reserve(string id, string kind) {
                    if (string.IsNullOrWhiteSpace(id)) {
                        throw new InvalidOperationException($"{kind} id cannot be null or whitespace on page '{page.Name}'.");
                    }

                    if (!ids.Add(id)) {
                        throw new InvalidOperationException($"Duplicate {kind.ToLowerInvariant()} id '{id}' found on page '{page.Name}'.");
                    }
                }

                void VisitShape(VisioShape shape) {
                    Reserve(shape.Id, "Shape");
                    foreach (VisioShape child in shape.Children) {
                        VisitShape(child);
                    }
                }

                foreach (VisioShape shape in page.Shapes) {
                    VisitShape(shape);
                }

                foreach (VisioConnector connector in page.Connectors) {
                    Reserve(connector.Id, "Connector");
                }
            }
        }
    }
}
