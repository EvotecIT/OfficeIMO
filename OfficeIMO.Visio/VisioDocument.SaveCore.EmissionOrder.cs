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

        private static bool TryGetNextUnemittedShape(
            IReadOnlyList<VisioShape> shapes,
            ISet<VisioShape> emittedShapes,
            ref int nextShapeIndex,
            out VisioShape? shape) {
            while (nextShapeIndex < shapes.Count) {
                VisioShape candidate = shapes[nextShapeIndex++];
                if (emittedShapes.Contains(candidate)) {
                    continue;
                }

                shape = candidate;
                return true;
            }

            shape = null;
            return false;
        }

        private static bool TryGetNextUnemittedConnector(
            IReadOnlyList<VisioConnector> connectors,
            ISet<VisioConnector> emittedConnectors,
            ref int nextConnectorIndex,
            out VisioConnector? connector) {
            while (nextConnectorIndex < connectors.Count) {
                VisioConnector candidate = connectors[nextConnectorIndex++];
                if (emittedConnectors.Contains(candidate)) {
                    continue;
                }

                connector = candidate;
                return true;
            }

            connector = null;
            return false;
        }
    }
}
