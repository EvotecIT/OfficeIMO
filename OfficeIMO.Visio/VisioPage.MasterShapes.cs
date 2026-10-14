using System;

namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>
        /// Adds a shape to the page.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="master">Master associated with the shape.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="w">Width of the shape.</param>
        /// <param name="h">Height of the shape.</param>
        /// <param name="text">Optional text override. Imported, grouped and styled masters retain their text when omitted.</param>
        /// <param name="unit">
        /// Optional measurement unit. When omitted, values are interpreted using
        /// the page <see cref="DefaultUnit"/>.
        /// </param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, VisioMaster master, double x, double y, double w, double h, string? text = null, VisioMeasurementUnit? unit = null) {
            VisioMeasurementUnit effectiveUnit = unit ?? DefaultUnit;
            x = x.ToInches(effectiveUnit);
            y = y.ToInches(effectiveUnit);
            w = w.ToInches(effectiveUnit);
            h = h.ToInches(effectiveUnit);

            if (VisioDocument.RequiresMasterInstanceCopy(master))
                return VisioDuplicationExtensions.CreateMasterInstance(this, master, id, x, y, w, h, text);

            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) {
                Master = master,
                NameU = master.NameU,
                TextStyle = VisioDocument.CreateSimpleMasterTextFrameStyle(master, w, h)
            };
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds a shape using a document-registered master by its NameU.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="masterNameU">Registered master universal name.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="w">Width.</param>
        /// <param name="h">Height.</param>
        /// <param name="text">Optional text override. Imported, grouped and styled masters retain their text when omitted.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (OwnerDocument == null) {
                throw new InvalidOperationException("This page is not attached to a VisioDocument, so master lookup by name is unavailable.");
            }

            VisioMaster master = OwnerDocument.GetMaster(masterNameU);
            x = x.ToInches(unit);
            y = y.ToInches(unit);
            w = w.ToInches(unit);
            h = h.ToInches(unit);

            if (VisioDocument.RequiresMasterInstanceCopy(master))
                return VisioDuplicationExtensions.CreateMasterInstance(this, master, id, x, y, w, h, text);

            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) {
                Master = master,
                NameU = master.NameU,
                TextStyle = VisioDocument.CreateSimpleMasterTextFrameStyle(master, w, h)
            };
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>Adds a shape from a registered master using the page's default measurement unit.</summary>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h) =>
            AddShape(id, masterNameU, x, y, w, h, text: null, unit: DefaultUnit);

        /// <summary>
        /// Adds a shape using the page <see cref="DefaultUnit"/> and a document-registered master.
        /// </summary>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h, string? text = null) =>
            AddShape(id, masterNameU, x, y, w, h, text, DefaultUnit);
    }
}
