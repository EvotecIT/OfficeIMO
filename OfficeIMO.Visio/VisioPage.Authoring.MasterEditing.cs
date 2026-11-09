using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>
        /// Replaces the master for a known shape id while preserving page placement, text, metadata, and connector endpoints.
        /// </summary>
        public VisioPage ReplaceMaster(string shapeId, string masterNameU, bool resizeToMaster = false) {
            this.ReplaceMaster(GetShape(shapeId), masterNameU, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for a known shape id using an existing master instance.
        /// </summary>
        public VisioPage ReplaceMaster(string shapeId, VisioMaster master, bool resizeToMaster = false) {
            this.ReplaceMaster(GetShape(shapeId), master, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for a known shape id using an OfficeIMO stencil definition.
        /// </summary>
        public VisioPage ReplaceMaster(string shapeId, VisioStencilShape stencil, bool resizeToMaster = false) {
            this.ReplaceMaster(GetShape(shapeId), stencil, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for every selected shape.
        /// </summary>
        public VisioPage ReplaceMasters(Func<VisioShape, bool> predicate, string masterNameU, bool resizeToMaster = false) {
            ArgumentNullException.ThrowIfNull(predicate);

            this.SelectShapes(predicate).ReplaceMaster(masterNameU, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for every selected shape using an existing master instance.
        /// </summary>
        public VisioPage ReplaceMasters(Func<VisioShape, bool> predicate, VisioMaster master, bool resizeToMaster = false) {
            ArgumentNullException.ThrowIfNull(predicate);

            this.SelectShapes(predicate).ReplaceMaster(master, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for every selected shape using an OfficeIMO stencil definition.
        /// </summary>
        public VisioPage ReplaceMasters(Func<VisioShape, bool> predicate, VisioStencilShape stencil, bool resizeToMaster = false) {
            ArgumentNullException.ThrowIfNull(predicate);

            this.SelectShapes(predicate).ReplaceMaster(stencil, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces shapes using a matching current master universal name.
        /// </summary>
        public VisioPage ReplaceMastersByMaster(string currentMasterNameU, string replacementMasterNameU, bool resizeToMaster = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            this.SelectByMaster(currentMasterNameU, comparison).ReplaceMaster(replacementMasterNameU, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces shapes using a matching current master universal name and an OfficeIMO stencil definition.
        /// </summary>
        public VisioPage ReplaceMastersByMaster(string currentMasterNameU, VisioStencilShape replacementStencil, bool resizeToMaster = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            this.SelectByMaster(currentMasterNameU, comparison).ReplaceMaster(replacementStencil, resizeToMaster);
            return this;
        }

    }
}
