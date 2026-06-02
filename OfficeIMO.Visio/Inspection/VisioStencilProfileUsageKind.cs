using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Classification for a stencil profile usage group.
    /// </summary>
    public enum VisioStencilProfileUsageKind {
        /// <summary>Shape uses a master imported from a stencil package or document package.</summary>
        PackageBackedMaster = 0,

        /// <summary>Shape uses a generated OfficeIMO master.</summary>
        GeneratedMaster = 1,

        /// <summary>Shape is direct geometry rather than a registered master instance.</summary>
        BasicGeometry = 2,

        /// <summary>Shape has no useful geometry or master identity and is grouped by semantic kind.</summary>
        SemanticOnly = 3
    }
}
