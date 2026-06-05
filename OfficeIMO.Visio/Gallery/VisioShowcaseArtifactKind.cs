namespace OfficeIMO.Visio {
    /// <summary>
    /// Describes the role a generated file plays in a Visio showcase proof bundle.
    /// </summary>
    public enum VisioShowcaseArtifactKind {
        /// <summary>A generated VSDX package.</summary>
        Package,

        /// <summary>A dependency-free OfficeIMO-native SVG or PNG preview.</summary>
        NativePreview,

        /// <summary>A preview exported through Microsoft Visio desktop automation.</summary>
        DesktopPreview,

        /// <summary>A preview artifact whose proof lane is not classified more specifically.</summary>
        Preview,

        /// <summary>A deterministic inspection snapshot text artifact for a generated package.</summary>
        Inspection,

        /// <summary>A deterministic stencil/profile usage text artifact for a generated package.</summary>
        StencilProfile,

        /// <summary>A deterministic visual-quality analysis text artifact for a generated package.</summary>
        VisualQuality,

        /// <summary>A structural proof artifact whose proof lane is not classified more specifically.</summary>
        Proof
    }
}
