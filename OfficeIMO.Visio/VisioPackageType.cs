namespace OfficeIMO.Visio {
    /// <summary>Identifies an Open XML Visio package family.</summary>
    public enum VisioPackageType {
        /// <summary>Macro-free drawing package (<c>.vsdx</c>).</summary>
        Drawing,

        /// <summary>Macro-free template package (<c>.vstx</c>).</summary>
        Template,

        /// <summary>Macro-free stencil package (<c>.vssx</c>).</summary>
        Stencil,

        /// <summary>Macro-enabled drawing package (<c>.vsdm</c>).</summary>
        MacroEnabledDrawing,

        /// <summary>Macro-enabled template package (<c>.vstm</c>).</summary>
        MacroEnabledTemplate,

        /// <summary>Macro-enabled stencil package (<c>.vssm</c>).</summary>
        MacroEnabledStencil
    }
}
