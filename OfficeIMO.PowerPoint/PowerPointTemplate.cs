using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Optional adapter for inspecting a PowerPoint template and creating an editable presentation from it.
    /// </summary>
    public static class PowerPointTemplate {
        /// <summary>Inspects masters, layouts, placeholders, theme tokens, and reusable assets in a template file.</summary>
        public static PowerPointTemplateInventory Inspect(string templatePath) {
            return PowerPointPresentation.InspectTemplate(templatePath);
        }

        /// <summary>Inspects the template-capable structure already loaded by a presentation.</summary>
        public static PowerPointTemplateInventory Inspect(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return presentation.InspectTemplate();
        }

        /// <summary>
        ///     Copies a `.pptx` or `.potx` template into a new editable presentation while preserving native
        ///     masters, layouts, themes, relationships, and assets.
        /// </summary>
        public static PowerPointPresentation CreatePresentation(string templatePath, string outputPath,
            PowerPointTemplateCreationOptions? options = null) {
            return PowerPointPresentation.CreateFromTemplate(templatePath, outputPath, options);
        }
    }
}
