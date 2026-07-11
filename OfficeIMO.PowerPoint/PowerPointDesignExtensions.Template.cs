using System;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        internal static PowerPointSlide AddDesignerSlide(PowerPointPresentation presentation,
            PowerPointDesignerSlideOptions options) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return options.TemplateLayout == null
                ? presentation.AddSlide()
                : presentation.AddSlide(options.TemplateLayout);
        }

        /// <summary>
        ///     Creates a designer facade whose semantic slide kinds use explicit named template layouts.
        ///     Theme tokens are imported into the design brief while the copied template remains the native owner.
        /// </summary>
        public static PowerPointDeckComposer UseTemplateDesigner(this PowerPointPresentation presentation,
            PowerPointTemplateInventory inventory, PowerPointTemplateLayoutMap layoutMap, string seed,
            string? purpose = null, int alternativeIndex = 0, bool applyTheme = false) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (inventory == null) throw new ArgumentNullException(nameof(inventory));
            if (layoutMap == null) throw new ArgumentNullException(nameof(layoutMap));
            PowerPointDesignBrief brief = inventory.CreateDesignBrief(seed, purpose);
            return new PowerPointDeckComposer(presentation, brief.CreateDesign(alternativeIndex), applyTheme,
                layoutMap);
        }
    }
}
