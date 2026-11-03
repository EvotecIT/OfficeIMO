using System;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
        internal static PowerPointSlide AddDesignerSlide(PowerPointDocument presentation,
            PowerPointDesignerSlideOptions options) {
            ArgumentNullException.ThrowIfNull(presentation);
            ArgumentNullException.ThrowIfNull(options);
            return options.TemplateLayout == null
                ? presentation.AddSlide()
                : presentation.AddSlide(options.TemplateLayout);
        }

        /// <summary>
        ///     Creates a designer facade whose semantic slide kinds use explicit named template layouts.
        ///     Theme tokens are imported into the design brief while the copied template remains the native owner.
        /// </summary>
        public static PowerPointDeckComposer UseTemplateDesigner(this PowerPointDocument presentation,
            PowerPointTemplateInventory inventory, PowerPointTemplateLayoutMap layoutMap, string seed,
            string? purpose = null, int alternativeIndex = 0, bool applyTheme = false) {
            ArgumentNullException.ThrowIfNull(presentation);
            ArgumentNullException.ThrowIfNull(inventory);
            ArgumentNullException.ThrowIfNull(layoutMap);
            PowerPointDesignBrief brief = inventory.CreateDesignBrief(seed, purpose);
            return new PowerPointDeckComposer(presentation, brief.CreateDesign(alternativeIndex), applyTheme,
                layoutMap);
        }
    }
}
