using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint {
    /// <summary>Maps semantic deck-plan slide kinds to named template layouts.</summary>
    public sealed class PowerPointTemplateLayoutMap {
        private readonly Dictionary<PowerPointDeckPlanSlideKind, PowerPointTemplateLayoutInfo> _layouts = new();

        /// <summary>Configured semantic layout mappings.</summary>
        public IReadOnlyDictionary<PowerPointDeckPlanSlideKind, PowerPointTemplateLayoutInfo> Layouts =>
            new ReadOnlyDictionary<PowerPointDeckPlanSlideKind, PowerPointTemplateLayoutInfo>(_layouts);

        /// <summary>Maps a semantic kind to an inventoried layout.</summary>
        public PowerPointTemplateLayoutMap Map(PowerPointDeckPlanSlideKind kind,
            PowerPointTemplateLayoutInfo layout) {
            _layouts[kind] = layout ?? throw new ArgumentNullException(nameof(layout));
            return this;
        }

        /// <summary>Maps a semantic kind by resolving a template layout name.</summary>
        public PowerPointTemplateLayoutMap Map(PowerPointDeckPlanSlideKind kind,
            PowerPointTemplateInventory inventory, string layoutName) {
            if (inventory == null) throw new ArgumentNullException(nameof(inventory));
            return Map(kind, inventory.ResolveLayout(layoutName));
        }

        /// <summary>Returns the mapped layout, or null when the semantic kind uses the default layout.</summary>
        public PowerPointTemplateLayoutInfo? Resolve(PowerPointDeckPlanSlideKind kind) =>
            _layouts.TryGetValue(kind, out PowerPointTemplateLayoutInfo? layout) ? layout : null;
    }
}
