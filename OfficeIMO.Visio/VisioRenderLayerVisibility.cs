using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio;

/// <summary>
/// Resolves each shape or connector's own layer memberships for one render operation.
/// Group children remain independent: callers must traverse them even when the group is hidden.
/// </summary>
internal sealed class VisioRenderLayerVisibility {
    private readonly Dictionary<string, bool> _layers = new(StringComparer.OrdinalIgnoreCase);
    private readonly VisioLayerRenderMode _mode;

    internal VisioRenderLayerVisibility(VisioPage page, VisioLayerRenderMode mode) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (!Enum.IsDefined(typeof(VisioLayerRenderMode), mode)) throw new ArgumentOutOfRangeException(nameof(mode));
        _mode = mode;

        foreach (VisioLayer layer in page.Layers) {
            bool enabled = mode == VisioLayerRenderMode.Printable ? layer.Print : layer.Visible;
            AddName(layer.Name, enabled);
            AddName(layer.NameU, enabled);
        }
    }

    /// <summary>Whether the shape's own geometry, artwork and text participate in this render.</summary>
    internal bool IsVisible(VisioShape shape) {
        if (shape == null) throw new ArgumentNullException(nameof(shape));
        return IsVisible(shape.LayerNames);
    }

    /// <summary>Whether the connector participates, independently of its endpoint shapes.</summary>
    internal bool IsVisible(VisioConnector connector) {
        if (connector == null) throw new ArgumentNullException(nameof(connector));
        return IsVisible(connector.LayerNames);
    }

    private void AddName(string? name, bool enabled) {
        // Preserve VisioPage.FindLayer's first matching display/universal name convention.
        if (!string.IsNullOrWhiteSpace(name) && !_layers.ContainsKey(name!)) _layers.Add(name!, enabled);
    }

    private bool IsVisible(ISet<string> names) {
        if (_mode == VisioLayerRenderMode.All || names.Count == 0) return true;
        foreach (string name in names) {
            // Missing layer declarations must not silently discard otherwise renderable content.
            if (string.IsNullOrWhiteSpace(name) || !_layers.TryGetValue(name, out bool enabled) || enabled) return true;
        }
        return false;
    }
}
