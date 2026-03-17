using OfficeIMO.Markdown;
using System.Runtime.CompilerServices;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Bridges renderer plugins and feature packs into source-side markdown reader configuration.
/// </summary>
public static class MarkdownReaderOptionsExtensions {
    private sealed class ReaderContractState {
        public HashSet<string> PluginIds { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> FeaturePackIds { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private static readonly ConditionalWeakTable<MarkdownReaderOptions, ReaderContractState> ReaderStates = new();

    /// <summary>
    /// Applies a renderer plugin's reader-side contract to markdown reader options.
    /// </summary>
    public static void ApplyPlugin(this MarkdownReaderOptions options, MarkdownRendererPlugin plugin) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (plugin == null) {
            throw new ArgumentNullException(nameof(plugin));
        }

        var state = ReaderStates.GetOrCreateValue(options);
        if (!state.PluginIds.Add(plugin.Id)) {
            return;
        }

        plugin.ApplyReader(options);
    }

    /// <summary>
    /// Returns <see langword="true"/> when a renderer plugin's reader-side contract has already been applied.
    /// </summary>
    public static bool HasPlugin(this MarkdownReaderOptions options, MarkdownRendererPlugin plugin) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (plugin == null) {
            throw new ArgumentNullException(nameof(plugin));
        }

        var state = ReaderStates.GetOrCreateValue(options);
        return state.PluginIds.Contains(plugin.Id) || plugin.ReaderContractMatches(options);
    }

    /// <summary>
    /// Applies a renderer feature pack's reader-side contract to markdown reader options.
    /// </summary>
    public static void ApplyFeaturePack(this MarkdownReaderOptions options, MarkdownRendererFeaturePack featurePack) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (featurePack == null) {
            throw new ArgumentNullException(nameof(featurePack));
        }

        var state = ReaderStates.GetOrCreateValue(options);
        if (!state.FeaturePackIds.Add(featurePack.Id)) {
            return;
        }

        featurePack.ApplyReader(options);
    }

    /// <summary>
    /// Returns <see langword="true"/> when a renderer feature pack's reader-side contract has already been applied.
    /// </summary>
    public static bool HasFeaturePack(this MarkdownReaderOptions options, MarkdownRendererFeaturePack featurePack) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (featurePack == null) {
            throw new ArgumentNullException(nameof(featurePack));
        }

        var state = ReaderStates.GetOrCreateValue(options);
        return state.FeaturePackIds.Contains(featurePack.Id) || featurePack.ReaderContractMatches(options);
    }
}
