using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using OfficeIMO.Markdown;

namespace OfficeIMO.Notebook.Html;

/// <summary>Projects approved saved widget models into browser controls without loading notebook code.</summary>
internal sealed partial class NotebookHtmlWidgets {
    private const string StateMime = "application/vnd.jupyter.widget-state+json";
    private const string ViewMime = "application/vnd.jupyter.widget-view+json";
    private const string ControlsModule = "@jupyter-widgets/controls";
    private const int MaxModels = 2048;
    private const int MaxViews = 1024;
    private const int MaxNodesPerView = 64;
    private const int MaxDepth = 12;
    private const int MaxModelIdLength = 128;
    private const int MaxWidgetCharacters = 2 * 1024 * 1024;
    private readonly NotebookDocument _notebook;
    private readonly CancellationToken _cancellationToken;
    private readonly List<OfficeDiagnostic> _diagnostics = new();
    private readonly Dictionary<string, JsonElement> _models = new(StringComparer.Ordinal);
    private readonly Dictionary<string, WidgetValue> _values = new(StringComparer.Ordinal);
    private bool _readState;
    private bool _renderedStateIsUsable;
    private bool _needsRuntime;
    private int _views;
    private int _characters;

    internal NotebookHtmlWidgets(NotebookDocument notebook, CancellationToken cancellationToken) {
        _notebook = notebook;
        _cancellationToken = cancellationToken;
    }

    internal IReadOnlyList<OfficeDiagnostic> Diagnostics => _diagnostics;

    internal IMarkdownBlock? RenderOutput(NotebookOutput output, string location) {
        if (!output.Data.TryGetValue(ViewMime, out JsonElement view)) return null;
        _cancellationToken.ThrowIfCancellationRequested();
        if (!HasSupportedVersion(view, maximumMinor: 1) || !TryText(Property(view, "model_id"), MaxModelIdLength, out string id)
            || id.Length == 0 || !ReadState() || !_models.ContainsKey(id)) {
            Report("notebook-widget-state-unavailable", location, "The saved widget view has no matching supported widget state.");
            return null;
        }
        if (_views >= MaxViews || _characters >= MaxWidgetCharacters) {
            Report("notebook-widget-view-limit", location, "The saved widget view exceeds the HTML conversion budget.");
            return null;
        }
        var viewRender = new WidgetRender();
        if (!TryRenderModel(id, viewRender, 0)) {
            Report("notebook-widget-view-unsupported", location,
                "The saved widget uses a model, value, or layout outside the supported browser controls.");
            return null;
        }
        if (viewRender.Markup.Length > MaxWidgetCharacters - _characters) {
            Report("notebook-widget-view-limit", location, "The saved widget view exceeds the HTML conversion budget.");
            return null;
        }

        _views++;
        _characters += viewRender.Markup.Length;
        _needsRuntime |= viewRender.NeedsRuntime;
        foreach (var pair in viewRender.Values) {
            if (_values.ContainsKey(pair.Key)) _needsRuntime = true;
            else _values.Add(pair.Key, pair.Value);
        }
        if (viewRender.ApproximatesAppearance) {
            _diagnostics.Add(new OfficeDiagnostic("notebook-widget-appearance", OfficeDiagnosticSeverity.Warning,
                "The browser controls approximate the saved widget layout or appearance.", location,
                "Jupyter", OfficeConversionLossKind.Approximation));
        }
        return new NotebookGeneratedHtmlBlock("<div class=\"officeimo-widget\">" + viewRender.Markup.ToString() + "</div>");
    }

    private bool ReadState() {
        if (_readState) return _renderedStateIsUsable;
        _readState = true;
        if (!_notebook.Metadata.TryGetValue("widgets", out JsonElement metadata)) return false;
        JsonElement envelope = Property(metadata, StateMime);
        JsonElement state = Property(envelope, "state");
        if (!HasSupportedVersion(envelope) || state.ValueKind != JsonValueKind.Object) return false;
        foreach (JsonProperty model in state.EnumerateObject()) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_models.Count >= MaxModels || !TryReadModelKey(model, out string name) || _models.ContainsKey(name)) return false;
            _models.Add(name, model.Value);
        }
        return _renderedStateIsUsable = true;
    }

    private bool TryRenderModel(string id, WidgetRender render, int depth) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (depth > MaxDepth || ++render.Nodes > MaxNodesPerView || !render.Path.Add(id)) return false;
        bool result = TryReadModel(id, out string name, out JsonElement state)
            && (name is "BoxModel" or "HBoxModel" or "VBoxModel"
                ? TryRenderBox(name, state, render, depth) : TryRenderControl(id, name, state, render));
        if (result && !render.ApproximatesAppearance) {
            render.ApproximatesAppearance = HasCustomAppearance(state, "layout") || HasCustomAppearance(state, "style")
                || name is "ToggleButtonModel" or "RadioButtonsModel" or "SelectionSliderModel";
        }
        render.Path.Remove(id);
        return result;
    }

    private bool TryReadModel(string id, out string name, out JsonElement state) {
        name = string.Empty;
        state = default;
        if (!_models.TryGetValue(id, out JsonElement model)
            || !IsString(Property(model, "model_module"), ControlsModule)
            || !TryText(Property(model, "model_module_version"), 24, out string version)
            || !IsSupportedControlsVersion(version)
            || !TryText(Property(model, "model_name"), 48, out name)) return false;
        state = Property(model, "state");
        JsonElement viewModule = Property(state, "_view_module");
        return state.ValueKind == JsonValueKind.Object
            && Property(state, "_esm").ValueKind == JsonValueKind.Undefined
            && (viewModule.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null || IsString(viewModule, ControlsModule));
    }

    private bool HasCustomAppearance(JsonElement state, string property) =>
        TryReference(Property(state, property), out string id) && _models.TryGetValue(id, out JsonElement model)
        && Property(model, "state") is { ValueKind: JsonValueKind.Object } appearance
        && appearance.EnumerateObject().Any(item => item.Value.ValueKind != JsonValueKind.Null
            && !IsString(item.Value, string.Empty)
            && !item.NameEquals("_model_module") && !item.NameEquals("_model_module_version") && !item.NameEquals("_model_name")
            && !item.NameEquals("_view_module") && !item.NameEquals("_view_module_version") && !item.NameEquals("_view_name")
            && !item.NameEquals("_view_count"));

    private static bool TryReadModelKey(JsonProperty model, out string name) {
        name = string.Empty;
        try {
            name = model.Name;
            return name.Length is > 0 and <= MaxModelIdLength;
        } catch (InvalidOperationException) {
            return false;
        }
    }

    private bool TryRenderBox(string name, JsonElement state, WidgetRender render, int depth) {
        JsonElement children = Property(state, "children");
        if (children.ValueKind != JsonValueKind.Array || children.GetArrayLength() > MaxNodesPerView) return false;
        render.Markup.Append(name == "VBoxModel"
            ? "<div style=\"display:flex;flex-direction:column;gap:.5rem\">"
            : "<div style=\"display:flex;flex-wrap:wrap;gap:.75rem\">");
        foreach (JsonElement child in children.EnumerateArray()) {
            if (!TryReference(child, out string childId) || !TryRenderModel(childId, render, depth + 1)) return false;
        }
        render.Markup.Append("</div>");
        return true;
    }

    private static bool HasSupportedVersion(JsonElement item, int maximumMinor = 0) =>
        TryInt(Property(item, "version_major"), out int major) && major == 2
        && TryInt(Property(item, "version_minor"), out int minor) && minor >= 0 && minor <= maximumMinor;

    private static bool IsSupportedControlsVersion(string version) {
        if (version.Length > 0 && version[0] is '^' or '~') version = version.Substring(1);
        return Version.TryParse(version, out Version? parsed) && parsed.Major is 1 or 2
            && parsed.Build >= 0 && parsed.Revision < 0;
    }

    private static bool TryReference(JsonElement item, out string id) {
        const string prefix = "IPY_MODEL_";
        id = string.Empty;
        if (!TryText(item, prefix.Length + MaxModelIdLength, out string reference)
            || !reference.StartsWith(prefix, StringComparison.Ordinal) || reference.Length == prefix.Length) return false;
        id = reference.Substring(prefix.Length);
        return true;
    }

    private static JsonElement Property(JsonElement value, string name) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out JsonElement property) ? property : default;

    private static bool IsString(JsonElement value, string expected) =>
        value.ValueKind == JsonValueKind.String && value.ValueEquals(expected);

    private static bool TryInt(JsonElement value, out int number) {
        number = 0;
        return value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out number);
    }

    private static bool TryText(JsonElement value, int maxLength, out string text) {
        text = string.Empty;
        if (value.ValueKind != JsonValueKind.String) return false;
        try {
            string candidate = value.GetString() ?? string.Empty;
            if (candidate.Length > maxLength || candidate.Any(character =>
                character is '\ufffe' or '\uffff' || char.IsControl(character) && character is not ('\t' or '\n' or '\r'))) return false;
            text = candidate;
            return true;
        } catch (InvalidOperationException) {
            return false;
        }
    }

    private static string Encode(string value) => WebUtility.HtmlEncode(value);

    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    private void Report(string code, string location, string message) =>
        _diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Warning, message, location, "Jupyter", OfficeConversionLossKind.Omission));

    private sealed class WidgetRender {
        internal StringBuilder Markup { get; } = new();
        internal Dictionary<string, WidgetValue> Values { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Path { get; } = new(StringComparer.Ordinal);
        internal int Nodes { get; set; }
        internal bool NeedsRuntime { get; set; }
        internal bool ApproximatesAppearance { get; set; }
    }

}
