using System.Text;
using System.Text.Json;

namespace OfficeIMO.Notebook.Html;

internal sealed partial class NotebookHtmlWidgets {
    private const int MaxControlText = 4096;
    private const int MaxChoices = 128;
    private const double MaxNumericMagnitude = 1_000_000_000;

    private bool TryRenderControl(string id, string name, JsonElement state, WidgetRender render) {
        if (!TryOptionalText(state, "description", MaxControlText, out string description)
            || !TryBoolean(Property(state, "disabled"), false, out bool disabled)) return false;
        string label = description.Length == 0 ? DefaultLabel(name) : description;
        StringBuilder markup = render.Markup;
        if (name == "LabelModel") {
            if (!TryOptionalText(state, "value", MaxControlText, out string value)) return false;
            markup.Append("<span").Append(Bind(id, "text-value")).Append('>').Append(Encode(value)).Append("</span>");
            return Register(id, new WidgetValue("text", "value", Multiline: true), render);
        }
        if (name is "DropdownModel" or "SelectModel" or "RadioButtonsModel" or "SelectionSliderModel") {
            return TryRenderSelection(id, state, label, disabled, render);
        }
        if (name is "CheckboxModel" or "ToggleButtonModel") {
            if (!TryBoolean(Property(state, "value"), false, out bool value)) return false;
            StartLabel(markup, label, compact: true);
            markup.Append("<input type=\"checkbox\"").Append(Bind(id, "boolean"));
            if (value) markup.Append(" checked");
            AppendDisabled(markup, disabled);
            markup.Append("></label>");
            return Register(id, new WidgetValue("boolean", "value"), render);
        }
        if (name is "TextModel" or "TextareaModel") {
            return TryRenderText(id, name == "TextareaModel", state, label, disabled, render);
        }
        if (name is "IntSliderModel" or "FloatSliderModel" or "IntTextModel" or "FloatTextModel"
            or "BoundedIntTextModel" or "BoundedFloatTextModel" or "IntProgressModel" or "FloatProgressModel" or "PlayModel") {
            return TryRenderNumber(id, name, state, label, disabled, render);
        }
        return false;
    }

    private bool TryRenderText(string id, bool multiline, JsonElement state, string label, bool disabled, WidgetRender render) {
        if (!TryOptionalText(state, "value", MaxControlText, out string value)
            || !TryOptionalText(state, "placeholder", MaxControlText, out string placeholder)
            || !TryBoolean(Property(state, "continuous_update"), true, out bool continuous)
            || !multiline && (value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0)) return false;
        StringBuilder markup = render.Markup;
        StartLabel(markup, label);
        markup.Append(multiline ? "<textarea rows=\"4\"" : "<input type=\"text\"")
            .Append(Bind(id, "text", !continuous))
            .Append(" maxlength=\"4096\" placeholder=\"").Append(Encode(placeholder)).Append('"');
        AppendDisabled(markup, disabled);
        if (multiline) {
            // HTML removes the first LF immediately after a textarea start tag.
            markup.Append(">\n").Append(Encode(value)).Append("</textarea>");
        } else {
            markup.Append(" value=\"").Append(Encode(value)).Append("\">");
        }
        markup.Append("</label>");
        return Register(id, new WidgetValue("text", "value", Multiline: multiline), render);
    }

    private bool TryRenderSelection(string id, JsonElement state, string label, bool disabled, WidgetRender render) {
        JsonElement values = Property(state, "_options_labels");
        JsonElement selected = Property(state, "index");
        int index = -1;
        if (values.ValueKind != JsonValueKind.Array || values.GetArrayLength() > MaxChoices
            || (selected.ValueKind is not (JsonValueKind.Undefined or JsonValueKind.Null)
                && (!TryInt(selected, out index) || index < 0 || index >= values.GetArrayLength()))) return false;
        StringBuilder markup = render.Markup;
        StartLabel(markup, label);
        markup.Append("<select").Append(Bind(id, "index"));
        AppendDisabled(markup, disabled);
        markup.Append("><option value=\"-1\" hidden");
        if (index == -1) markup.Append(" selected");
        markup.Append("></option>");
        int current = 0;
        foreach (JsonElement value in values.EnumerateArray()) {
            if (render.ExceedsLimit || !TryText(value, MaxControlText, out string option)) return false;
            markup.Append("<option value=\"").Append(Number(current)).Append('"');
            if (index == current++) markup.Append(" selected");
            markup.Append('>').Append(Encode(option)).Append("</option>");
        }
        markup.Append("</select></label>");
        return Register(id, new WidgetValue("index", "index", -1, values.GetArrayLength() - 1), render);
    }

    private bool TryRenderNumber(string id, string name, JsonElement state, string label, bool disabled, WidgetRender render) {
        bool play = name == "PlayModel";
        bool slider = name is "IntSliderModel" or "FloatSliderModel";
        bool progress = name is "IntProgressModel" or "FloatProgressModel";
        bool bounded = play || slider || progress || name is "BoundedIntTextModel" or "BoundedFloatTextModel";
        bool integer = play || name.IndexOf("Int", StringComparison.Ordinal) >= 0;
        string kind = play ? "play" : progress ? "progress" : slider ? "range" : "number";
        double min = bounded ? 0 : -MaxNumericMagnitude;
        double max = bounded ? 100 : MaxNumericMagnitude;
        JsonElement savedStep = Property(state, "step");
        bool unrestrictedStep = !integer && !slider && !play && (!bounded || progress
            || savedStep.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null);
        bool vertical = false;
        if (slider) {
            JsonElement orientation = Property(state, "orientation");
            vertical = IsString(orientation, "vertical");
            if (!vertical && orientation.ValueKind != JsonValueKind.Undefined && !IsString(orientation, "horizontal")) return false;
        }
        if (!TryNumber(Property(state, "min"), min, integer, out min)
            || !TryNumber(Property(state, "max"), max, integer, out max)
            || !TryNumber(Property(state, "value"), 0, integer, out double value)
            || !TryNumber(savedStep, integer ? 1 : 0.1, integer, out double step, allowNull: !slider && !play)
            || !TryBoolean(Property(state, "continuous_update"), true, out bool continuous)
            || max < min || value < min || value > max || step <= 0 || progress && min == max) return false;
        double gridStep = progress ? (integer ? 1 : 0) : (unrestrictedStep ? 0 : step);
        if ((gridStep > 0 && !IsGridPoint(value, min, gridStep)) || (play && !IsGridPoint(max, min, step))) return false;

        StringBuilder markup = render.Markup;
        if (play) {
            return TryRenderPlay(id, state, label, disabled, render, min, max, value, step);
        }
        StartLabel(markup, label);
        if (progress) {
            markup.Append("<progress").Append(Bind(id, kind))
                .Append(" value=\"").Append(Number(value - min)).Append("\" max=\"").Append(Number(max - min))
                .Append("\" data-officeimo-widget-min=\"").Append(Number(min))
                .Append("\" data-officeimo-widget-value=\"").Append(Number(value)).Append("\"></progress>");
        } else {
            markup.Append("<input type=\"").Append(slider ? "range" : "number").Append('"').Append(Bind(id, kind, !continuous))
                .Append(" min=\"").Append(Number(min)).Append("\" max=\"").Append(Number(max))
                .Append("\" step=\"").Append(unrestrictedStep ? "any" : Number(step))
                .Append("\" value=\"").Append(Number(value)).Append('"');
            if (vertical) markup.Append(" style=\"writing-mode:vertical-lr;direction:rtl\"");
            AppendDisabled(markup, disabled);
            markup.Append('>');
        }
        markup.Append("</label>");
        return Register(id, new WidgetValue("numeric", "value", min, max, integer, gridStep), render);
    }

    private bool TryRenderPlay(string id, JsonElement state, string label, bool disabled, WidgetRender render,
        double min, double max, double value, double step) {
        if (!TryNumber(Property(state, "interval"), 100, true, out double interval) || interval is < 16 or > 60_000
            || !TryBoolean(Property(state, "_repeat"), false, out bool repeat)) return false;
        StringBuilder markup = render.Markup;
        markup.Append("<span style=\"display:inline-flex;align-items:center;gap:.5rem\"><button type=\"button\"")
            .Append(Bind(id, "play")).Append(" aria-label=\"").Append(Encode(label))
            .Append("\" aria-pressed=\"false\" value=\"").Append(Number(value))
            .Append("\" data-officeimo-widget-min=\"").Append(Number(min))
            .Append("\" data-officeimo-widget-max=\"").Append(Number(max))
            .Append("\" data-officeimo-widget-step=\"").Append(Number(step))
            .Append("\" data-officeimo-widget-interval=\"").Append(Number(interval))
            .Append("\" data-officeimo-widget-repeat=\"").Append(repeat ? "true" : "false").Append('"');
        AppendDisabled(markup, disabled || min == max);
        markup.Append(">Play</button><output>").Append(Number(value)).Append("</output></span>");
        render.NeedsRuntime = true;
        return Register(id, new WidgetValue("numeric", "value", min, max, Integer: true, Step: step), render);
    }

    private static bool Register(string id, WidgetValue value, WidgetRender render) {
        if (render.Values.ContainsKey(id)) render.NeedsRuntime = true;
        else render.Values.Add(id, value);
        return true;
    }

    private static void StartLabel(StringBuilder markup, string label, bool compact = false) =>
        markup.Append(compact
            ? "<label style=\"display:inline-flex;align-items:center;gap:.5rem\">"
            : "<label style=\"display:inline-grid;gap:.25rem;max-width:100%\">")
            .Append("<span>").Append(Encode(label)).Append("</span>");

    private static void AppendDisabled(StringBuilder markup, bool disabled) {
        if (disabled) markup.Append(" disabled");
    }

    private static string Bind(string id, string kind, bool commitOnly = false) => " data-officeimo-widget-id=\"" + Encode(id)
        + "\" data-officeimo-widget-kind=\"" + kind + "\""
        + (commitOnly ? " data-officeimo-widget-commit-only" : string.Empty);

    private static string DefaultLabel(string name) => name.EndsWith("Model", StringComparison.Ordinal)
        ? name.Substring(0, name.Length - "Model".Length) : name;

    private static bool TryOptionalText(JsonElement item, string property, int maxLength, out string text) {
        JsonElement value = Property(item, property);
        text = string.Empty;
        return value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null || TryText(value, maxLength, out text);
    }

    private static bool TryBoolean(JsonElement property, bool fallback, out bool value) {
        value = fallback;
        if (property.ValueKind == JsonValueKind.Undefined) return true;
        if (property.ValueKind is not (JsonValueKind.True or JsonValueKind.False)) return false;
        value = property.GetBoolean();
        return true;
    }

    private static bool TryNumber(JsonElement property, double fallback, bool integer, out double value, bool allowNull = false) {
        value = fallback;
        if (property.ValueKind == JsonValueKind.Undefined || allowNull && property.ValueKind == JsonValueKind.Null) return true;
        return property.ValueKind == JsonValueKind.Number && property.TryGetDouble(out value)
            && value is >= -MaxNumericMagnitude and <= MaxNumericMagnitude && (!integer || Math.Truncate(value) == value);
    }

    private readonly record struct WidgetValue(
        string Kind, string Trait, double Min = 0, double Max = 0, bool Integer = false, double Step = 0,
        bool Multiline = false);
}
