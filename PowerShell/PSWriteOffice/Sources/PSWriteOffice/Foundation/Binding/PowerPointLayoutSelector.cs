using System;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Foundation;

/// <summary>A template layout selected by literal name, one-based position, native type or returned layout metadata.</summary>
public sealed class PowerPointLayoutSelector {
    private PowerPointLayoutSelector(string? name = null, int? position = null,
        PowerPointSlideLayoutType? type = null, PowerPointSlideLayoutInfo? layout = null) {
        Name = name; Position = position; Type = type; Layout = layout;
    }
    /// <summary>Gets the literal layout name; numeric strings remain names.</summary>
    public string? Name { get; }
    /// <summary>Gets a one-based layout position.</summary>
    public int? Position { get; }
    /// <summary>Gets the native layout type.</summary>
    public PowerPointSlideLayoutType? Type { get; }
    /// <summary>Gets layout metadata returned by Get-OfficePowerPointLayout.</summary>
    public PowerPointSlideLayoutInfo? Layout { get; }

    /// <summary>Creates a selector from a PowerShell argument without loading or changing a document.</summary>
    public static PowerPointLayoutSelector From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw OfficeArgumentValue.Invalid("A template layout is required.");
        if (value is PowerPointLayoutSelector selector) return selector;
        if (value is string name) {
            if (string.IsNullOrWhiteSpace(name)) throw OfficeArgumentValue.Invalid("A layout name cannot be empty.");
            return new(name: name);
        }
        if (value is PowerPointSlideLayoutType type) return new(type: type);
        if (value is PowerPointSlideLayoutInfo layout) return new(layout: layout);
        if (value is PowerPointTemplateLayoutInfo template) return new(layout:
            new PowerPointSlideLayoutInfo(template.MasterIndex, template.LayoutIndex, template.Name, template.Type, null));
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try {
                int position = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                if (position >= 1) return new(position: position);
            } catch (OverflowException) { }
            throw OfficeArgumentValue.Invalid("Template layout positions start at 1 and must fit an Int32.");
        }
        throw OfficeArgumentValue.Invalid("Expected a layout name, one-based integer position, native layout type or layout metadata.");
    }

    internal PowerPointTemplateLayoutInfo Resolve(PowerPointPresentation document, int master, bool masterBound) {
        int masterIndex = Layout?.MasterIndex ?? master - 1;
        if (Layout.HasValue && masterBound && masterIndex != master - 1)
            throw new ArgumentException("Master does not match the supplied layout metadata.", nameof(master));
        var owner = PowerPointTemplate.Inspect(document).Masters.SingleOrDefault(item => item.MasterIndex == masterIndex)
            ?? throw new ArgumentOutOfRangeException(nameof(master), "The selected master does not exist in this document.");
        var layouts = owner.Layouts;
        if (Name != null) {
            var exact = layouts.Where(item => string.Equals(item.Name, Name, StringComparison.Ordinal)).ToArray();
            if (exact.Length == 1) return exact[0];
            var matches = exact.Length > 0 ? exact : layouts.Where(item => string.Equals(item.Name, Name, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 1) return matches[0];
            throw new ArgumentException(matches.Length == 0 ? "No layout has the supplied name." : "The layout name is ambiguous; select its position.");
        }
        var candidates = layouts.Where(item => Layout.HasValue
            ? item.LayoutIndex == Layout.Value.LayoutIndex && string.Equals(item.Name, Layout.Value.Name, StringComparison.Ordinal)
            : Position.HasValue ? item.LayoutIndex == Position.Value - 1 : item.Type == Type).ToArray();
        if (candidates.Length == 1) return candidates[0];
        throw new ArgumentException(candidates.Length == 0
            ? "The selected layout does not exist in this document."
            : "Several layouts have this type; select a layout name or position.");
    }
}

/// <summary>Binds template layout names, positions, native types and metadata.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class PowerPointLayoutTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => PowerPointLayoutSelector.From(inputData);
}
