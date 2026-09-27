using System.Globalization;
using Avalonia.Data.Converters;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Small presentation converters shared by Studio views.</summary>
public static class StudioConverters {
    /// <summary>True when the bound value's name equals the converter parameter (enum tools, modes).</summary>
    public static readonly IValueConverter NameEquals = new NameEqualsConverter();

    /// <summary>Maps a Studio command id to its icon geometry so palettes and tool cards share one vocabulary.</summary>
    public static readonly IValueConverter CommandIcon = new CommandIconConverter();

    internal static string CommandIconKey(string? id) => id switch {
        "Open" => "IconOpen",
        "Save" => "IconSave",
        "SaveAs" => "IconSaveCopy",
        "Undo" => "IconUndo",
        "Redo" => "IconRedo",
        "Read" => "IconDocument",
        "FocusReading" => "IconFocus",
        "Comment" => "IconNote",
        "Edit" => "IconText",
        "Pages" => "IconPages",
        "MovePages" => "IconMoveTo",
        "MovePagesUp" => "IconArrowUp",
        "MovePagesDown" => "IconArrowDown",
        "RotatePagesLeft" => "IconRotateLeft",
        "RotatePagesRight" => "IconRotateRight",
        "DuplicatePages" => "IconDuplicate",
        "ExtractPages" => "IconExport",
        "DeletePages" => "IconTrash",
        "InsertBlankPage" => "IconBlankPage",
        "Forms" => "IconRectangle",
        "Protect" => "IconProtect",
        "Redact" => "IconRedact",
        "Convert" => "IconConvert",
        "Jobs" => "IconJobs",
        "Ocr" => "IconOcr",
        "Assemble" => "IconMerge",
        "Export" => "IconExport",
        "Print" => "IconPrint",
        "Inspect" => "IconSearch",
        "Compare" => "IconCompare",
        "Optimize" => "IconCompress",
        "Repair" => "IconShieldCheck",
        "Sanitize" => "IconShieldCheck",
        "FitPage" => "IconFitPage",
        "ActualSize" => "IconFitPage",
        "FitWidth" => "IconFitWidth",
        "ZoomIn" => "IconPlus",
        "ZoomOut" => "IconMinus",
        "Home" => "IconHome",
        "Tools" => "IconTools",
        "Settings" => "IconSettings",
        _ => "IconMore"
    };

    private sealed class CommandIconConverter : IValueConverter {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
            Avalonia.Application.Current?.TryGetResource(CommandIconKey(value as string), null, out object? resource) == true
                ? resource
                : null;

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
            Avalonia.Data.BindingOperations.DoNothing;
    }

    /// <summary>True when a count is zero; drives empty-state hints.</summary>
    public static readonly IValueConverter IsZero = new FuncValueConverter<int, bool>(count => count == 0);

    /// <summary>Shows an engine code such as "render.resource.font-substitution" as plain words.</summary>
    /// <summary>Shows only the file or folder name of a location; bind the full location to a tooltip.</summary>
    public static readonly IValueConverter FileName = new FuncValueConverter<string?, string>(location =>
        string.IsNullOrWhiteSpace(location) ? string.Empty
            : Path.GetFileName(location.TrimEnd('\\', '/')) is { Length: > 0 } name ? name : location);

    public static readonly IValueConverter FriendlyCode = new FuncValueConverter<string?, string>(StudioMessages.Humanize);

    /// <summary>Turns an indent in pixels into a left margin, for nested outline rows.</summary>
    public static readonly IValueConverter LeftMargin = new FuncValueConverter<double, Avalonia.Thickness>(indent => new Avalonia.Thickness(indent, 0, 0, 0));

    /// <summary>Parses a #RRGGBB text into a brush preview; invalid text shows no fill.</summary>
    public static readonly IValueConverter HexBrush = new FuncValueConverter<string?, Avalonia.Media.IBrush?>(text =>
        Avalonia.Media.Color.TryParse(text, out Avalonia.Media.Color color) ? new Avalonia.Media.SolidColorBrush(color) : null);

    private sealed class NameEqualsConverter : IValueConverter {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
            value is not null && parameter is not null &&
            string.Equals(value.ToString(), parameter.ToString(), StringComparison.OrdinalIgnoreCase);

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
            Avalonia.Data.BindingOperations.DoNothing;
    }
}
