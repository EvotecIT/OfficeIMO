using Avalonia.Markup.Xaml;

namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Resolves a stable Studio localization key while Avalonia constructs a view.</summary>
internal sealed class LocExtension : MarkupExtension {
    public LocExtension(string key) {
        Key = key;
    }

    public string Key { get; }

    public override object ProvideValue(IServiceProvider serviceProvider) => StudioLocalization.Current.Get(Key);
}
