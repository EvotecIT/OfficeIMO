using System.Collections;
using System.Globalization;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioTranslationResourceTests {
    [Fact]
    public void AvailableTranslationsContainEveryUiResourceAndPreserveFormatContracts() {
        var manager = new ResourceManager("OfficeIMO.Studio.Localization.StudioStrings", typeof(StudioLocalizer).Assembly);
        ResourceSet neutral = manager.GetResourceSet(CultureInfo.InvariantCulture, true, false)!;
        var source = neutral.Cast<DictionaryEntry>().ToDictionary(e => (string)e.Key, e => (string)e.Value!);
        foreach (var descriptor in StudioCultureCatalog.Available.Where(c => c.Name is not "en" and not StudioCultureCatalog.PseudoCulture)) {
            ResourceSet? translated = manager.GetResourceSet(CultureInfo.GetCultureInfo(descriptor.Name), true, false);
            Assert.NotNull(translated);
            var values = translated.Cast<DictionaryEntry>().ToDictionary(e => (string)e.Key, e => (string)e.Value!);
            Assert.Equal(source.Keys.Order(), values.Keys.Order());
            foreach ((string key, string english) in source) {
                string value = values[key];
                Assert.False(string.IsNullOrWhiteSpace(value), descriptor.Name + ": " + key);
                Assert.Equal(Tokens(english), Tokens(value));
                if (Regex.IsMatch(english, @"\{\d+[,:}]")) {
                    var format = CompositeFormat.Parse(value);
                    Assert.Equal(CompositeFormat.Parse(english).MinimumArgumentCount, format.MinimumArgumentCount);
                    object[] arguments = Enumerable.Repeat<object>(2, format.MinimumArgumentCount).ToArray();
                    _ = string.Format(CultureInfo.GetCultureInfo(descriptor.Name), format, arguments);
                }
            }
        }
    }

    private static string[] Tokens(string value) => Regex.Matches(value, @"(?<!\{)\{[^{}]+\}(?!\})")
        .Select(match => match.Value).Order().ToArray();
}
