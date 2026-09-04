using System.Globalization;

namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Resolves culture-aware Studio presentation text by stable resource key.</summary>
internal interface IStudioLocalizer {
    CultureInfo Culture { get; }

    string Get(string key);

    string GetOrDefault(string key, string fallback);

    string Format(string key, params object?[] arguments);

    string FormatOrDefault(string key, string fallback, params object?[] arguments);
}
