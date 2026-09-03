using System.Globalization;

namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Resolves culture-aware Studio presentation text by stable resource key.</summary>
internal interface IStudioLocalizer {
    CultureInfo Culture { get; }

    string Get(string key);

    string Format(string key, params object?[] arguments);
}
