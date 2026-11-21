using Microsoft.Extensions.FileProviders;
using OfficeIMO.Yaml.Configuration;

namespace Microsoft.Extensions.Configuration;

/// <summary>Extension methods for adding YAML configuration sources.</summary>
public static class YamlConfigurationExtensions {
    /// <summary>Adds the YAML configuration provider at <paramref name="path"/> to <paramref name="builder"/>.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/> to add to.</param>
    /// <param name="path">Path relative to the base path stored in <see cref="IConfigurationBuilder.Properties"/> of
    /// <paramref name="builder"/>.</param>
    /// <returns>The <see cref="IConfigurationBuilder"/>.</returns>
    public static IConfigurationBuilder AddYamlFile(this IConfigurationBuilder builder, string path) =>
        AddYamlFile(builder, provider: null, path: path, optional: false, reloadOnChange: false);

    /// <summary>Adds the YAML configuration provider at <paramref name="path"/> to <paramref name="builder"/>.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/> to add to.</param>
    /// <param name="path">Path relative to the base path stored in <see cref="IConfigurationBuilder.Properties"/> of
    /// <paramref name="builder"/>.</param>
    /// <param name="optional">Whether the file is optional.</param>
    /// <returns>The <see cref="IConfigurationBuilder"/>.</returns>
    public static IConfigurationBuilder AddYamlFile(this IConfigurationBuilder builder, string path, bool optional) =>
        AddYamlFile(builder, provider: null, path: path, optional: optional, reloadOnChange: false);

    /// <summary>Adds the YAML configuration provider at <paramref name="path"/> to <paramref name="builder"/>.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/> to add to.</param>
    /// <param name="path">Path relative to the base path stored in <see cref="IConfigurationBuilder.Properties"/> of
    /// <paramref name="builder"/>.</param>
    /// <param name="optional">Whether the file is optional.</param>
    /// <param name="reloadOnChange">Whether the configuration should be reloaded if the file changes.</param>
    /// <returns>The <see cref="IConfigurationBuilder"/>.</returns>
    public static IConfigurationBuilder AddYamlFile(this IConfigurationBuilder builder, string path, bool optional, bool reloadOnChange) =>
        AddYamlFile(builder, provider: null, path: path, optional: optional, reloadOnChange: reloadOnChange);

    /// <summary>Adds a YAML configuration source to <paramref name="builder"/>.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/> to add to.</param>
    /// <param name="provider">The <see cref="IFileProvider"/> to use to access the file.</param>
    /// <param name="path">Path relative to the base path stored in <see cref="IConfigurationBuilder.Properties"/> of
    /// <paramref name="builder"/>.</param>
    /// <param name="optional">Whether the file is optional.</param>
    /// <param name="reloadOnChange">Whether the configuration should be reloaded if the file changes.</param>
    /// <returns>The <see cref="IConfigurationBuilder"/>.</returns>
    public static IConfigurationBuilder AddYamlFile(this IConfigurationBuilder builder, IFileProvider? provider, string path, bool optional, bool reloadOnChange) {
        ArgumentNullException.ThrowIfNull(builder);
        if (string.IsNullOrEmpty(path)) throw new ArgumentException("File path must be a non-empty string.", nameof(path));

        return builder.AddYamlFile(source => {
            source.FileProvider = provider;
            source.Path = path;
            source.Optional = optional;
            source.ReloadOnChange = reloadOnChange;
            source.ResolveFileProvider();
        });
    }

    /// <summary>Adds a YAML configuration source to <paramref name="builder"/>.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/> to add to.</param>
    /// <param name="configureSource">Configures the source.</param>
    /// <returns>The <see cref="IConfigurationBuilder"/>.</returns>
    public static IConfigurationBuilder AddYamlFile(this IConfigurationBuilder builder, Action<YamlConfigurationSource>? configureSource) =>
        builder.Add(configureSource);

    /// <summary>Adds a YAML configuration source that reads <paramref name="stream"/> to <paramref name="builder"/>.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/> to add to.</param>
    /// <param name="stream">The <see cref="Stream"/> to read the YAML configuration data from.</param>
    /// <returns>The <see cref="IConfigurationBuilder"/>.</returns>
    public static IConfigurationBuilder AddYamlStream(this IConfigurationBuilder builder, Stream stream) {
        ArgumentNullException.ThrowIfNull(builder);
        return builder.Add<YamlStreamConfigurationSource>(source => source.Stream = stream);
    }
}
