using Microsoft.Extensions.Configuration;

namespace OfficeIMO.Yaml.Configuration;

/// <summary>A YAML file as an <see cref="IConfigurationSource"/>.</summary>
public class YamlConfigurationSource : FileConfigurationSource {
    /// <summary>Builds the <see cref="YamlConfigurationProvider"/> for this source.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/>.</param>
    /// <returns>A <see cref="YamlConfigurationProvider"/>.</returns>
    public override IConfigurationProvider Build(IConfigurationBuilder builder) {
        EnsureDefaults(builder);
        return new YamlConfigurationProvider(this);
    }
}

/// <summary>
/// Reads a YAML file into configuration keys. Mapping keys and sequence indexes are joined with <c>:</c>, scalars keep their
/// decoded text, YAML nulls become empty strings, and aliases and <c>&lt;&lt;</c> merge keys are resolved.
/// </summary>
public class YamlConfigurationProvider : FileConfigurationProvider {
    /// <summary>Initializes a new instance with the specified source.</summary>
    /// <param name="source">The source settings.</param>
    public YamlConfigurationProvider(YamlConfigurationSource source) : base(source) { }

    /// <summary>Loads the YAML data from a stream.</summary>
    /// <param name="stream">The stream to read.</param>
    /// <exception cref="FormatException">The YAML is malformed, has duplicate keys, more than one document, or a root that is
    /// not a mapping.</exception>
    public override void Load(Stream stream) => Data = YamlConfigurationFileParser.Parse(stream);
}
