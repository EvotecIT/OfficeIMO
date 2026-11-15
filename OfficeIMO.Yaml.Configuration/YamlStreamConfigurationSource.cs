using Microsoft.Extensions.Configuration;

namespace OfficeIMO.Yaml.Configuration;

/// <summary>A YAML stream as an <see cref="IConfigurationSource"/>.</summary>
public class YamlStreamConfigurationSource : StreamConfigurationSource {
    /// <summary>Builds the <see cref="YamlStreamConfigurationProvider"/> for this source.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/>.</param>
    /// <returns>A <see cref="YamlStreamConfigurationProvider"/>.</returns>
    public override IConfigurationProvider Build(IConfigurationBuilder builder) => new YamlStreamConfigurationProvider(this);
}

/// <summary>Reads YAML from a stream into configuration keys, with the rules of <see cref="YamlConfigurationProvider"/>.</summary>
public class YamlStreamConfigurationProvider : StreamConfigurationProvider {
    /// <summary>Initializes a new instance with the specified source.</summary>
    /// <param name="source">The <see cref="YamlStreamConfigurationSource"/>.</param>
    public YamlStreamConfigurationProvider(YamlStreamConfigurationSource source) : base(source) { }

    /// <summary>Loads the YAML data from a stream.</summary>
    /// <param name="stream">The stream to read.</param>
    /// <exception cref="FormatException">The YAML is malformed, has duplicate keys, more than one document, or a root that is
    /// not a mapping.</exception>
    public override void Load(Stream stream) => Data = YamlConfigurationFileParser.Parse(stream);
}
