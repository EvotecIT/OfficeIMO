using Microsoft.Extensions.Configuration;

namespace OfficeIMO.Toml.Configuration;

/// <summary>A TOML stream as an <see cref="IConfigurationSource"/>.</summary>
public class TomlStreamConfigurationSource : StreamConfigurationSource {
    /// <summary>Builds the <see cref="TomlStreamConfigurationProvider"/> for this source.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/>.</param>
    /// <returns>A <see cref="TomlStreamConfigurationProvider"/>.</returns>
    public override IConfigurationProvider Build(IConfigurationBuilder builder) => new TomlStreamConfigurationProvider(this);
}

/// <summary>Reads TOML from a stream into configuration keys, with the rules of <see cref="TomlConfigurationProvider"/>.</summary>
public class TomlStreamConfigurationProvider : StreamConfigurationProvider {
    /// <summary>Initializes a new instance with the specified source.</summary>
    /// <param name="source">The <see cref="TomlStreamConfigurationSource"/>.</param>
    public TomlStreamConfigurationProvider(TomlStreamConfigurationSource source) : base(source) { }

    /// <summary>Loads the TOML data from a stream.</summary>
    /// <param name="stream">The stream to read.</param>
    /// <exception cref="FormatException">The TOML is malformed or has keys that differ only in case.</exception>
    public override void Load(Stream stream) => Data = TomlConfigurationFileParser.Parse(stream);
}
