using Microsoft.Extensions.Configuration;

namespace OfficeIMO.Toml.Configuration;

/// <summary>A TOML file as an <see cref="IConfigurationSource"/>.</summary>
public class TomlConfigurationSource : FileConfigurationSource {
    /// <summary>Builds the <see cref="TomlConfigurationProvider"/> for this source.</summary>
    /// <param name="builder">The <see cref="IConfigurationBuilder"/>.</param>
    /// <returns>A <see cref="TomlConfigurationProvider"/>.</returns>
    public override IConfigurationProvider Build(IConfigurationBuilder builder) {
        EnsureDefaults(builder);
        return new TomlConfigurationProvider(this);
    }
}

/// <summary>
/// Reads a TOML file into configuration keys. Table keys, dotted keys and the indexes of arrays and arrays of tables are
/// joined with <c>:</c>; strings keep their decoded text, numbers and booleans their invariant spelling, and date-times
/// become ISO 8601.
/// </summary>
public class TomlConfigurationProvider : FileConfigurationProvider {
    /// <summary>Initializes a new instance with the specified source.</summary>
    /// <param name="source">The source settings.</param>
    public TomlConfigurationProvider(TomlConfigurationSource source) : base(source) { }

    /// <summary>Loads the TOML data from a stream.</summary>
    /// <param name="stream">The stream to read.</param>
    /// <exception cref="FormatException">The TOML is malformed or has keys that differ only in case.</exception>
    public override void Load(Stream stream) => Data = TomlConfigurationFileParser.Parse(stream);
}
