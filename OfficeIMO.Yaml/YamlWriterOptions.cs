namespace OfficeIMO.Yaml;

/// <summary>
/// Layout settings for <see cref="YamlWriter"/>. Indentation follows the three numbers ruamel.yaml made familiar:
/// <see cref="MappingIndent"/>, <see cref="SequenceIndent"/> and <see cref="SequenceDashOffset"/>.
/// </summary>
/// <remarks>
/// <para>The defaults write <c>key:\n- item</c>, the compact form most real files use (mapping 2, sequence 2, offset 0).
/// Kubernetes manifests and Helm charts usually indent sequences under their key (mapping 2, sequence 4, offset 2), which
/// writes <c>key:\n  - item</c>.</para>
/// <para>Output is deterministic: the same calls with the same options always produce the same text. The writer never
/// folds long lines and always ends the stream with a line break.</para>
/// </remarks>
public sealed class YamlWriterOptions {
    private int _mappingIndent = 2;
    private int _sequenceIndent = 2;
    private int _sequenceDashOffset;
    private string _newLine = "\n";

    /// <summary>Gets the shared default options. The instance must not be modified.</summary>
    internal static YamlWriterOptions Default { get; } = new YamlWriterOptions();

    /// <summary>Gets or sets how far the keys of a nested mapping are indented from their parent key. Defaults to 2; 1 to 9.</summary>
    public int MappingIndent {
        get => _mappingIndent;
        set => _mappingIndent = CheckRange(value, 1, 9, nameof(MappingIndent));
    }

    /// <summary>
    /// Gets or sets how far the content of a sequence item under a mapping key is indented from that key. Defaults to 2;
    /// it must be at least <see cref="SequenceDashOffset"/> + 2 so the dash and a space fit before the content.
    /// </summary>
    public int SequenceIndent {
        get => _sequenceIndent;
        set => _sequenceIndent = CheckRange(value, 2, 10, nameof(SequenceIndent));
    }

    /// <summary>
    /// Gets or sets how far the dash of a sequence item under a mapping key is indented from that key. Defaults to 0,
    /// which writes <c>key:\n- item</c>; 2 with a sequence indent of 4 writes <c>key:\n  - item</c>.
    /// </summary>
    public int SequenceDashOffset {
        get => _sequenceDashOffset;
        set => _sequenceDashOffset = CheckRange(value, 0, 8, nameof(SequenceDashOffset));
    }

    /// <summary>Gets or sets the line break, <c>"\n"</c> (the default) or <c>"\r\n"</c>.</summary>
    public string NewLine {
        get => _newLine;
        set {
            if (value != "\n" && value != "\r\n") throw new ArgumentException("The line break must be \"\\n\" or \"\\r\\n\".", nameof(value));
            _newLine = value;
        }
    }

    /// <summary>The spaces between a dash and its item's content, at least one.</summary>
    internal int ItemPadding => Math.Max(1, _sequenceIndent - _sequenceDashOffset - 1);

    internal void Validate() {
        if (_sequenceIndent < _sequenceDashOffset + 2) {
            throw new InvalidOperationException("SequenceIndent (" + _sequenceIndent + ") must be at least SequenceDashOffset (" + _sequenceDashOffset + ") + 2.");
        }
    }

    private static int CheckRange(int value, int min, int max, string name) {
        if (value < min || value > max) throw new ArgumentOutOfRangeException(name, value, name + " must be between " + min + " and " + max + ".");
        return value;
    }
}
