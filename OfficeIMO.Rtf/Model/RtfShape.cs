namespace OfficeIMO.Rtf;

/// <summary>
/// Dependency-free representation of an RTF drawing shape, including raw instructions, named properties, and optional text-box content.
/// </summary>
public sealed class RtfShape : IRtfInline, IRtfBlock {
    private readonly List<RtfShapeInstruction> _instructions = new List<RtfShapeInstruction>();
    private readonly List<RtfShapeProperty> _properties = new List<RtfShapeProperty>();
    private readonly List<RtfParagraph> _textBoxParagraphs = new List<RtfParagraph>();

    /// <summary>Raw shape instruction controls from <c>\shpinst</c>.</summary>
    public IReadOnlyList<RtfShapeInstruction> Instructions => _instructions.AsReadOnly();

    /// <summary>Named shape properties from <c>\sp</c> groups.</summary>
    public IReadOnlyList<RtfShapeProperty> Properties => _properties.AsReadOnly();

    /// <summary>Paragraphs contained by the shape text box destination.</summary>
    public IReadOnlyList<RtfParagraph> TextBoxParagraphs => _textBoxParagraphs.AsReadOnly();

    /// <summary>Adds a raw shape instruction control.</summary>
    public RtfShapeInstruction AddInstruction(string name, int? parameter = null, bool hasParameter = true) {
        var instruction = new RtfShapeInstruction(name, parameter, hasParameter);
        _instructions.Add(instruction);
        return instruction;
    }

    /// <summary>Adds a named shape property.</summary>
    public RtfShapeProperty AddProperty(string name, string? value = null) {
        var property = new RtfShapeProperty(name, value);
        _properties.Add(property);
        return property;
    }

    /// <summary>Adds a paragraph to the shape text box.</summary>
    public RtfParagraph AddTextBoxParagraph(string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) {
            paragraph.AddText(text!);
        }

        _textBoxParagraphs.Add(paragraph);
        return paragraph;
    }

    /// <summary>Returns text-box content with paragraphs separated by new lines.</summary>
    public string ToPlainText() {
        var builder = new StringBuilder();
        for (int i = 0; i < _textBoxParagraphs.Count; i++) {
            if (i > 0) {
                builder.AppendLine();
            }

            builder.Append(_textBoxParagraphs[i].ToPlainText());
        }

        return builder.ToString();
    }

    internal void AddParsedTextBoxParagraph(RtfParagraph paragraph) {
        _textBoxParagraphs.Add(paragraph ?? throw new ArgumentNullException(nameof(paragraph)));
    }
}
