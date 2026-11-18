using System.Text;

namespace OfficeIMO.CSV;

/// <summary>
/// Rewrites RFC 4180 CSV text, as every CSV writer produces it, into fields enclosed by another quote character
/// and escaped with another escape character. Quoted fields are re-enclosed; unquoted fields that contain the new
/// quote character are enclosed too, so the output reads back with the same characters. The writers keep their
/// one quoting implementation and this sits between them and the destination.
/// </summary>
internal sealed class CsvQuotingTextWriter : TextWriter
{
    private const int OutputFlushThreshold = 16 * 1024;

    private enum State : byte
    {
        FieldStart,
        Unquoted,
        Quoted,
        QuoteSeen,
        Separator
    }

    private readonly TextWriter _inner;
    private readonly bool _leaveOpen;
    private readonly char _quote;
    private readonly char _escape;
    private readonly string _delimiter;
    private readonly string _recordSeparator;
    private readonly StringBuilder _field = new();
    private readonly StringBuilder _separator = new();
    private readonly StringBuilder _output = new();
    private State _state;
    private bool _completed;

    private CsvQuotingTextWriter(TextWriter inner, CsvSaveOptions options, bool leaveOpen)
        : base(inner.FormatProvider)
    {
        _inner = inner;
        _leaveOpen = leaveOpen;
        _quote = options.QuoteCharacter;
        _escape = options.EscapeCharacter ?? options.QuoteCharacter;
        _delimiter = CsvWriter.GetDelimiterText(options);
        _recordSeparator = options.NewLine;
        ValidateDialect(options);
    }

    /// <summary>
    /// Rejects quote and escape characters that cannot be written unambiguously. Path writers call it before they
    /// open the destination, so a rejected dialect never truncates an existing file.
    /// </summary>
    internal static void ValidateDialect(CsvSaveOptions options)
    {
        ValidateEncodableDialect(options, options.Encoding);
        if (!options.UsesCustomQuoting)
        {
            return;
        }

        char quote = options.QuoteCharacter;
        char escape = options.EscapeCharacter ?? quote;
        string delimiter = CsvWriter.GetDelimiterText(options);
        if (quote is '\r' or '\n' || escape is '\r' or '\n' ||
            delimiter.IndexOf(quote) >= 0 || delimiter.IndexOf(escape) >= 0)
        {
            throw new ArgumentException("Quote and escape characters cannot be line breaks or part of the delimiter.");
        }
    }

    /// <summary>
    /// Rejects a delimiter, quote or escape character that the output encoding cannot write. An unpaired surrogate,
    /// or a character outside a legacy code page, is replaced on output by the same substitute that replaces such
    /// characters in field text, so a reader can no longer tell the separators from the data.
    /// </summary>
    /// <param name="options">Save options that define the dialect.</param>
    /// <param name="encoding">Output encoding; null means the options' encoding or UTF-8.</param>
    internal static void ValidateEncodableDialect(CsvSaveOptions options, Encoding? encoding)
    {
        encoding ??= options.Encoding ?? Encoding.UTF8;
        ThrowIfNotEncodable(CsvWriter.GetDelimiterText(options), encoding, "delimiter");
        ThrowIfNotEncodable(options.QuoteCharacter.ToString(), encoding, "quote character");
        if (options.EscapeCharacter is char escape)
        {
            ThrowIfNotEncodable(escape.ToString(), encoding, "escape character");
        }
    }

    private static void ThrowIfNotEncodable(string text, Encoding encoding, string role)
    {
        var strict = (Encoding)encoding.Clone();
        strict.EncoderFallback = EncoderFallback.ExceptionFallback;
        try
        {
            strict.GetByteCount(text);
        }
        catch (EncoderFallbackException)
        {
            throw new ArgumentException(
                $"The CSV {role} cannot be written in the {encoding.WebName} output encoding; it would be indistinguishable from replaced field text.");
        }
    }

    /// <summary>Returns a writer that applies the configured quote and escape characters, or the writer itself when they are the defaults.</summary>
    internal static TextWriter Wrap(TextWriter writer, CsvSaveOptions options, bool leaveOpen) =>
        options.UsesCustomQuoting ? new CsvQuotingTextWriter(writer, options, leaveOpen) : writer;

    /// <summary>Rewrites one complete RFC 4180 record, without its line break, with the configured quote and escape characters.</summary>
    internal static string RewriteRecord(string record, CsvSaveOptions options)
    {
        if (!options.UsesCustomQuoting)
        {
            return record;
        }

        using var target = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
        using (var quoting = new CsvQuotingTextWriter(target, options, leaveOpen: true))
        {
            quoting.Write(record);
        }

        return target.ToString();
    }

    public override Encoding Encoding => _inner.Encoding;

    public override void Write(char value)
    {
        Process(value);
        FlushOutputWhenLarge();
    }

    public override void Write(char[] buffer, int index, int count)
    {
        for (int position = index; position < index + count; position++)
        {
            Process(buffer[position]);
        }

        FlushOutputWhenLarge();
    }

    public override void Write(string? value)
    {
        if (value is null)
        {
            return;
        }

        foreach (char character in value)
        {
            Process(character);
        }

        FlushOutputWhenLarge();
    }

#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<char> buffer)
    {
        foreach (char character in buffer)
        {
            Process(character);
        }

        FlushOutputWhenLarge();
    }

    public override void Write(StringBuilder? value)
    {
        if (value is null)
        {
            return;
        }

        foreach (ReadOnlyMemory<char> chunk in value.GetChunks())
        {
            Write(chunk.Span);
        }
    }
#endif

    public override void Flush()
    {
        WriteOutput();
        _inner.Flush();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && !_completed)
        {
            _completed = true;
            try
            {
                Complete();
                if (_leaveOpen)
                {
                    _inner.Flush();
                }
            }
            finally
            {
                if (!_leaveOpen)
                {
                    _inner.Dispose();
                }
            }
        }

        base.Dispose(disposing);
    }

    /// <summary>Writes a field still open at the end of the output, such as a last record without a line break.</summary>
    internal void Complete()
    {
        switch (_state)
        {
            case State.Unquoted:
                EmitField(enclose: false);
                break;
            case State.Quoted:
            case State.QuoteSeen:
                EmitField(enclose: true);
                break;
        }

        _state = State.FieldStart;
        WriteOutput();
    }

    private void Process(char character)
    {
        switch (_state)
        {
            case State.FieldStart when character == '"':
                _state = State.Quoted;
                return;
            case State.FieldStart:
            case State.Unquoted:
                _state = State.Unquoted;
                _field.Append(character);
                if (TryEndUnquotedField(_delimiter) || TryEndUnquotedField(_recordSeparator))
                {
                    _state = State.FieldStart;
                }

                return;
            case State.Quoted:
                if (character == '"')
                {
                    _state = State.QuoteSeen;
                }
                else
                {
                    _field.Append(character);
                }

                return;
            case State.QuoteSeen:
                if (character == '"')
                {
                    _field.Append('"');
                    _state = State.Quoted;
                    return;
                }

                EmitField(enclose: true);
                _state = State.Separator;
                _separator.Clear();
                AppendSeparator(character);
                return;
            default:
                AppendSeparator(character);
                return;
        }
    }

    /// <summary>Copies the delimiter or record separator that follows a quoted field.</summary>
    private void AppendSeparator(char character)
    {
        _output.Append(character);
        _separator.Append(character);
        if (EndsWith(_separator, _delimiter) || EndsWith(_separator, _recordSeparator) ||
            _separator.Length >= Math.Max(_delimiter.Length, _recordSeparator.Length))
        {
            _state = State.FieldStart;
        }
    }

    private bool TryEndUnquotedField(string separator)
    {
        if (separator.Length == 0 || !EndsWith(_field, separator))
        {
            return false;
        }

        _field.Length -= separator.Length;
        EmitField(enclose: false);
        _output.Append(separator);
        return true;
    }

    private void EmitField(bool enclose)
    {
        if (!enclose)
        {
            for (int index = 0; index < _field.Length && !enclose; index++)
            {
                enclose = _field[index] == _quote;
            }
        }

        if (!enclose)
        {
            _output.Append(_field);
            _field.Clear();
            return;
        }

        _output.Append(_quote);
        for (int index = 0; index < _field.Length; index++)
        {
            char character = _field[index];
            if (character == _quote || (character == _escape && _escape != _quote))
            {
                _output.Append(_escape);
            }

            _output.Append(character);
        }

        _output.Append(_quote);
        _field.Clear();
    }

    private static bool EndsWith(StringBuilder text, string value)
    {
        if (text.Length < value.Length)
        {
            return false;
        }

        int start = text.Length - value.Length;
        for (int index = 0; index < value.Length; index++)
        {
            if (text[start + index] != value[index])
            {
                return false;
            }
        }

        return true;
    }

    private void FlushOutputWhenLarge()
    {
        if (_output.Length >= OutputFlushThreshold)
        {
            WriteOutput();
        }
    }

    private void WriteOutput()
    {
        if (_output.Length == 0)
        {
            return;
        }

#if NET8_0_OR_GREATER
        _inner.Write(_output);
#else
        _inner.Write(_output.ToString());
#endif
        _output.Clear();
    }
}
