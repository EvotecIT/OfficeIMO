using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellLineWriter : TextWriter
{
    private const int PendingFlushThreshold = 32 * 1024;
    private static readonly char[] RecordStopCharacters = { '"', '\r', '\n' };
    private readonly PSCmdlet _cmdlet;
    private readonly string _delimiterText;
    private readonly bool _parseQuotedRecords;
    private readonly StringBuilder _line = new();
    private readonly StringBuilder _pending = new();
    private readonly char[] _specialCharacters;
    private bool _pendingCarriageReturn;
    private bool _inQuotes;
    private bool _pendingQuoteInQuotedField;
    private bool _atFieldStart = true;
    private int _delimiterMatchIndex;

    public CsvPowerShellLineWriter(PSCmdlet cmdlet, char delimiter, CsvQuoteMode quoteMode)
        : this(cmdlet, delimiter.ToString(), quoteMode)
    {
    }

    public CsvPowerShellLineWriter(PSCmdlet cmdlet, string delimiterText, CsvQuoteMode quoteMode)
    {
        _cmdlet = cmdlet ?? throw new ArgumentNullException(nameof(cmdlet));
        _delimiterText = string.IsNullOrEmpty(delimiterText) ? "," : delimiterText;
        _parseQuotedRecords = quoteMode != CsvQuoteMode.Never;
        _specialCharacters = new[] { '"', '\r', '\n', _delimiterText[0] };
    }

    public override Encoding Encoding => Encoding.UTF8;

    // Writes are buffered and split into records in bulk on Flush. Scanning every character through virtual
    // Write(char) calls cost more than formatting the CSV itself.
    public override void Write(char value)
    {
        _pending.Append(value);
        FlushLargeBuffer();
    }

    public override void Write(char[] buffer, int index, int count)
    {
        _pending.Append(buffer, index, count);
        FlushLargeBuffer();
    }

    public override void Write(string? value)
    {
        _pending.Append(value);
        FlushLargeBuffer();
    }

#if NET6_0_OR_GREATER
    public override void Write(StringBuilder? value)
    {
        _pending.Append(value);
        FlushLargeBuffer();
    }
#endif

    // Multi-row writes such as a DataTable keep streaming records instead of buffering the whole output.
    private void FlushLargeBuffer()
    {
        if (_pending.Length >= PendingFlushThreshold)
        {
            Flush();
        }
    }

    public override void Flush()
    {
        if (_pending.Length == 0)
        {
            return;
        }

        var text = _pending.ToString();
        _pending.Clear();
        var index = 0;
        while (index < text.Length)
        {
            if (TryEmitQuoteFreeRecord(text, ref index))
            {
                continue;
            }

            if (!_pendingCarriageReturn && !_pendingQuoteInQuotedField && _delimiterMatchIndex == 0)
            {
                var next = text.IndexOfAny(_specialCharacters, index);
                if (next < 0)
                {
                    next = text.Length;
                }

                if (next > index)
                {
                    // A run of ordinary characters only clears the field-start state, as it would one character at a time.
                    _line.Append(text, index, next - index);
                    _atFieldStart = false;
                    index = next;
                    continue;
                }
            }

            ProcessCharacter(text[index]);
            index++;
        }
    }

    // A record that starts cleanly and reaches its line break without a quote is exactly the text before that
    // break; the character state machine would produce the same line. Quoted records take the full path.
    private bool TryEmitQuoteFreeRecord(string text, ref int index)
    {
        if (_line.Length != 0 || !_atFieldStart || _pendingCarriageReturn || _inQuotes || _pendingQuoteInQuotedField || _delimiterMatchIndex != 0)
        {
            return false;
        }

        var stop = FindRecordEnd(text, index);
        if (stop < 0)
        {
            return false;
        }

        int next;
        if (text[stop] == '\n')
        {
            next = stop + 1;
        }
        else if (text[stop] == '\r' && stop + 1 < text.Length && text[stop + 1] == '\n')
        {
            next = stop + 2;
        }
        else
        {
            return false;
        }

        _cmdlet.WriteObject(text.Substring(index, stop - index));
        _atFieldStart = true;
        index = next;
        return true;
    }

    // Returns the index of the line break that ends the record starting at recordStart, or -1 when the record is
    // incomplete or needs the character state machine. Quotes follow the same rules as that state machine: a quote
    // opens a field only at a field start, a doubled quote stays inside the field, and line breaks inside quotes
    // belong to the record. Multi-character delimiters keep the state machine for quoted records.
    private int FindRecordEnd(string text, int recordStart)
    {
        var position = recordStart;
        var inQuotes = false;
        while (true)
        {
            var stop = inQuotes
                ? text.IndexOf('"', position)
                : text.IndexOfAny(RecordStopCharacters, position);
            if (stop < 0)
            {
                return -1;
            }

            if (text[stop] != '"')
            {
                return stop;
            }

            if (!_parseQuotedRecords)
            {
                position = stop + 1;
                continue;
            }

            if (inQuotes)
            {
                if (stop + 1 >= text.Length)
                {
                    return -1;
                }

                if (text[stop + 1] == '"')
                {
                    position = stop + 2;
                    continue;
                }

                inQuotes = false;
                position = stop + 1;
                continue;
            }

            if (_delimiterText.Length != 1 ||
                (stop != recordStart && text[stop - 1] != _delimiterText[0]))
            {
                return -1;
            }

            inQuotes = true;
            position = stop + 1;
        }
    }

    private void ProcessCharacter(char value)
    {
        if (ResolvePendingQuote(value))
        {
            return;
        }

        if (_pendingCarriageReturn)
        {
            if (value == '\n')
            {
                if (_inQuotes)
                {
                    _line.Append("\r\n");
                }
                else
                {
                    EmitLine();
                }

                _pendingCarriageReturn = false;
                return;
            }

            if (_inQuotes)
            {
                _line.Append('\r');
            }
            else
            {
                EmitLine();
            }

            _pendingCarriageReturn = false;
        }

        if (value == '"')
        {
            _line.Append(value);
            if (!_parseQuotedRecords)
            {
                _atFieldStart = false;
            }
            else if (_inQuotes)
            {
                _pendingQuoteInQuotedField = true;
            }
            else if (_atFieldStart)
            {
                _inQuotes = true;
            }

            _atFieldStart = false;
            return;
        }

        if (value == '\r')
        {
            _pendingCarriageReturn = true;
            return;
        }

        if (value == '\n')
        {
            if (_inQuotes)
            {
                _line.Append(value);
            }
            else
            {
                EmitLine();
            }

            return;
        }

        _line.Append(value);
        UpdateFieldStartState(value);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Flush();
            if (_pendingCarriageReturn)
            {
                if (_inQuotes)
                {
                    _line.Append('\r');
                }
                else
                {
                    EmitLine();
                }

                _pendingCarriageReturn = false;
            }

            if (_line.Length > 0)
            {
                EmitLine();
            }
        }

        base.Dispose(disposing);
    }

    private void EmitLine()
    {
        _cmdlet.WriteObject(_line.ToString());
        _line.Clear();
        _inQuotes = false;
        _pendingQuoteInQuotedField = false;
        _atFieldStart = true;
        _delimiterMatchIndex = 0;
    }

    private bool ResolvePendingQuote(char value)
    {
        if (!_pendingQuoteInQuotedField)
        {
            return false;
        }

        _pendingQuoteInQuotedField = false;
        if (value == '"')
        {
            _line.Append(value);
            _atFieldStart = false;
            return true;
        }

        _inQuotes = false;
        _atFieldStart = false;
        return false;
    }

    private void UpdateFieldStartState(char value)
    {
        if (_inQuotes)
        {
            _atFieldStart = false;
            _delimiterMatchIndex = 0;
            return;
        }

        if (_atFieldStart && _delimiterText.Length > 1 && value == _delimiterText[0])
        {
            _delimiterMatchIndex = 1;
            return;
        }

        if (value == _delimiterText[_delimiterMatchIndex])
        {
            _delimiterMatchIndex++;
            if (_delimiterMatchIndex == _delimiterText.Length)
            {
                _atFieldStart = true;
                _delimiterMatchIndex = 0;
            }
            else
            {
                _atFieldStart = false;
            }

            return;
        }

        _delimiterMatchIndex = value == _delimiterText[0] ? 1 : 0;
        _atFieldStart = false;
    }
}
