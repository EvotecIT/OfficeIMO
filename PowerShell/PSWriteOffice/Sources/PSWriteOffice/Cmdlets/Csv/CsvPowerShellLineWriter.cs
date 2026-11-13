using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellLineWriter : TextWriter
{
    private const int PendingFlushThreshold = 32 * 1024;
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
        if (TryGetWholeRecordLength(new ReadOnlySpan<char>(buffer, index, count), out var length))
        {
            EmitRecord(new string(buffer, index, length));
            return;
        }

        _pending.Append(buffer, index, count);
        FlushLargeBuffer();
    }

    public override void Write(string? value)
    {
        if (value != null && TryGetWholeRecordLength(value.AsSpan(), out var length))
        {
            EmitRecord(value.Substring(0, length));
            return;
        }

        _pending.Append(value);
        FlushLargeBuffer();
    }

#if NET6_0_OR_GREATER
    public override void Write(StringBuilder? value)
    {
        // The row writer passes one record per call in its reusable buffer; emit it without copying it first.
        if (value != null && _pending.Length == 0)
        {
            var chunks = 0;
            ReadOnlyMemory<char> single = default;
            foreach (var chunk in value.GetChunks())
            {
                single = chunk;
                if (++chunks > 1)
                {
                    break;
                }
            }

            if (chunks == 1 && TryGetWholeRecordLength(single.Span, out var length))
            {
                EmitRecord(value.ToString(0, length));
                return;
            }
        }

        _pending.Append(value);
        FlushLargeBuffer();
    }
#endif

    // True when nothing is buffered and the text is exactly one complete record followed by its line break.
    private bool TryGetWholeRecordLength(ReadOnlySpan<char> text, out int length)
    {
        length = 0;
        if (_pending.Length != 0 || !IsAtRecordStart())
        {
            return false;
        }

        var stop = FindRecordEnd(text, 0);
        if (stop < 0 || GetRecordBreakEnd(text, stop) != text.Length)
        {
            return false;
        }

        length = stop;
        return true;
    }

    private void EmitRecord(string record)
    {
        _cmdlet.WriteObject(record);
        _atFieldStart = true;
    }

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

    // A record that starts cleanly and whose end FindRecordEnd can locate is exactly the text before its line break;
    // the character state machine would produce the same line.
    private bool TryEmitQuoteFreeRecord(string text, ref int index)
    {
        if (!IsAtRecordStart())
        {
            return false;
        }

        var span = text.AsSpan(index);
        var stop = FindRecordEnd(span, 0);
        var next = stop < 0 ? -1 : GetRecordBreakEnd(span, stop);
        if (next < 0)
        {
            return false;
        }

        EmitRecord(text.Substring(index, stop));
        index += next;
        return true;
    }

    private bool IsAtRecordStart() =>
        _line.Length == 0 && _atFieldStart && !_pendingCarriageReturn && !_inQuotes && !_pendingQuoteInQuotedField && _delimiterMatchIndex == 0;

    // Index just past the line break at stop, or -1 when the break is a lone or trailing carriage return.
    private static int GetRecordBreakEnd(ReadOnlySpan<char> text, int stop)
    {
        if (text[stop] == '\n')
        {
            return stop + 1;
        }

        return stop + 1 < text.Length && text[stop + 1] == '\n' ? stop + 2 : -1;
    }

    // Returns the index of the line break that ends the record starting at recordStart, or -1 when the record is
    // incomplete or needs the character state machine. Quotes follow the same rules as that state machine: a quote
    // opens a field only at a field start, a doubled quote stays inside the field, and line breaks inside quotes
    // belong to the record. Multi-character delimiters keep the state machine for quoted records.
    private int FindRecordEnd(ReadOnlySpan<char> text, int recordStart)
    {
        var position = recordStart;
        var inQuotes = false;
        while (true)
        {
            var remaining = text.Slice(position);
            var found = inQuotes
                ? remaining.IndexOf('"')
                : remaining.IndexOfAny('"', '\r', '\n');
            if (found < 0)
            {
                return -1;
            }

            var stop = position + found;

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
