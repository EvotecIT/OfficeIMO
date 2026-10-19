using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellLineWriter : TextWriter
{
    private readonly PSCmdlet _cmdlet;
    private readonly string _delimiterText;
    private readonly bool _parseQuotedRecords;
    private readonly StringBuilder _line = new();
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
    }

    public override Encoding Encoding => Encoding.UTF8;

    public override void Write(char value)
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

    public override void Write(string? value)
    {
        if (value is null)
        {
            return;
        }

        foreach (var character in value)
        {
            Write(character);
        }
    }

#if NET6_0_OR_GREATER
    public override void Write(StringBuilder? value)
    {
        if (value is null)
        {
            return;
        }

        for (var i = 0; i < value.Length; i++)
        {
            Write(value[i]);
        }
    }
#endif

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
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
