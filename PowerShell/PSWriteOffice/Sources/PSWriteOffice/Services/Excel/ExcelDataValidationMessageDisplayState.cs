using System;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal sealed class ExcelDataValidationMessageDisplayState
{
    // OfficeIMO 0.6.48 updates message text through a public API, but it does not expose
    // the existing showInputMessage/showErrorMessage flags in its validation snapshot.
    private static readonly PropertyInfo? WorksheetRootProperty = typeof(ExcelSheet).GetProperty("WorksheetRoot", BindingFlags.Instance | BindingFlags.NonPublic);
    private const int MaxExcelRow = 1048576;
    private const int MaxExcelColumn = 16384;

    private readonly List<DataValidationDisplaySnapshot> _snapshots;

    private ExcelDataValidationMessageDisplayState(List<DataValidationDisplaySnapshot> snapshots)
    {
        _snapshots = snapshots;
    }

    public static ExcelDataValidationMessageDisplayState Capture(ExcelSheet sheet, string targetRange)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (string.IsNullOrWhiteSpace(targetRange)) throw new ArgumentNullException(nameof(targetRange));

        Worksheet worksheet = GetWorksheetRoot(sheet);
        var filter = ParseReferenceArgument(targetRange);
        var snapshots = new List<DataValidationDisplaySnapshot>();
        var validations = worksheet.GetFirstChild<DataValidations>();
        if (validations == null)
        {
            return new ExcelDataValidationMessageDisplayState(snapshots);
        }

        foreach (DataValidation validation in validations.Elements<DataValidation>())
        {
            string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
            if (!ReferenceListOverlaps(range, filter)) continue;

            snapshots.Add(new DataValidationDisplaySnapshot(
                range,
                validation.PromptTitle?.Value,
                validation.Prompt?.Value,
                validation.ErrorTitle?.Value,
                validation.Error?.Value,
                validation.ShowInputMessage?.Value,
                validation.ShowErrorMessage?.Value,
                !string.IsNullOrEmpty(validation.PromptTitle?.Value) || !string.IsNullOrEmpty(validation.Prompt?.Value),
                !string.IsNullOrEmpty(validation.ErrorTitle?.Value) || !string.IsNullOrEmpty(validation.Error?.Value)));
        }

        return new ExcelDataValidationMessageDisplayState(snapshots);
    }

    public static bool ReferenceListOverlapsTarget(string referenceList, string targetRange)
        => ReferenceListOverlaps(referenceList, ParseReferenceArgument(targetRange));

    public void Restore(
        ExcelSheet sheet,
        string targetRange,
        string? promptTitle,
        bool setPromptTitle,
        string? prompt,
        bool setPrompt,
        string? errorTitle,
        bool setErrorTitle,
        string? error,
        bool setError,
        bool? showInputMessage,
        bool? showErrorMessage)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (string.IsNullOrWhiteSpace(targetRange)) throw new ArgumentNullException(nameof(targetRange));

        Worksheet worksheet = GetWorksheetRoot(sheet);
        var filter = ParseReferenceArgument(targetRange);
        var validations = worksheet.GetFirstChild<DataValidations>();
        if (validations == null)
        {
            return;
        }

        var usedSnapshots = new bool[_snapshots.Count];
        bool changed = false;
        foreach (DataValidation validation in validations.Elements<DataValidation>())
        {
            string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
            if (!ReferenceListOverlaps(range, filter)) continue;

            int snapshotIndex = FindSnapshot(range, usedSnapshots);
            if (snapshotIndex >= 0)
            {
                usedSnapshots[snapshotIndex] = true;
            }

            DataValidationDisplaySnapshot? snapshot = snapshotIndex >= 0 ? _snapshots[snapshotIndex] : null;
            if (setPromptTitle)
            {
                changed |= SetString(validation, MessageField.PromptTitle, promptTitle);
            }
            else if (snapshot != null)
            {
                changed |= SetString(validation, MessageField.PromptTitle, snapshot.PromptTitle);
            }

            if (setPrompt)
            {
                changed |= SetString(validation, MessageField.Prompt, prompt);
            }
            else if (snapshot != null)
            {
                changed |= SetString(validation, MessageField.Prompt, snapshot.Prompt);
            }

            if (setErrorTitle)
            {
                changed |= SetString(validation, MessageField.ErrorTitle, errorTitle);
            }
            else if (snapshot != null)
            {
                changed |= SetString(validation, MessageField.ErrorTitle, snapshot.ErrorTitle);
            }

            if (setError)
            {
                changed |= SetString(validation, MessageField.Error, error);
            }
            else if (snapshot != null)
            {
                changed |= SetString(validation, MessageField.Error, snapshot.Error);
            }

            bool hasInputMessageText = HasMessageText(validation.PromptTitle?.Value, validation.Prompt?.Value);
            bool hasErrorMessageText = HasMessageText(validation.ErrorTitle?.Value, validation.Error?.Value);
            if (showInputMessage.HasValue)
            {
                changed |= SetBoolean(validation, isInputMessage: true, showInputMessage);
            }
            else if (snapshot?.HasInputMessageText == true)
            {
                changed |= SetBoolean(validation, isInputMessage: true, snapshot.ShowInputMessage);
            }
            else
            {
                changed |= SetBoolean(validation, isInputMessage: true, hasInputMessageText);
            }

            if (showErrorMessage.HasValue)
            {
                changed |= SetBoolean(validation, isInputMessage: false, showErrorMessage);
            }
            else if (snapshot?.HasErrorMessageText == true)
            {
                changed |= SetBoolean(validation, isInputMessage: false, snapshot.ShowErrorMessage);
            }
            else
            {
                changed |= SetBoolean(validation, isInputMessage: false, hasErrorMessageText);
            }
        }

        if (changed)
        {
            worksheet.Save();
        }
    }

    private static Worksheet GetWorksheetRoot(ExcelSheet sheet)
    {
        if (WorksheetRootProperty?.GetValue(sheet) is Worksheet worksheet)
        {
            return worksheet;
        }

        throw new InvalidOperationException("Unable to access the worksheet XML for Excel data validation message display state.");
    }

    private int FindSnapshot(string range, bool[] usedSnapshots)
    {
        for (int i = 0; i < _snapshots.Count; i++)
        {
            if (!usedSnapshots[i] && string.Equals(_snapshots[i].Range, range, StringComparison.Ordinal))
            {
                return i;
            }
        }

        return -1;
    }

    private enum MessageField
    {
        PromptTitle,
        Prompt,
        ErrorTitle,
        Error
    }

    private static bool SetString(DataValidation validation, MessageField field, string? value)
    {
        string? current = field switch
        {
            MessageField.PromptTitle => validation.PromptTitle?.Value,
            MessageField.Prompt => validation.Prompt?.Value,
            MessageField.ErrorTitle => validation.ErrorTitle?.Value,
            MessageField.Error => validation.Error?.Value,
            _ => null
        };

        if (string.Equals(current, value, StringComparison.Ordinal))
        {
            return false;
        }

        switch (field)
        {
            case MessageField.PromptTitle:
                validation.PromptTitle = value;
                break;
            case MessageField.Prompt:
                validation.Prompt = value;
                break;
            case MessageField.ErrorTitle:
                validation.ErrorTitle = value;
                break;
            case MessageField.Error:
                validation.Error = value;
                break;
        }

        return true;
    }

    private static bool SetBoolean(DataValidation validation, bool isInputMessage, bool? value)
    {
        BooleanValue? current = isInputMessage ? validation.ShowInputMessage : validation.ShowErrorMessage;
        if (current?.Value == value)
        {
            return false;
        }

        BooleanValue? next = value.HasValue ? new BooleanValue(value.Value) : null;
        if (isInputMessage)
        {
            validation.ShowInputMessage = next;
        }
        else
        {
            validation.ShowErrorMessage = next;
        }

        return true;
    }

    private static bool HasMessageText(string? title, string? message)
        => !string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(message);

    private static (int r1, int c1, int r2, int c2) ParseReferenceArgument(string reference)
    {
        if (TryParseReference(reference, out var bounds))
        {
            return bounds;
        }

        throw new ArgumentException($"Invalid A1 reference '{reference}'.", nameof(reference));
    }

    private static bool ReferenceListOverlaps(string referenceList, (int r1, int c1, int r2, int c2) filter)
    {
        foreach (string reference in referenceList.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            if (TryParseReference(reference, out var bounds) && RangesOverlapInclusive(filter, bounds))
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseReference(string reference, out (int r1, int c1, int r2, int c2) bounds)
    {
        bounds = default;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        string normalized = reference.Trim();
        int sheetSeparator = normalized.LastIndexOf('!');
        if (sheetSeparator >= 0)
        {
            normalized = normalized.Substring(sheetSeparator + 1);
        }

        string[] parts = normalized.Split(':');
        if (parts.Length > 2)
        {
            return false;
        }

        if (!TryParseReferenceEndpoint(parts[0], out ReferenceEndpoint first))
        {
            return false;
        }

        ReferenceEndpoint second = first;
        if (parts.Length == 2 && !TryParseReferenceEndpoint(parts[1], out second))
        {
            return false;
        }

        int r1;
        int c1;
        int r2;
        int c2;
        if (first.Row.HasValue || second.Row.HasValue)
        {
            r1 = first.Row ?? 1;
            r2 = second.Row ?? MaxExcelRow;
        }
        else
        {
            r1 = 1;
            r2 = MaxExcelRow;
        }

        if (first.Column.HasValue || second.Column.HasValue)
        {
            c1 = first.Column ?? 1;
            c2 = second.Column ?? MaxExcelColumn;
        }
        else
        {
            c1 = 1;
            c2 = MaxExcelColumn;
        }

        if (r2 < r1)
        {
            (r1, r2) = (r2, r1);
        }

        if (c2 < c1)
        {
            (c1, c2) = (c2, c1);
        }

        bounds = (r1, c1, r2, c2);
        return true;
    }

    private static bool TryParseReferenceEndpoint(string reference, out ReferenceEndpoint endpoint)
    {
        endpoint = default;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        string normalized = reference.Trim().Replace("$", string.Empty);
        int index = 0;
        int column = 0;
        while (index < normalized.Length && char.IsLetter(normalized[index]))
        {
            column = (column * 26) + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
            index++;
        }

        int row = 0;
        for (; index < normalized.Length; index++)
        {
            if (!char.IsDigit(normalized[index]))
            {
                return false;
            }

            row = (row * 10) + (normalized[index] - '0');
        }

        if (row == 0 && column == 0)
        {
            return false;
        }

        endpoint = new ReferenceEndpoint(row == 0 ? null : row, column == 0 ? null : column);
        return true;
    }

    private static bool RangesOverlapInclusive((int r1, int c1, int r2, int c2) first, (int r1, int c1, int r2, int c2) second)
        => first.r1 <= second.r2
            && second.r1 <= first.r2
            && first.c1 <= second.c2
            && second.c1 <= first.c2;

    private sealed class DataValidationDisplaySnapshot
    {
        public DataValidationDisplaySnapshot(
            string range,
            string? promptTitle,
            string? prompt,
            string? errorTitle,
            string? error,
            bool? showInputMessage,
            bool? showErrorMessage,
            bool hasInputMessageText,
            bool hasErrorMessageText)
        {
            Range = range;
            PromptTitle = promptTitle;
            Prompt = prompt;
            ErrorTitle = errorTitle;
            Error = error;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
            HasInputMessageText = hasInputMessageText;
            HasErrorMessageText = hasErrorMessageText;
        }

        public string Range { get; }

        public string? PromptTitle { get; }

        public string? Prompt { get; }

        public string? ErrorTitle { get; }

        public string? Error { get; }

        public bool? ShowInputMessage { get; }

        public bool? ShowErrorMessage { get; }

        public bool HasInputMessageText { get; }

        public bool HasErrorMessageText { get; }
    }

    private readonly struct ReferenceEndpoint
    {
        public ReferenceEndpoint(int? row, int? column)
        {
            Row = row;
            Column = column;
        }

        public int? Row { get; }

        public int? Column { get; }
    }
}
