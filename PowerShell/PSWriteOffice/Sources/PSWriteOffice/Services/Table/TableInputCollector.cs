using System.Collections;
using System.Collections.Generic;
using System.Data;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Table;

internal static class TableInputCollector
{
    public static void AddInput(ICollection<object?> items, object? value, bool preserveTabularInput = false)
    {
        if (value is null)
        {
            return;
        }

        if (ShouldExpand(value, preserveTabularInput))
        {
            foreach (var entry in (IEnumerable)value)
            {
                if (entry is not null)
                {
                    items.Add(entry);
                }
            }

            return;
        }

        items.Add(value);
    }

    public static object[] RequireRows(IReadOnlyCollection<object?> items, string parameterName)
    {
        if (items.Count == 0)
        {
            throw new System.Management.Automation.PSArgumentException("Provide at least one data row.", parameterName);
        }

        var rows = new object[items.Count];
        var index = 0;
        foreach (var item in items)
        {
            if (item is not null)
            {
                rows[index++] = item;
            }
        }

        if (index == rows.Length)
        {
            return rows;
        }

        System.Array.Resize(ref rows, index);
        return rows;
    }

    private static bool ShouldExpand(object value, bool preserveTabularInput)
    {
        return value is IEnumerable
            and not string
            and not IDictionary
            and not IDataReader
            and not DataSet
            && !PowerShellDictionaryAdapter.IsDictionaryLike(value)
            && (!preserveTabularInput || value is not DataTable && value is not DataView);
    }
}
