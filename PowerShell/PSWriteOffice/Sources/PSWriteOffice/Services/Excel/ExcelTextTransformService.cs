using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTextTransformService
{
    public static string Apply(ScriptBlock? script, string value)
    {
        if (script == null)
        {
            return value;
        }

        var result = script.InvokeReturnAsIs(value);
        if (result == null)
        {
            return string.Empty;
        }

        return LanguagePrimitives.ConvertTo<string>(result);
    }
}
