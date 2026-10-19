using OfficeIMO.Excel;
using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelHttpLoadService
{
    public static ExcelHttpLoadOptions CreateOptions(SwitchParameter allowHttp)
    {
        return new ExcelHttpLoadOptions
        {
            SchemePolicy = allowHttp.IsPresent
                ? ExcelUriSchemePolicy.HttpAndHttps
                : ExcelUriSchemePolicy.HttpsOnly
        };
    }
}
