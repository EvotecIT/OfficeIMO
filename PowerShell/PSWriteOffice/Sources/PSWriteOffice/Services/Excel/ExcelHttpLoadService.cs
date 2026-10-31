using OfficeIMO.Excel;
using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelHttpLoadService
{
    public static ExcelLoadOptions CreateOptions(SwitchParameter allowHttp)
    {
        return new ExcelLoadOptions
        {
            SchemePolicy = allowHttp.IsPresent
                ? ExcelUriSchemePolicy.HttpAndHttps
                : ExcelUriSchemePolicy.HttpsOnly
        };
    }
}
