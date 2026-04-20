using System.Runtime.InteropServices;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using System.Threading;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public class OfficeMarkupExcelInteropTests {
    [Fact]
    public void ExcelExporter_OpensMarkupWorkbookThroughExcelCom_WhenExcelIsAvailable() {
        if (!IsWindowsPlatform()) {
            return;
        }

        if (!IsExcelComAvailable()) {
            return;
        }

        var markup = """
---
profile: workbook
title: Workbook COM Smoke
---

::range address=Data!A1
Quarter,Revenue,Costs
Q1,120,85
Q2,180,94
Q3,260,132
Q4,320,150

::table name="RevenueTable" range=Data!A1:C5 header=true

::format target=Data!B2:C5 numberFormat="#,##0" fill=#D9EAD3 color=#112233 align=center valign=middle border=thin border-color=#445566

::chart type=column title="Quarterly Revenue" source=Data!RevenueTable cell=Dashboard!B2 width=480 height=320 category-title=Quarter value-title=Amount value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path
            });

            var workbookSheetCount = OpenWorkbookViaExcelCom(path);
            Assert.True(workbookSheetCount >= 2, "Expected the exported workbook to open through Excel COM and contain at least two worksheets.");
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

#if NET5_0_OR_GREATER
    [SupportedOSPlatformGuard("windows")]
#endif
    private static bool IsWindowsPlatform() =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    private static bool IsExcelComAvailable() =>
        Type.GetTypeFromProgID("Excel.Application") != null;

#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    private static int OpenWorkbookViaExcelCom(string path) {
        Exception? failure = null;
        var worksheetCount = 0;

        var thread = new Thread(() => {
            object? excel = null;
            object? workbook = null;

            try {
                var excelType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Excel COM automation is not available.");
                excel = Activator.CreateInstance(excelType)
                    ?? throw new InvalidOperationException("Failed to create Excel COM automation instance.");

                excelType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });
                excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });

                var workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excel, null);
                var workbooksType = workbooks!.GetType();
                workbook = workbooksType.InvokeMember("Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    workbooks,
                    new object[] { path, 0, true });

                var workbookType = workbook!.GetType();
                var worksheets = workbookType.InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, workbook, null);
                worksheetCount = (int)worksheets!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, worksheets, null)!;

                if (Marshal.IsComObject(worksheets)) {
                    Marshal.FinalReleaseComObject(worksheets);
                }

                if (Marshal.IsComObject(workbooks)) {
                    Marshal.FinalReleaseComObject(workbooks);
                }
            } catch (Exception ex) {
                failure = ex;
            } finally {
                try {
                    if (workbook != null) {
                        workbook.GetType().InvokeMember("Close",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null,
                            workbook,
                            new object[] { false });
                    }
                } catch {
                }

                try {
                    if (excel != null) {
                        excel.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excel, null);
                    }
                } catch {
                }

                if (workbook != null && Marshal.IsComObject(workbook)) {
                    Marshal.FinalReleaseComObject(workbook);
                }

                if (excel != null && Marshal.IsComObject(excel)) {
                    Marshal.FinalReleaseComObject(excel);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (failure != null) {
            throw new InvalidOperationException("Excel COM smoke test failed for the exported markup workbook.", failure);
        }

        return worksheetCount;
    }
}
