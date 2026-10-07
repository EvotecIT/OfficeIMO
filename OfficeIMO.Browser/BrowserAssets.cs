namespace OfficeIMO.Browser;

/// <summary>Dependency-free export scripts for inline reports, offline bundles and JavaScript modules.</summary>
/// <remarks>Classic scripts extend <c>globalThis.OfficeIMO</c>. Modules have no relative runtime imports.</remarks>
public static class BrowserAssets {
    /// <summary>The classic XLSX, CSV and PDF runtime for a single script tag.</summary>
    public static BrowserAsset Script { get; } = new BrowserAsset("officeimo.js");

    /// <summary>The standalone classic XLSX runtime.</summary>
    public static BrowserAsset XlsxScript { get; } = new BrowserAsset("officeimo-xlsx.js");

    /// <summary>The standalone classic CSV runtime.</summary>
    public static BrowserAsset CsvScript { get; } = new BrowserAsset("officeimo-csv.js");

    /// <summary>The combined ES module.</summary>
    public static BrowserAsset Module { get; } = new BrowserAsset("officeimo.mjs");

    /// <summary>The standalone XLSX ES module.</summary>
    public static BrowserAsset XlsxModule { get; } = new BrowserAsset("officeimo-xlsx.mjs");

    /// <summary>The standalone CSV ES module.</summary>
    public static BrowserAsset CsvModule { get; } = new BrowserAsset("officeimo-csv.mjs");

    /// <summary>The standalone classic PDF table runtime.</summary>
    public static BrowserAsset PdfScript { get; } = new BrowserAsset("officeimo-pdf.js");

    /// <summary>The standalone PDF table ES module.</summary>
    public static BrowserAsset PdfModule { get; } = new BrowserAsset("officeimo-pdf.mjs");

    /// <summary>The optional DataTables Excel/CSV/PDF integration, including its document writers.</summary>
    public static BrowserAsset DataTablesScript { get; } = new BrowserAsset("officeimo-datatables.js");

    /// <summary>The optional standalone DataTables ES module, including its document writers.</summary>
    public static BrowserAsset DataTablesModule { get; } = new BrowserAsset("officeimo-datatables.mjs");
}
