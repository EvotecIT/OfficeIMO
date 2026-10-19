namespace OfficeIMO.JavaScript.Assets;

/// <summary>Dependency-free export scripts for inline reports, offline bundles and JavaScript modules.</summary>
/// <remarks>Classic scripts extend <c>globalThis.OfficeIMO</c>. Modules have no relative runtime imports.</remarks>
public static class JavaScriptAssets {
    /// <summary>The classic XLSX, CSV and PDF runtime for a single script tag.</summary>
    public static JavaScriptAsset Script { get; } = new JavaScriptAsset("officeimo.js");

    /// <summary>The standalone classic XLSX runtime.</summary>
    public static JavaScriptAsset XlsxScript { get; } = new JavaScriptAsset("officeimo-xlsx.js");

    /// <summary>The standalone classic CSV runtime.</summary>
    public static JavaScriptAsset CsvScript { get; } = new JavaScriptAsset("officeimo-csv.js");

    /// <summary>The combined ES module.</summary>
    public static JavaScriptAsset Module { get; } = new JavaScriptAsset("officeimo.mjs");

    /// <summary>The standalone XLSX ES module.</summary>
    public static JavaScriptAsset XlsxModule { get; } = new JavaScriptAsset("officeimo-xlsx.mjs");

    /// <summary>The standalone CSV ES module.</summary>
    public static JavaScriptAsset CsvModule { get; } = new JavaScriptAsset("officeimo-csv.mjs");

    /// <summary>The standalone classic PDF document runtime.</summary>
    public static JavaScriptAsset PdfScript { get; } = new JavaScriptAsset("officeimo-pdf.js");

    /// <summary>The standalone PDF document ES module.</summary>
    public static JavaScriptAsset PdfModule { get; } = new JavaScriptAsset("officeimo-pdf.mjs");

    /// <summary>The optional DataTables Excel/CSV/PDF integration, including its document writers.</summary>
    public static JavaScriptAsset DataTablesScript { get; } = new JavaScriptAsset("officeimo-datatables.js");

    /// <summary>The optional standalone DataTables ES module, including its document writers.</summary>
    public static JavaScriptAsset DataTablesModule { get; } = new JavaScriptAsset("officeimo-datatables.mjs");

    /// <summary>The optional CanopyX portable grid capture integration, including its document writers.</summary>
    public static JavaScriptAsset CanopyXScript { get; } = new JavaScriptAsset("officeimo-canopyx.js");

    /// <summary>The optional standalone CanopyX ES module, including its document writers.</summary>
    public static JavaScriptAsset CanopyXModule { get; } = new JavaScriptAsset("officeimo-canopyx.mjs");
}
