namespace OfficeIMO.JavaScript.Assets;

/// <summary>Dependency-free export scripts for inline reports, offline bundles and JavaScript modules.</summary>
/// <remarks>
/// Classic scripts extend <c>globalThis.OfficeIMO</c>. Modules have no relative runtime imports, except the integration
/// adapter modules, which import the sibling <c>officeimo.mjs</c>. Every asset has a minified <see cref="JavaScriptAsset.Minified"/> variant.
/// </remarks>
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

    /// <summary>Optional bounded worker transport; the host owns worker creation and the export operation.</summary>
    public static JavaScriptAsset ExecutionScript { get; } = new JavaScriptAsset("officeimo-execution.js");

    /// <summary>Standalone bounded worker transport ES module.</summary>
    public static JavaScriptAsset ExecutionModule { get; } = new JavaScriptAsset("officeimo-execution.mjs");

    /// <summary>The optional DataTables Excel/CSV/PDF integration, including its document writers.</summary>
    public static JavaScriptAsset DataTablesScript { get; } = new JavaScriptAsset("officeimo-datatables.js");

    /// <summary>The optional standalone DataTables ES module, including its document writers.</summary>
    public static JavaScriptAsset DataTablesModule { get; } = new JavaScriptAsset("officeimo-datatables.mjs");

    /// <summary>The DataTables integration without document writers; load <see cref="Script"/> from the same build first.</summary>
    public static JavaScriptAsset DataTablesAdapterScript { get; } = new JavaScriptAsset("officeimo-datatables-adapter.js");

    /// <summary>The DataTables integration ES module without document writers; it imports the sibling <c>officeimo.mjs</c> (<see cref="Module"/>).</summary>
    public static JavaScriptAsset DataTablesAdapterModule { get; } = new JavaScriptAsset("officeimo-datatables-adapter.mjs");

    /// <summary>The optional CanopyX portable grid capture integration, including its document writers.</summary>
    public static JavaScriptAsset CanopyXScript { get; } = new JavaScriptAsset("officeimo-canopyx.js");

    /// <summary>The optional standalone CanopyX ES module, including its document writers.</summary>
    public static JavaScriptAsset CanopyXModule { get; } = new JavaScriptAsset("officeimo-canopyx.mjs");

    /// <summary>The CanopyX integration without document writers; load <see cref="Script"/> from the same build first.</summary>
    public static JavaScriptAsset CanopyXAdapterScript { get; } = new JavaScriptAsset("officeimo-canopyx-adapter.js");

    /// <summary>The CanopyX integration ES module without document writers; it imports the sibling <c>officeimo.mjs</c> (<see cref="Module"/>).</summary>
    public static JavaScriptAsset CanopyXAdapterModule { get; } = new JavaScriptAsset("officeimo-canopyx-adapter.mjs");

    /// <summary>The optional ChartForgeX chart capture integration (typed XLSX series and vector PDF), including its document writers.</summary>
    public static JavaScriptAsset ChartForgeXScript { get; } = new JavaScriptAsset("officeimo-chartforgex.js");

    /// <summary>The optional standalone ChartForgeX ES module, including its document writers.</summary>
    public static JavaScriptAsset ChartForgeXModule { get; } = new JavaScriptAsset("officeimo-chartforgex.mjs");

    /// <summary>The ChartForgeX integration without document writers; load <see cref="Script"/> from the same build first.</summary>
    public static JavaScriptAsset ChartForgeXAdapterScript { get; } = new JavaScriptAsset("officeimo-chartforgex-adapter.js");

    /// <summary>The ChartForgeX integration ES module without document writers; it imports the sibling <c>officeimo.mjs</c> (<see cref="Module"/>).</summary>
    public static JavaScriptAsset ChartForgeXAdapterModule { get; } = new JavaScriptAsset("officeimo-chartforgex-adapter.mjs");
}
