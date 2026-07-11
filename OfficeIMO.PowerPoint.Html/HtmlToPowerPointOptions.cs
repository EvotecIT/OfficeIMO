namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Options for importing semantic OfficeIMO PowerPoint HTML into a presentation.
/// </summary>
public sealed class HtmlToPowerPointOptions {
    /// <summary>
    /// Imports semantic slide tables as native PowerPoint tables.
    /// </summary>
    public bool ImportTables { get; set; } = true;

    /// <summary>
    /// Imports embedded data URI pictures from the semantic picture inventory.
    /// </summary>
    public bool ImportPictures { get; set; } = true;

    /// <summary>
    /// Imports chart inventory items as native charts with reconstructed placeholder data.
    /// </summary>
    public bool ImportChartInventory { get; set; } = true;

    /// <summary>
    /// Imports presenter notes when they are present in the semantic extraction proof.
    /// </summary>
    public bool ImportNotes { get; set; } = true;
}
