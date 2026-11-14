using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a QR code or barcode anchored at a worksheet cell.</summary>
/// <para>The code is an SVG picture with a PNG fallback, so Excel draws it sharply and older readers still show it; OfficeIMO's PDF conversion places it at print resolution. Two-dimensional codes default to 1.25 inches wide; linear barcodes to 0.33 mm per bar module, with their text underneath. The alternative text defaults to the symbology and the encoded text.</para>
/// <example><summary>Put a QR code next to a product row.</summary>
/// <code>New-OfficeExcel -Path .\Catalog.xlsx {
///     ExcelSheet 'Products' {
///         ExcelCell -Address A2 -Value 'Contoso Router X2'
///         ExcelBarcode 'https://contoso.example/x2' -Address B2 -Width 2cm
///     }
/// }</code></example>
/// <example><summary>Add an EAN-13 barcode to an existing workbook.</summary>
/// <code>ExcelBarcode '590123412345' -Type Ean -Document $book -Sheet Labels -Address C4</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelBarcode")]
[OutputType(typeof(ExcelImage))]
[Alias("ExcelBarcode")]
public sealed class AddOfficeExcelBarcodeCommand : OfficeBarcodeCommandBase {
    /// <summary>Workbook to modify outside the Excel DSL.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Sheet name, one-based position or native worksheet belonging to Document.</summary>
    [Parameter]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 reference of the anchor cell; defaults to A1.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string Address { get; set; } = "A1";

    /// <summary>Return the native editable ExcelImage.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Address.Contains(":"))
            throw new PSArgumentException("Address must be a single cell; the code keeps its own proportions.", nameof(Address));
        ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        ExcelImage image = sheet.AddBarcode(Address, Type, Text!, Configure);
        if (PassThru) WriteObject(image);
    }
}
