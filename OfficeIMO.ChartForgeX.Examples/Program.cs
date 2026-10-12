using OfficeIMO.ChartForgeX.Examples;
using OfficeIMO.PowerPoint;

if (args.Length is < 1 or > 2 || (args.Length == 2 && args[1] != "--native-powerpoint")) {
    Console.Error.WriteLine("Usage: OfficeIMO.ChartForgeX.Examples <output-directory> [--native-powerpoint]");
    return 1;
}

string output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var specimens = VisualSpecimens.Create();
DocumentDelivery.Write(output, specimens);
DeliveryEvidence.RenderSavedDocuments(output);

if (args.Length == 2) {
    var result = PowerPointDesktopReferenceRenderer.TryRender(
        Path.Combine(output, "service-review.pptx"), Path.Combine(output, "powerpoint-desktop"), enabled: true);
    DeliveryEvidence.SaveJson(Path.Combine(output, "powerpoint-desktop.json"), result);
    Console.WriteLine("PowerPoint Desktop: " + result.Status + ": " + result.Message);
    if (!result.IsSuccessful) return 2;
}

Console.WriteLine("Saved documents and managed previews: " + output);
Console.WriteLine("Managed previews describe OfficeIMO layout; they do not establish Microsoft Office application rendering.");
return 0;
