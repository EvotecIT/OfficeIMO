using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static void ReadResourceCapacity(ProjectResource resource, XElement element, CancellationToken token) {
        var document = resource.Document;
        foreach (var period in Children(element, "AvailabilityPeriods", "AvailabilityPeriod")) {
            token.ThrowIfCancellationRequested();
            var item = resource.AvailabilityPeriods.Add(); Attach(document, item, period);
            ReadFields(item, period, document, ProjectXmlFields.ResourceAvailability);
        }
        foreach (var period in Children(element, "Rates", "Rate")) {
            token.ThrowIfCancellationRequested();
            var item = resource.Rates.Add(); Attach(document, item, period);
            ReadFields(item, period, document, ProjectXmlFields.ResourceRate);
        }
    }
    private static void WriteResourceCapacity(ProjectResource resource, XElement element, CancellationToken token) {
        var document = resource.Document;
        ReplaceContainer(element, "AvailabilityPeriods", "AvailabilityPeriod", resource.AvailabilityPeriods.Select(item => {
            token.ThrowIfCancellationRequested(); var node = NewNode(document, item, "AvailabilityPeriod");
            ProjectXmlFields.Write(item, node, document, ProjectXmlFields.ResourceAvailability, new[] { "AvailableFrom", "AvailableTo", "AvailableUnits" });
            return node;
        }), ResourceOrder);
        ReplaceContainer(element, "Rates", "Rate", resource.Rates.Select(item => {
            token.ThrowIfCancellationRequested(); var node = NewNode(document, item, "Rate");
            ProjectXmlFields.Write(item, node, document, ProjectXmlFields.ResourceRate,
                new[] { "RatesFrom", "RatesTo", "RateTable", "StandardRate", "StandardRateFormat", "OvertimeRate", "OvertimeRateFormat", "CostPerUse" });
            return node;
        }), ResourceOrder);
    }
}
