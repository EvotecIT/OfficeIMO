namespace OfficeIMO.Project;

internal sealed partial class ProjectCustomFieldCalculator {
    private ProjectFormulaValue Builtin(ProjectEntity entity, string name) {
        // Stable PjField identifiers, including the automatically scheduled StartText/DurationText aliases exported in formulas.
        string field = name.ToLowerInvariant();
        object? value;
        if (entity is ProjectTask task) value = field switch {
            "name" or "188743694" => task.Name ?? "",
            "id" or "188743703" => task.DisplayId is int display ? (decimal)display : null,
            "unique id" or "188743766" => (decimal)task.Uid,
            "start" or "188743715" or "188744965" => task.Start,
            "finish" or "188743716" or "188744966" => task.Finish,
            "duration" or "188743709" or "188744967" => task.Duration is ProjectDuration duration ? duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, duration.IsElapsed, _document) : null,
            "work" or "188743680" => task.Work?.Minutes,
            "actual work" or "188743682" => task.ActualWork?.Minutes,
            "remaining work" or "188743684" => task.RemainingWork?.Minutes,
            "cost" or "188743685" => task.Cost,
            "actual cost" or "188743687" => task.ActualCost,
            "remaining cost" or "188743690" => task.RemainingCost,
            "% complete" or "188743712" => task.PercentComplete is int percent ? (decimal)percent : null,
            "% work complete" or "188743713" => task.PercentWorkComplete is int workPercent ? (decimal)workPercent : null,
            "physical % complete" or "188744799" => task.PhysicalPercentComplete is int physical ? (decimal)physical : null,
            "priority" or "188743705" => task.Priority is int priority ? (decimal)priority : null,
            "summary" or "188743772" => task.IsSummary,
            "milestone" or "188743704" => task.IsMilestone,
            _ => throw new NotSupportedException("Task field '" + name + "' is outside the supported formula profile.")
        };
        else {
            var resource = (ProjectResource)entity;
            value = field switch {
                "name" or "205520897" => resource.Name ?? "",
                "id" or "205520896" => resource.DisplayId is int display ? (decimal)display : null,
                "unique id" or "205520923" => (decimal)resource.Uid,
                "work" or "205520909" => resource.Work?.Minutes,
                "actual work" or "205520910" => resource.ActualWork?.Minutes,
                "remaining work" or "205520918" => resource.RemainingWork?.Minutes,
                "cost" or "205520908" => resource.Cost,
                "actual cost" or "205520907" => resource.ActualCost,
                "remaining cost" or "205520917" => resource.RemainingCost,
                _ => throw new NotSupportedException("Resource field '" + name + "' is outside the supported formula profile.")
            };
        }
        if (value == null) throw new InvalidDataException("The referenced field has no stored value: " + name);
        return new ProjectFormulaValue(value);
    }
}
