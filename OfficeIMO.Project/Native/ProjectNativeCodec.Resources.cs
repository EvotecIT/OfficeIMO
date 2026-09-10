namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    private static void ReadResources(ProjectDocument document, ProjectNativeTable table, ProjectLoadOptions options, CancellationToken token) {
        foreach (var record in table.Records) {
            token.ThrowIfCancellationRequested();
            record.Uid = record.Integer(0x0c40001b) ?? throw new InvalidDataException("Native resource UID missing.");
            if (record.Uid < 0) continue;
            CheckEntityBudget(document, options);
            var resource = new ProjectResource(document, record.Uid) {
                DisplayId = record.Integer(0x0c400000), Name = record.Text(0x0c400001),
                Type = record.Boolean(0x0c4002df) == true ? ProjectResourceType.Cost : record.Boolean(0x0c40012a) == true ? ProjectResourceType.Work : ProjectResourceType.Material,
                MaterialLabel = record.Text(0x0c40012b),
                StandardRate = record.Number(0x0c400006), OvertimeRate = record.Number(0x0c400007), CostPerUse = record.Number(0x0c400012) / 100m,
                Cost = record.Number(0x0c40000c) / 100m, ActualCost = record.Number(0x0c40000b) / 100m,
                Work = Work(record, 0x0c40000d), ActualWork = Work(record, 0x0c40000e), RemainingWork = Work(record, 0x0c400016)
            };
            decimal? units = record.Number(0x0c400004);
            if (units.HasValue) resource.MaxUnits = ProjectUnits.Fraction(units.Value / 10000m);
            if (record.Value(0x0c4002d8) is ProjectNativeValue guid) resource.Guid = new Guid(guid.Copy());
            int calendarId = record.Integer(0x0c400038) ?? -1;
            if (document.CalendarIndex.TryGetValue(calendarId, out var calendar)) {
                resource.Calendar = calendar;
                if (string.IsNullOrEmpty(calendar.Name)) calendar.Name = resource.Name;
            } else if (calendarId > 0) resource.SourceCalendarUid = calendarId;
            if (document.ResourceIndex.ContainsKey(resource.Uid)) throw new InvalidDataException("Duplicate native resource UID.");
            document.ResourceIndex.Add(resource.Uid, resource); document.Resources.Items.Add(resource);
            ReadCustomValues(document, record, resource.CustomFields, false, token);
            for (int number = 0; number <= 10; number++) {
                uint workId = number == 0 ? 0x0c40000fu : 0x0c400156u + (uint)(number - 1) * 10;
                var baselineWork = Work(record, workId); var baselineCost = record.Number(number == 0 ? 0x0c400011u : workId + 1) / 100m;
                if (baselineWork == null && baselineCost == null) continue;
                var baseline = resource.Baselines.Add(); baseline.Number = number; baseline.Work = baselineWork; baseline.Cost = baselineCost;
            }
        }
    }
    private static void ReadAssignments(ProjectDocument document, ProjectNativeTable table, ProjectLoadOptions options, CancellationToken token) {
        foreach (var record in table.Records) {
            token.ThrowIfCancellationRequested();
            record.Uid = record.Integer(0x0f400000) ?? throw new InvalidDataException("Native assignment UID missing.");
            CheckEntityBudget(document, options);
            int taskId = record.Integer(0x0f400001) ?? -1, resourceId = record.Integer(0x0f400002) ?? -1;
            document.TaskIndex.TryGetValue(taskId, out var task); document.ResourceIndex.TryGetValue(resourceId, out var resource);
            var assignment = new ProjectAssignment(document, record.Uid) {
                Task = task, Resource = resource, SourceTaskUid = taskId, SourceResourceUid = resourceId,
                Start = record.Date(0x0f400014), Finish = record.Date(0x0f400015),
                Work = Work(record, 0x0f400008), ActualWork = Work(record, 0x0f40000a), RemainingWork = Work(record, 0x0f40000c),
                OvertimeWork = Work(record, 0x0f400009), ActualOvertimeWork = Work(record, 0x0f40000d),
                ActualStart = record.Date(0x0f400016), ActualFinish = record.Date(0x0f400017), PercentWorkComplete = record.Integer(0x0f40002b),
                Cost = record.Number(0x0f40001a) / 100m, ActualCost = record.Number(0x0f40001c) / 100m, RemainingCost = record.Number(0x0f40001d) / 100m
            };
            decimal? units = record.Number(0x0f400007);
            if (units.HasValue) assignment.Units = ProjectUnits.Fraction(units.Value / 10000m);
            if (record.Value(0x0f40027c) is ProjectNativeValue guid) assignment.Guid = new Guid(guid.Copy());
            if (document.AssignmentIndex.ContainsKey(assignment.Uid)) throw new InvalidDataException("Duplicate native assignment UID.");
            document.AssignmentIndex.Add(assignment.Uid, assignment); document.Assignments.Items.Add(assignment);
            if (task != null && resource != null) document.AssignmentPairs.Add(ProjectDocument.PairKey(taskId, resourceId));
            for (int number = 0; number <= 10; number++) {
                uint workId = number == 0 ? 0x0f400010u : 0x0f400121u + (uint)(number - 1) * 9;
                var baselineWork = Work(record, workId); var baselineCost = record.Number(number == 0 ? 0x0f400020u : workId + 1) / 100m;
                var start = record.Date(number == 0 ? 0x0f400092u : workId + 6); var finish = record.Date(number == 0 ? 0x0f400093u : workId + 7);
                if (baselineWork == null && baselineCost == null && start == null && finish == null) continue;
                var baseline = assignment.Baselines.Add(); baseline.Number = number; baseline.Work = baselineWork; baseline.Cost = baselineCost; baseline.Start = start; baseline.Finish = finish;
            }
        }
    }
}
