namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    private static void ReadTasks(ProjectDocument document, ProjectNativeTable table, ProjectLoadOptions options, CancellationToken token) {
        var ordered = new List<ProjectTask>();
        foreach (var record in table.Records) {
            token.ThrowIfCancellationRequested();
            record.Uid = record.Integer(0x0b400056) ?? throw new InvalidDataException("A task has no native UID.");
            if (record.Uid < 0) continue;
            CheckEntityBudget(document, options);
            if (ordered.Count >= options.MaxTasks) throw new InvalidDataException("Native task budget exceeded.");
            var task = new ProjectTask(document, record.Uid) {
                DisplayId = record.Integer(0x0b400017), Name = record.Text(0x0b40000e),
                SourceOutlineLevel = record.Integer(0x0b4000f9), SourceSummary = record.Boolean(0x0b40005c) == true,
                Duration = Duration(record, 0x0b40001d, 0x0b4000b5, document),
                ActualDuration = Duration(record, 0x0b40001c, 0x0b4000b5, document),
                RemainingDuration = Duration(record, 0x0b40001f, 0x0b4000b5, document),
                Start = record.Date(0x0b400023), Finish = record.Date(0x0b400024),
                EarlyStart = record.Date(0x0b400025), EarlyFinish = record.Date(0x0b400026), LateStart = record.Date(0x0b400027), LateFinish = record.Date(0x0b400028),
                FreeSlackMinutes = record.Integer(0x0b400015) / 10m, TotalSlackMinutes = record.Integer(0x0b400016) / 10m,
                ActualStart = record.Date(0x0b400029), ActualFinish = record.Date(0x0b40002a),
                Deadline = record.Date(0x0b4001b5), ConstraintDate = record.Date(0x0b400012),
                ConstraintType = (ProjectConstraintType?)record.Integer(0x0b400011), Type = (ProjectTaskType?)record.Integer(0x0b400080),
                IsManual = record.Boolean(0x0b400500), IsActive = record.Boolean(0x0b4004ff), IsMilestone = record.Boolean(0x0b400018),
                IsCritical = record.Boolean(0x0b400013), EffortDriven = record.Boolean(0x0b400084),
                Work = Work(record, 0x0b400000), ActualWork = Work(record, 0x0b400002), RemainingWork = Work(record, 0x0b400004),
                Cost = record.Number(0x0b400005) / 100m, ActualCost = record.Number(0x0b400007) / 100m,
                RemainingCost = record.Number(0x0b40000a) / 100m, FixedCost = record.Number(0x0b400008) / 100m,
                PercentComplete = record.Integer(0x0b400020), PercentWorkComplete = record.Integer(0x0b400021), Priority = record.Integer(0x0b400019)
            };
            if (record.Value(0x0b400477) is ProjectNativeValue guid) task.Guid = new Guid(guid.Copy());
            int calendarId = record.Integer(0x0b400191) ?? -1;
            if (document.CalendarIndex.TryGetValue(calendarId, out var calendar)) task.Calendar = calendar;
            else if (calendarId > 0) task.SourceCalendarUid = calendarId;
            if (record.Value(0x0b40000f).HasValue) Warn(document, "PROJECT_NATIVE_RTF_NOTES", "RTF task notes remain in the native source. Plain-text conversion requires the RTF document owner.", "/Task[UID=" + task.Uid + "]/Notes");
            ReadTaskBaseline(document, record, task);
            ReadCustomValues(document, record, task.CustomFields, true, token);
            if (document.TaskIndex.ContainsKey(task.Uid)) throw new InvalidDataException("Duplicate native task UID.");
            document.TaskIndex.Add(task.Uid, task); ordered.Add(task);
        }
        ordered.Sort((left, right) => (left.DisplayId ?? int.MaxValue).CompareTo(right.DisplayId ?? int.MaxValue));
        var parents = new Stack<ProjectTask>();
        foreach (var task in ordered) {
            int level = task.SourceOutlineLevel ?? throw new InvalidDataException("Native task outline level is absent.");
            if (level < 0 || level > options.MaxOutlineDepth) throw new InvalidDataException("Native outline depth exceeds the configured limit.");
            while (parents.Count != 0 && parents.Peek().SourceOutlineLevel >= level) parents.Pop();
            // A project summary remains a separate row; ordinary top-level tasks are not its structural children.
            var parent = parents.Count != 0 && parents.Peek().SourceOutlineLevel > 0 ? parents.Peek() : null;
            if (level > 1 && (parent == null || parent.SourceOutlineLevel != level - 1)) throw new InvalidDataException("Native task outline skips a parent level.");
            task.Parent = parent;
            (parent?.Children ?? document.Tasks).Items.Add(task);
            parents.Push(task);
        }
    }

    private static void ReadTaskBaseline(ProjectDocument document, ProjectNativeRecord record, ProjectTask task) {
        var start = record.Date(0x0b40002b); var finish = record.Date(0x0b40002c); var work = Work(record, 0x0b400001);
        var cost = record.Number(0x0b400006) / 100m;
        if (start != null || finish != null || work != null || cost != null) {
            var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Start = start; baseline.Finish = finish;
            baseline.Work = work; baseline.Cost = cost; baseline.Duration = Duration(record, 0x0b40001b, 0x0b4000b3, document);
        }
        var firstFields = new uint[] { 0x1e2, 0x1ed, 0x1f8, 0x203, 0x20e, 0x220, 0x22b, 0x236, 0x241, 0x24c };
        for (int index = 0; index < firstFields.Length; index++) {
            uint first = firstFields[index];
            uint id = 0x0b400000 | first;
            var extraStart = record.Date(id); var extraFinish = record.Date(id + 1);
            var extraWork = Work(record, id + 3); var extraCost = record.Number(id + 2) / 100m;
            if (extraStart == null && extraFinish == null && extraWork == null && extraCost == null) continue;
            var item = task.Baselines.Add(); item.Number = index + 1;
            item.Start = extraStart; item.Finish = extraFinish; item.Work = extraWork; item.Cost = extraCost;
            item.Duration = Duration(record, id + 5, id + 6, document);
        }
    }

    private static void ReadDependencies(ProjectDocument document, ProjectNativeTable table, CancellationToken token) {
        foreach (var record in table.Records) {
            token.ThrowIfCancellationRequested();
            record.Uid = record.Integer(0x0e400000) ?? throw new InvalidDataException("Native dependency UID missing.");
            int predecessorId = record.Integer(0x0e400002) ?? -1, successorId = record.Integer(0x0e400005) ?? -1;
            if (!document.TaskIndex.TryGetValue(successorId, out var successor)) throw new InvalidDataException("Native dependency successor is absent.");
            document.TaskIndex.TryGetValue(predecessorId, out var predecessor);
            var link = new ProjectDependency(document) { Predecessor = predecessor, Successor = successor, SourcePredecessorUid = predecessorId,
                Type = (ProjectDependencyType?)record.Integer(0x0e400007) };
            int format = record.Integer(0x0e40000a) ?? 7;
            if (format == 19 || format == 20 || format == 51 || format == 52) link.LagPercent = record.Integer(0x0e400009);
            else link.Lag = Duration(record, 0x0e400009, 0x0e40000a, document);
            document.Dependencies.Items.Add(link);
            if (predecessor != null) document.DependencyPairs.Add(ProjectDocument.PairKey(predecessorId, successorId));
        }
    }
}
