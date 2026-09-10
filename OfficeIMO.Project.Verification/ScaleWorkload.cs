using OfficeIMO.Project;
using System.Diagnostics;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

internal static class ScaleWorkload {
    internal static int Run(string operation, string path, int count, string shape) {
        if (count < 1 || count > 100_000) throw new ArgumentOutOfRangeException(nameof(count));
        if (shape != "flat" && shape != "deep" && shape != "dense" && shape != "timephased" && shape != "calendar-chain" && shape != "calendar-mirrors") throw new ArgumentException("Unknown workload shape.");
        if (operation == "scale-create") {
            using var document = ProjectDocument.Create();
            document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
            document.Calendar = document.Calendars.AddStandardWorkingWeek();
            var resource = document.Resources.AddWork("Engineer");
            var tasks = new List<ProjectTask>();
            for (int i = 0; i < count; i++) {
                var collection = shape == "deep" && tasks.Count != 0 ? tasks[tasks.Count - 1].Children : document.Tasks;
                var task = collection.Add("Task " + (i + 1));
                task.Duration = ProjectDuration.WorkingHours(1);
                task.Work = ProjectWork.Hours(1);
                var assignment = document.Assignments.Add(task, resource);
                assignment.Work = ProjectWork.Hours(1);
                if (shape == "dense") for (int preceding = Math.Max(0, i - 4); preceding < i; preceding++) document.Dependencies.Add(tasks[preceding], task);
                if (shape == "timephased") for (int interval = 0; interval < 10; interval++) {
                    var value = assignment.TimephasedData.Add();
                    value.Type = 1; value.Uid = assignment.Uid; value.Unit = 2;
                    value.Start = new DateTime(2026, 10, 5, 8, 0, 0).AddHours(interval);
                    value.Finish = value.Start.Value.AddHours(1); value.Value = "PT0H6M0S";
                }
                tasks.Add(task);
            }
            if (shape == "calendar-chain") {
                var calendars = Enumerable.Range(0, count).Select(i => document.Calendars.Add("Calendar " + i)).ToArray();
                calendars[0].IsBaseCalendar = true;
                for (int i = count - 1; i > 0; i--) { calendars[i].IsBaseCalendar = false; calendars[i].BaseCalendar = calendars[i - 1]; }
            }
            if (shape == "calendar-mirrors") {
                for (int i = 0; i < count; i++) {
                    var day = document.Calendar!.WeekDays.Add(); day.FromDate = new DateTime(2026, 10, 5).AddDays(i); day.ToDate = day.FromDate; day.IsWorking = true;
                    var legacy = day.WorkingTimes.Add(); legacy.From = TimeSpan.FromHours(8); legacy.To = TimeSpan.FromHours(12);
                    var exception = document.Calendar.Exceptions.Add(); exception.FromDate = day.FromDate; exception.ToDate = day.ToDate; exception.IsWorking = true;
                    var modern = exception.WorkingTimes.Add(); modern.From = legacy.From; modern.To = legacy.To;
                }
            }
            document.Save(path, new ProjectSaveOptions { Indent = false });
            Console.WriteLine(new FileInfo(path).Length);
            return 0;
        }
        long before = GC.GetAllocatedBytesForCurrentThread();
        using var output = new MemoryStream();
        using (var loaded = ProjectDocument.Load(path, new ProjectLoadOptions { MaxOutlineDepth = shape == "deep" ? count : 128 })) {
            loaded.Validate().ThrowIfErrors();
            if (shape == "calendar-chain" && (loaded.Calendars.Count != count + 1 || loaded.Calendars.Last().BaseCalendar?.Uid != count))
                throw new InvalidDataException("Calendar chain reference binding failed.");
            if (shape == "calendar-mirrors" && (loaded.Calendar!.Exceptions.Count != count || loaded.Calendar.WeekDays.Count != 7))
                throw new InvalidDataException("Calendar mirrors were not consolidated.");
            loaded.Tasks.GetByUid(count).Notes = "Measured scalar edit";
            loaded.Save(output, new ProjectSaveOptions { Indent = false });
        }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        output.Position = 0;
        var uids = new HashSet<int>();
        int taskCount = 0, assignmentCount = 0, dependencies = 0, timephased = 0;
        long uidSum = 0;
        decimal workMinutes = 0;
        string? editedNote = null;
        XNamespace ns = "http://schemas.microsoft.com/project";
        using var reader = XmlReader.Create(output, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, CloseInput = false });
        while (reader.Read()) {
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != ns.NamespaceName) continue;
            if (reader.LocalName != "Task" && reader.LocalName != "Assignment") continue;
            using var subtree = reader.ReadSubtree();
            var record = XElement.Load(subtree);
            if (record.Name.LocalName == "Task") {
                taskCount++;
                int uid = (int)record.Element(ns + "UID")!;
                if (!uids.Add(uid)) throw new InvalidDataException("Duplicate serialized task UID.");
                uidSum += uid;
                workMinutes += (decimal)XmlConvert.ToTimeSpan((string)record.Element(ns + "Work")!).TotalMinutes;
                dependencies += record.Elements(ns + "PredecessorLink").Count();
                if (uid == count) editedNote = (string?)record.Element(ns + "Notes");
            } else {
                assignmentCount++;
                timephased += record.Elements(ns + "TimephasedData").Count();
            }
        }
        if (taskCount != count || uidSum != (long)count * (count + 1) / 2 ||
            assignmentCount != count || editedNote != "Measured scalar edit" || workMinutes != count * 60m ||
            (shape == "dense" && dependencies != (count >= 4 ? (long)count * 4 - 10 : (long)count * (count - 1) / 2)) || (shape == "timephased" && timephased != count * 10))
            throw new InvalidDataException("Workload readback failed identity/work/count proof.");
        Console.WriteLine(JsonSerializer.Serialize(new {
            tasks = taskCount, assignments = assignmentCount, uidSum, dependencies, timephased,
            inputBytes = new FileInfo(path).Length, outputBytes = output.Length, allocatedBytes = allocated,
            peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
            runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription, shape
        }));
        return 0;
    }
}
