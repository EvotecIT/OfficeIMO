using OfficeIMO;
using OfficeIMO.Project;
using Xunit;

namespace OfficeIMO.Project.Tests;

public class ProjectNativeRepeatedEditTests {
    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void DependencyFieldEditsKeepTheLogicalProjectAndFileSizeStable(ProjectFileFormat format) {
        byte[] bytes = Create(format); long originalLength = bytes.Length;
        for (int edit = 0; edit < 30; edit++) {
            using var document = ProjectDocument.Load(new MemoryStream(bytes));
            document.Dependencies[0].Lag = ProjectDuration.WorkingMinutes(edit % 2);
            bytes = Save(document, format);
            Assert.Equal(originalLength, bytes.LongLength);
            using var reopened = ProjectDocument.Load(new MemoryStream(bytes));
            Assert.Equal(3, reopened.AllTasks.Count()); Assert.Equal(2, reopened.Dependencies.Count);
            Assert.Equal(edit % 2, reopened.Dependencies[0].Lag!.Value.Value);
        }
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void DependencyAddRemoveCyclesDoNotAccumulateDiscardedTailRecords(ProjectFileFormat format) {
        byte[] bytes = Create(format); long originalLength = bytes.Length;
        for (int edit = 0; edit < 20; edit++) {
            using (var document = ProjectDocument.Load(new MemoryStream(bytes))) {
                document.Dependencies.Remove(document.Dependencies.Single(d => d.Successor.Uid == 3)); bytes = Save(document, format);
            }
            using (var document = ProjectDocument.Load(new MemoryStream(bytes))) {
                document.Dependencies.Add(document.Tasks.GetByUid(1), document.Tasks.GetByUid(3)); bytes = Save(document, format);
            }
            Assert.Equal(originalLength, bytes.LongLength);
            using var reopened = ProjectDocument.Load(new MemoryStream(bytes));
            Assert.Equal(new[] { (1, 2), (1, 3) }, reopened.Dependencies.Select(d => (d.Predecessor!.Uid, d.Successor.Uid)).OrderBy(p => p));
        }
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ReplacingADependencyReusesOnlyItsDeletedRecord(ProjectFileFormat format) {
        byte[] bytes = Create(format); long originalLength = bytes.Length;
        for (int edit = 0; edit < 20; edit++) {
            using var document = ProjectDocument.Load(new MemoryStream(bytes));
            document.Dependencies.Remove(document.Dependencies.Single(d => d.Predecessor!.Uid != 1 || d.Successor.Uid != 3));
            if (edit % 2 == 0) document.Dependencies.Add(document.Tasks.GetByUid(2), document.Tasks.GetByUid(3));
            else document.Dependencies.Add(document.Tasks.GetByUid(1), document.Tasks.GetByUid(2));
            bytes = Save(document, format); Assert.Equal(originalLength, bytes.LongLength);
            using var reopened = ProjectDocument.Load(new MemoryStream(bytes));
            Assert.Contains(reopened.Dependencies, d => d.Predecessor!.Uid == 1 && d.Successor.Uid == 3);
            Assert.Contains(reopened.Dependencies, d => d.Predecessor!.Uid == (edit % 2 == 0 ? 2 : 1) && d.Successor.Uid == (edit % 2 == 0 ? 3 : 2));
        }
    }

    private static byte[] Create(ProjectFileFormat format) {
        using var document = ProjectDocument.Create();
        var start = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = start;
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        for (int index = 0; index < 3; index++) {
            var task = document.Tasks.Add("Task " + index); task.Duration = ProjectDuration.WorkingMinutes(60); task.RemainingDuration = task.Duration;
            task.Start = start; task.Finish = start.AddHours(1);
        }
        document.Dependencies.Add(document.Tasks.GetByUid(1), document.Tasks.GetByUid(2));
        document.Dependencies.Add(document.Tasks.GetByUid(1), document.Tasks.GetByUid(3));
        return Save(document, format);
    }
    private static byte[] Save(ProjectDocument document, ProjectFileFormat format) {
        using var output = new MemoryStream();
        document.Save(output, new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow });
        return output.ToArray();
    }
}
