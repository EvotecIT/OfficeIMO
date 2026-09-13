namespace OfficeIMO.Project.Tests;

public sealed class ProjectSerializationReviewTests {
    [Fact]
    public void NewXmlTaskRemainingDurationDefaultRequiresLossAcceptance() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingHours(8);
        var strict = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };

        var diagnostic = Assert.Single(document.AssessSave(strict).Diagnostics,
            item => item.Code == "PROJECT_XML_REMAINING_DURATION_DEFAULT");
        Assert.True(diagnostic.RepresentsLoss);
        Assert.Equal("/Task[UID=" + task.Uid + "]/RemainingDuration", diagnostic.Location);
        using var rejected = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => document.Save(rejected, strict)); Assert.Equal(0, rejected.Length);

        var allow = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var output = new MemoryStream(); document.Save(output, allow);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(task.Duration, reopened.Tasks.GetByUid(task.Uid).RemainingDuration);
    }

    [Fact]
    public void ConvertedMpxAndNativeTasksReportXmlRemainingDurationDefaults() {
        const string mpx = "MPX,Fixture,4.0,ANSI\r\n61,90,98,1,40\r\n70,1,1,Task,1d\r\n";
        using var mpxDocument = ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(mpx)));
        mpxDocument.Settings.CurrencyCode = "USD";
        var mpxTask = Assert.Single(mpxDocument.Tasks); Assert.Null(mpxTask.RemainingDuration);
        Assert.Contains(mpxDocument.AssessSave(new ProjectSaveOptions { Format = ProjectFileFormat.Xml }).Diagnostics,
            item => item.Code == "PROJECT_XML_REMAINING_DURATION_DEFAULT" && item.Location == "/Task[UID=" + mpxTask.Uid + "]/RemainingDuration");

        using var seed = ProjectNativeAuthoringTests.Create(); using var nativeBytes = new MemoryStream();
        seed.Save(nativeBytes, new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14 });
        using var nativeDocument = ProjectDocument.Load(new MemoryStream(nativeBytes.ToArray()));
        var nativeTask = Assert.Single(nativeDocument.AllTasks, item => item.Duration.HasValue && !item.RemainingDuration.HasValue);
        Assert.Contains(nativeDocument.AssessSave(new ProjectSaveOptions { Format = ProjectFileFormat.Xml }).Diagnostics,
            item => item.Code == "PROJECT_XML_REMAINING_DURATION_DEFAULT" && item.Location == "/Task[UID=" + nativeTask.Uid + "]/RemainingDuration");
    }

    [Theory]
    [InlineData("project")]
    [InlineData("settings")]
    [InlineData("task")]
    [InlineData("resource")]
    [InlineData("assignment")]
    [InlineData("calendar")]
    [InlineData("workweek")]
    [InlineData("dependency")]
    [InlineData("definition")]
    [InlineData("lookup")]
    [InlineData("outline")]
    [InlineData("timephased")]
    public void XmlForbiddenCharactersBecomeLocatedAssessmentErrors(string owner) {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Task");
        var resource = document.Resources.AddWork("Engineer"); var assignment = document.Assignments.Add(task, resource);
        const string invalid = "bad\u0001text";
        string expectedLocation;
        switch (owner) {
            case "project": document.Name = invalid; expectedLocation = "/Project/Name"; break;
            case "settings": document.Settings.CurrencySymbol = invalid; expectedLocation = "/Settings/CurrencySymbol"; break;
            case "task": task.Name = invalid; expectedLocation = "/Task[UID=" + task.Uid + "]/Name"; break;
            case "resource": resource.Notes = invalid; expectedLocation = "/Resource[UID=" + resource.Uid + "]/Notes"; break;
            case "assignment": assignment.Notes = invalid; expectedLocation = "/Assignment[UID=" + assignment.Uid + "]/Notes"; break;
            case "calendar": {
                var calendar = document.Calendars.Add("Calendar"); var item = calendar.Exceptions.Add(); item.Name = invalid;
                expectedLocation = "/Calendar[UID=" + calendar.Uid + "]/Exception[0]/Name"; break;
            }
            case "workweek": {
                var calendar = document.Calendars.Add("Calendar"); var item = calendar.WorkWeeks.Add(); item.Name = invalid;
                expectedLocation = "/Calendar[UID=" + calendar.Uid + "]/Week[0]/Name"; break;
            }
            case "dependency": {
                var successor = document.Tasks.Add("Successor"); var dependency = document.Dependencies.Add(task, successor); dependency.CrossProjectName = invalid;
                expectedLocation = "/Dependency[0]/CrossProjectName"; break;
            }
            case "definition": {
                var definition = document.CustomFields.Add(); definition.FieldId = "188743731"; definition.Alias = invalid;
                expectedLocation = "/Definition[0]/Alias"; break;
            }
            case "lookup": {
                var definition = document.CustomFields.Add(); definition.FieldId = "188743731"; var value = definition.LookupValues.Add(); value.Value = invalid;
                expectedLocation = "/Definition[0]/Lookup[0]/Value"; break;
            }
            case "outline": {
                var definition = document.OutlineCodes.Add(); definition.Alias = invalid;
                expectedLocation = "/OutlineCode[0]/Alias"; break;
            }
            default: {
                var value = task.TimephasedData.Add(); value.Uid = task.Uid; value.Value = invalid;
                expectedLocation = "/Task[UID=" + task.Uid + "]/Timephased[0]/Value"; break;
            }
        }

        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        var diagnostic = Assert.Single(document.AssessSave(options).Diagnostics,
            item => item.Code == "PROJECT_XML_TEXT" && item.Location == expectedLocation);
        Assert.Equal(ProjectDiagnosticSeverity.Error, diagnostic.Severity);
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(output, options)); Assert.Equal(0, output.Length);
    }
}
