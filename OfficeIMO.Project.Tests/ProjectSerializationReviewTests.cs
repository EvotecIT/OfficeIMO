namespace OfficeIMO.Project.Tests;

public sealed class ProjectSerializationReviewTests {
    [Theory]
    [InlineData(ProjectFileFormat.Mpx4)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void XmlOutputNeverRetainsPreviouslySavedNonXmlBytes(ProjectFileFormat sourceFormat) {
        using var document = sourceFormat == ProjectFileFormat.Mpp14
            ? ProjectNativeAuthoringTests.Create()
            : ProjectDocument.Create();
        if (sourceFormat == ProjectFileFormat.Mpx4) {
            var task = document.Tasks.Add("Task"); task.DisplayId = 1;
        }
        var allowSource = new ProjectSaveOptions { Format = sourceFormat, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var source = new MemoryStream(); document.Save(source, allowSource);
        Assert.NotEmpty(source.ToArray());

        var allowXml = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow };
        string xml = document.ToXml(allowXml);
        Assert.Equal("Project", System.Xml.Linq.XDocument.Parse(xml).Root!.Name.LocalName);

        string path = Path.Combine(Path.GetTempPath(), "officeimo-project-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            document.Save(path, allowXml);
            Assert.Equal("Project", System.Xml.Linq.XDocument.Load(path).Root!.Name.LocalName);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void RetainedMpxSourceKeepsRootConversionLosses() {
        const string xml = "<Project xmlns=\"http://schemas.microsoft.com/project\" xmlns:x=\"urn:fixture\"><x:Opaque>value</x:Opaque></Project>";
        using var document = ProjectDocument.Parse(xml);
        var allow = new ProjectSaveOptions { Format = ProjectFileFormat.Mpx4, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var output = new MemoryStream(); document.Save(output, allow);

        var strict = new ProjectSaveOptions { Format = ProjectFileFormat.Mpx4 };
        var diagnostic = Assert.Single(document.AssessSave(strict).Diagnostics,
            item => item.Code == "PROJECT_MPX_XML_CONTENT_LOSS" && item.Location == "/");
        Assert.True(diagnostic.RepresentsLoss);
        Assert.Throws<InvalidOperationException>(() => document.Save(new MemoryStream(), strict));
    }

    [Fact]
    public void RetainedNativeSourceKeepsRootGenerationLosses() {
        using var seed = ProjectNativeAuthoringTests.Create(); using var source = new MemoryStream();
        seed.Save(source, new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14 });
        using var document = ProjectDocument.Load(new MemoryStream(source.ToArray()));
        var allow = new ProjectSaveOptions { Format = ProjectFileFormat.Mpp8, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var converted = new MemoryStream(); document.Save(converted, allow);

        var strict = new ProjectSaveOptions { Format = ProjectFileFormat.Mpp8 };
        var diagnostic = Assert.Single(document.AssessSave(strict).Diagnostics,
            item => item.Code == "PROJECT_NATIVE_GENERATION_LOSS" && item.Location == "/");
        Assert.True(diagnostic.RepresentsLoss);
        Assert.Throws<InvalidOperationException>(() => document.Save(new MemoryStream(), strict));
    }

    [Fact]
    public void NativeToXmlReportsOmittedPresentationAsOmission() {
        using var seed = ProjectNativeAuthoringTests.Create();
        using var source = new MemoryStream();
        seed.Save(source, new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14 });
        using var document = ProjectDocument.Load(new MemoryStream(source.ToArray()));

        ProjectReport report = document.AssessSave(new ProjectSaveOptions { Format = ProjectFileFormat.Xml });

        ProjectDiagnostic diagnostic = Assert.Single(report.Diagnostics,
            item => item.Code == "PROJECT_NATIVE_PRESENTATION_LOSS");
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Equal(OfficeConversionLossKind.Omission, Assert.Single(report.FidelityDiagnostics,
            item => item.Code == diagnostic.Code).LossKind);
    }

    [Theory]
    [InlineData("80,external-reference", "PROJECT_MPX_CONVERSION_LOSS", "/MPX/Record[7]")]
    [InlineData("72,recurrence", "PROJECT_MPX_CONVERSION_LOSS", "/Task[UID=1]/Recurrence")]
    [InlineData("0,comment", "PROJECT_MPX_COMMENT_LOSS", "/")]
    public void RetainedNativeSourceKeepsMpxConversionLosses(string sourceRecord, string expectedCode, string expectedLocation) {
        string mpx = "MPX,Fixture,4.0,ANSI\r\n10,$,1,2,\",\",.\r\n12,1,0,480,/,:,AM,PM,20,9\r\n"
            + "30,Fixture,,,Standard\r\n61,90,1\r\n70,1,Task\r\n" + sourceRecord + "\r\n";
        using var document = ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(mpx)));
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var allow = new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var converted = new MemoryStream(); document.Save(converted, allow);

        var strict = new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14 };
        var diagnostic = Assert.Single(document.AssessSave(strict).Diagnostics,
            item => item.Code == expectedCode && item.Location == expectedLocation);
        Assert.True(diagnostic.RepresentsLoss);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Equal(OfficeConversionLossKind.Omission, Assert.Single(document.AssessSave(strict).FidelityDiagnostics,
            item => item.Code == expectedCode && item.Location == expectedLocation).LossKind);
        Assert.Throws<InvalidOperationException>(() => document.Save(new MemoryStream(), strict));

        document.Name = "Edited";
        using var rewritten = new MemoryStream(); document.Save(rewritten, allow);
        Assert.Contains(document.AssessSave(strict).Diagnostics,
            item => item.Code == expectedCode && item.Location == expectedLocation && item.RepresentsLoss);
    }

    [Fact]
    public void FreshXmlExplicitTaskDisplayRowsRequireLossAcceptance() {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Task"); task.DisplayId = 99;
        var strict = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        var diagnostic = Assert.Single(document.AssessSave(strict).Diagnostics,
            item => item.Code == "PROJECT_XML_TASK_ROWS");
        Assert.Equal("/Task[UID=" + task.Uid + "]/DisplayId", diagnostic.Location);
        Assert.True(diagnostic.RepresentsLoss); Assert.Contains("becomes 1", diagnostic.Message);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<InvalidOperationException>(() => document.Save(new MemoryStream(), strict));

        var allow = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var output = new MemoryStream(); document.Save(output, allow);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(1, reopened.Tasks.Single().DisplayId);
    }

    [Fact]
    public void FreshXmlTasksWithoutDisplayRowsRemainAbsent() {
        using var document = ProjectDocument.Create(); document.Tasks.Add("Task");
        var strict = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        Assert.DoesNotContain(document.AssessSave(strict).Diagnostics, item => item.Code == "PROJECT_XML_TASK_ROWS");
        using var output = new MemoryStream(); document.Save(output, strict);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Tasks.Single().DisplayId);
    }

    [Fact]
    public void StructuralXmlEditsReportEveryChangedTaskDisplayRow() {
        const string xml = "<Project xmlns=\"http://schemas.microsoft.com/project\"><Tasks>"
            + "<Task><UID>1</UID><ID>10</ID><Name>First</Name></Task>"
            + "<Task><UID>2</UID><ID>20</ID><Name>Second</Name></Task></Tasks></Project>";
        using var document = ProjectDocument.Parse(xml); document.Tasks.GetByUid(2).MoveTo(null, 0);
        var report = document.AssessSave(new ProjectSaveOptions { Format = ProjectFileFormat.Xml });
        Assert.Contains(report.Diagnostics, item => item.Code == "PROJECT_XML_TASK_ROWS" && item.Location == "/Task[UID=2]/DisplayId");
        Assert.Contains(report.Diagnostics, item => item.Code == "PROJECT_XML_TASK_ROWS" && item.Location == "/Task[UID=1]/DisplayId");
        var allow = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var output = new MemoryStream(); document.Save(output, allow);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { 2, 1 }, reopened.Tasks.Select(item => item.Uid));
        Assert.Equal(new int?[] { 1, 2 }, reopened.Tasks.Select(item => item.DisplayId));
    }

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
