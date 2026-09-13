using System.Xml.Linq;
using OfficeIMO.Core;

namespace OfficeIMO.Project.Tests;

public class SchemaContracts {
    [Fact]
    public void BaselineFieldsRespectTheirOwningRecord() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task");
        var resource = document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource);
        var taskBaseline = task.Baselines.Add(); taskBaseline.Number = 0;
        taskBaseline.Duration = ProjectDuration.WorkingDays(1); taskBaseline.FixedCost = 12.5m;
        var resourceBaseline = resource.Baselines.Add(); resourceBaseline.Number = 0; resourceBaseline.Work = new ProjectWork(60);
        var assignmentBaseline = assignment.Baselines.Add(); assignmentBaseline.Number = 0;
        assignmentBaseline.Start = new DateTime(2026, 10, 5, 8, 0, 0);
        using var copy = ProjectDocument.Parse(document.ToXml());
        Assert.Equal(12.5m, copy.Tasks[0].Baselines[0].FixedCost);
        resourceBaseline.Start = assignmentBaseline.Start;
        assignmentBaseline.Duration = ProjectDuration.WorkingDays(1);
        Assert.Equal(2, document.Validate().Diagnostics.Count(d => d.Code == "PROJECT_BASELINE_CONTEXT"));
        Assert.Throws<InvalidDataException>(() => document.ToXml());
    }

    [Theory]
    [InlineData(2.5, true, "3")]
    [InlineData(-2.5, true, "-3")]
    [InlineData(0.25, false, "3")]
    public void FractionalLagRequiresExplicitLossPermission(double value, bool percent, string expected) {
        using var document = ProjectDocument.Create();
        var link = document.Dependencies.Add(document.Tasks.Add("A"), document.Tasks.Add("B"));
        if (percent) link.LagPercent = (decimal)value;
        else link.Lag = new ProjectDuration((decimal)value, ProjectDurationUnit.Minute);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_LAG_PRECISION" && d.RepresentsLoss);
        Assert.Throws<InvalidOperationException>(() => document.ToXml());
        var xml = XDocument.Parse(document.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(expected, xml.Descendants(XName.Get("LinkLag", XmlContracts.Ns)).Single().Value);
    }

    [Theory]
    [InlineData("<UID>8</UID>")]
    [InlineData("<CalendarUID>1</CalendarUID><CalendarUID>2</CalendarUID>")]
    [InlineData("<Summary><value>1</value></Summary>")]
    public void ManuallyParsedScalarsRejectAmbiguousXml(string content) {
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(XmlContracts.Wrap(XmlContracts.TaskXml(content))));
    }
}
