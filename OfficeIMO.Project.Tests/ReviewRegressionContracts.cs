using System.Xml.Linq;

namespace OfficeIMO.Project.Tests;

public class ReviewRegressionContracts {
    [Fact]
    public void ProjectSummaryRowDoesNotAdoptNewRootTasksOnReload() {
        using var project = ProjectDocument.Parse(XmlContracts.Wrap("<Tasks><Task><UID>0</UID><OutlineLevel>0</OutlineLevel><Summary>1</Summary></Task></Tasks>"));
        var task = project.Tasks.Add("New root");
        Assert.Null(task.Parent);
        using var copy = ProjectDocument.Parse(project.ToXml());
        Assert.Null(copy.Tasks.GetByUid(task.Uid).Parent);
        Assert.Equal(2, copy.Tasks.Count);
        var summary = project.Tasks.GetByUid(0);
        Assert.Throws<InvalidOperationException>(() => summary.MoveTo(task));
        Assert.Throws<InvalidOperationException>(() => task.MoveTo(summary));
        Assert.Throws<InvalidOperationException>(() => summary.Children.Add("Invalid child"));
        Assert.Null(summary.Parent); Assert.Null(task.Parent); Assert.Empty(summary.Children);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CalendarBindingIsIndependentOfRecordOrderAndRejectsCycles(bool reverse) {
        var rows = Enumerable.Range(1, 1000).Select(i => "<Calendar><UID>" + i + "</UID><BaseCalendarUID>" + (i - 1) + "</BaseCalendarUID></Calendar>");
        using var document = ProjectDocument.Parse(XmlContracts.Wrap("<Calendars>" + string.Concat(reverse ? rows.Reverse() : rows) + "</Calendars>"));
        Assert.Same(document.Calendars.GetByUid(999), document.Calendars.GetByUid(1000).BaseCalendar);
        Assert.Null(document.Calendars.GetByUid(1).BaseCalendar);
        string cycle = "<Calendars><Calendar><UID>1</UID><BaseCalendarUID>2</BaseCalendarUID></Calendar><Calendar><UID>2</UID><BaseCalendarUID>1</BaseCalendarUID></Calendar></Calendars>";
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(XmlContracts.Wrap(cycle)));
    }

    [Fact]
    public void ScalarAndIntervalEditsKeepLegacyOnlyWorkingTimeExtensions() {
        const string period = "<TimePeriod><FromDate>2026-10-05T00:00:00</FromDate><ToDate>2026-10-05T00:00:00</ToDate></TimePeriod>";
        string working(string extra) => "<WorkingTimes><WorkingTime " + extra + "><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime>" +
            (extra.Length == 0 ? "" : "<x:Extra>retained</x:Extra>") + "</WorkingTime></WorkingTimes>";
        string source = XmlContracts.Wrap("<Calendars><Calendar><UID>1</UID><WeekDays><WeekDay><DayType>0</DayType><DayWorking>1</DayWorking>" + period +
            working("xmlns:x=\"urn:legacy\" x:keep=\"legacy\"") + "</WeekDay></WeekDays><Exceptions><Exception><DayWorking>1</DayWorking>" + period + working("") +
            "</Exception></Exceptions></Calendar></Calendars>");
        using var project = ProjectDocument.Parse(source);
        project.Name = "Unrelated edit";
        var output = XDocument.Parse(project.ToXml());
        Assert.Single(output.Descendants().Attributes(XName.Get("keep", "urn:legacy")));
        Assert.Single(output.Descendants(XName.Get("Extra", "urn:legacy")));
        project.Calendars[0].Exceptions[0].WorkingTimes[0].From = TimeSpan.FromHours(9);
        output = XDocument.Parse(project.ToXml());
        Assert.Single(output.Descendants().Attributes(XName.Get("keep", "urn:legacy")));
        Assert.Single(output.Descendants(XName.Get("Extra", "urn:legacy")));
        Assert.Equal(2, output.Descendants(XName.Get("FromTime", XmlContracts.Ns)).Count(e => e.Value == "09:00:00"));
    }
}
