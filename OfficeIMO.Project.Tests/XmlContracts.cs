using System.Xml.Linq;

namespace OfficeIMO.Project.Tests;

public class XmlContracts {
    internal const string Ns = "http://schemas.microsoft.com/project";
    internal static string Wrap(string content) => "<Project xmlns=\"" + Ns + "\">" + content + "</Project>";
    internal static string TaskXml(string content = "", int uid = 7) => "<Tasks><Task><UID>" + uid + "</UID><ID>42</ID><Name>Original</Name><OutlineLevel>1</OutlineLevel>" + content + "</Task></Tasks>";

    [Fact]
    public void FieldEditsKeepExtensionNodesBetweenTheirOriginalRecords() {
        string source = Wrap("<Tasks marker=\"keep\"><Task><UID>1</UID><Name>A</Name></Task>" +
            "<!--between--><x:extension xmlns:x=\"urn:fixture\">in place</x:extension>" +
            "<Task><UID>2</UID><Name>B</Name></Task></Tasks>");
        using var document = ProjectDocument.Parse(source);
        document.Tasks[0].Name = "Edited";
        var output = XDocument.Parse(document.ToXml());
        var tasks = output.Root!.Element(XName.Get("Tasks", Ns))!;
        Assert.Equal(new[] { "Task", "extension", "Task" }, tasks.Elements().Select(e => e.Name.LocalName));
        Assert.Equal("keep", (string?)tasks.Attribute("marker"));
        Assert.Equal("between", tasks.Nodes().OfType<XComment>().Single().Value);
        Assert.Equal("Edited", tasks.Elements().First().Element(XName.Get("Name", Ns))!.Value);
    }

    [Fact]
    public void NoOpSavePreservesBytesAndSubsequentSavesKeepEdits() {
        byte[] original = new UnicodeEncoding(false, true).GetPreamble().Concat(Encoding.Unicode.GetBytes(
            "<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n" + Wrap(TaskXml("<Notes>café Łódź 日本語</Notes>")))).ToArray();
        using var input = new MemoryStream(original);
        input.Position = 17;
        using var document = ProjectDocument.Load(input);
        Assert.Equal(17, input.Position);
        using var output = new MemoryStream();
        document.Save(output);
        Assert.Equal(original, output.ToArray());
        document.Tasks[0].Name = "Edited";
        document.Save(output);
        var edited = output.ToArray();
        document.Save(output);
        Assert.Equal(edited, output.ToArray());
        using var copy = ProjectDocument.Load(output);
        Assert.Equal("Edited", copy.Tasks[0].Name);
        Assert.Equal("café Łódź 日本語", copy.Tasks[0].Notes);
        Assert.Equal("42", XDocument.Parse(copy.ToXml()).Descendants(XName.Get("Task", Ns)).Single().Element(XName.Get("ID", Ns))!.Value);
    }

    [Fact]
    public void FieldEditsPreserveForeignContentAndStrictStructuralSaveDiagnosesRisk() {
        string xml = Wrap("<x:Metadata xmlns:x=\"urn:fixture\" x:mode=\"keep\"><x:TaskRef>7</x:TaskRef></x:Metadata>" +
            TaskXml("<FutureField><TaskUID>7</TaskUID></FutureField><Notes custom=\"retained\">note</Notes>"));
        using var document = ProjectDocument.Parse(xml);
        Assert.Contains(document.ReadDiagnostics, d => d.Code == "PROJECT_XML_PRESERVED");
        document.Tasks[0].Name = "Rename";
        var changed = XDocument.Parse(document.ToXml());
        Assert.Equal("7", changed.Descendants(XName.Get("TaskRef", "urn:fixture")).Single().Value);
        Assert.Equal("retained", changed.Descendants(XName.Get("Notes", Ns)).Single().Attribute("custom")!.Value);
        document.Tasks.Add("New task");
        Assert.True(document.AssessSave().HasLoss);
        Assert.Throws<InvalidOperationException>(() => document.ToXml());
        Assert.Contains("FutureField", document.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    [Fact]
    public void AbsentFieldsStayAbsentAndExplicitNullRemovesSourceField() {
        using var document = ProjectDocument.Parse(Wrap(TaskXml("<Cost>0</Cost><Notes>remove me</Notes>")));
        Assert.Null(document.Tasks[0].Duration);
        Assert.Equal(0m, document.Tasks[0].Cost);
        document.Tasks[0].Notes = null;
        var xml = XDocument.Parse(document.ToXml());
        var task = xml.Descendants(XName.Get("Task", Ns)).Single();
        Assert.Null(task.Element(XName.Get("Notes", Ns)));
        Assert.Null(task.Element(XName.Get("Duration", Ns)));
        Assert.Equal("0", task.Element(XName.Get("Cost", Ns))!.Value);
    }

    [Fact]
    public void WorkingElapsedDurationWorkAndMoneyUseProjectWireConventions() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Working");
        task.Duration = ProjectDuration.WorkingDays(5);
        task.Work = ProjectWork.Hours(40);
        task.Cost = 5000.25m;
        var elapsed = document.Tasks.Add("Elapsed");
        elapsed.Duration = ProjectDuration.ElapsedDays(2).Estimated();
        var xml = XDocument.Parse(document.ToXml());
        var tasks = xml.Descendants(XName.Get("Task", Ns)).ToArray();
        Assert.Equal("PT40H0M0S", tasks[0].Element(XName.Get("Work", Ns))!.Value);
        Assert.Equal("PT40H0M0S", tasks[0].Element(XName.Get("RemainingDuration", Ns))!.Value);
        Assert.Equal(500025m, (decimal)tasks[0].Element(XName.Get("Cost", Ns))!);
        Assert.Equal("PT48H0M0S", tasks[1].Element(XName.Get("Duration", Ns))!.Value);
        using var read = ProjectDocument.Parse(xml.ToString());
        Assert.Equal(task.Duration, read.Tasks[0].Duration);
        Assert.Equal(task.Work, read.Tasks[0].Work);
        Assert.Equal(task.Cost, read.Tasks[0].Cost);
        Assert.Equal(elapsed.Duration, read.Tasks[1].Duration);
    }

    [Fact]
    public void DtdMalformedIdentitiesAndConfiguredResourceLimitsRejectBeforeOutput() {
        Assert.ThrowsAny<Exception>(() => ProjectDocument.Parse("<!DOCTYPE Project [<!ENTITY x SYSTEM 'file:///does-not-exist'>]>" + Wrap("<Name>&x;</Name>")));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(Wrap("<Tasks><Task><UID>1</UID></Task><Task><UID>1</UID></Task></Tasks>")));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(Wrap(TaskXml()), new ProjectLoadOptions { MaxInputBytes = 8 }));
        Assert.ThrowsAny<Exception>(() => ProjectDocument.Parse(Wrap("<a><b><c/></b></a>"), new ProjectLoadOptions { MaxDepth = 2 }));
        Assert.Throws<NotSupportedException>(() => ProjectDocument.Parse(Wrap(""), new ProjectLoadOptions { PackageSecurity = new OfficePackageSecurityOptions() }));
        Assert.Throws<OperationCanceledException>(() => ProjectDocument.Parse(Wrap(TaskXml()), cancellationToken: new CancellationToken(true)));
        using var document = ProjectDocument.Parse(Wrap(TaskXml()));
        Assert.Throws<InvalidDataException>(() => document.ToXml(new ProjectSaveOptions { MaxOutputBytes = 8 }));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void DependencyTypesAndNegativeOrPercentageLagRoundTrip(int type) {
        using var document = ProjectDocument.Create();
        var a = document.Tasks.Add("A");
        var b = document.Tasks.Add("B");
        var dependency = document.Dependencies.Add(a, b, (ProjectDependencyType)type);
        dependency.Lag = ProjectDuration.WorkingHours(-2);
        using var copy = ProjectDocument.Parse(document.ToXml());
        Assert.Equal(dependency.Type, copy.Dependencies[0].Type);
        Assert.Equal(dependency.Lag, copy.Dependencies[0].Lag);
        dependency.Lag = null;
        dependency.LagPercent = 50;
        using var percentCopy = ProjectDocument.Parse(document.ToXml());
        Assert.Equal(50m, percentCopy.Dependencies[0].LagPercent);
    }
}
