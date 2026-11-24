using OfficeIMO.Access;
using OfficeIMO.OneNote;
using OfficeIMO.Project;

namespace OfficeIMO.Scaling.Benchmarks;

internal sealed class ProjectBuildLane : ScalingLane {
    public override string Name => "Project.Build";
    public override string Unit => "task";

    public override object? Run() {
        using ProjectDocument project = ProjectDocument.Create();
        for (int i = 0; i < Units; i++) {
            project.Tasks.Add($"Task {i}");
        }

        var stream = new MemoryStream();
        project.Save(stream, new ProjectSaveOptions { Format = ProjectFileFormat.Xml });
        return stream.ToArray();
    }

    public override void Validate(object? result) {
        using ProjectDocument loaded = ProjectDocument.Load(new MemoryStream((byte[])result!));
        Require(loaded.Tasks.Count >= Units, "task count");
    }
}

internal sealed class ProjectReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Project.Read";
    public override string Unit => "task";

    protected override void Prepare(int units) {
        var lane = new ProjectBuildLane();
        lane.Setup(units);
        _bytes = (byte[])lane.Run()!;
    }

    public override object? Run() {
        using ProjectDocument loaded = ProjectDocument.Load(new MemoryStream(_bytes, writable: false));
        return loaded.Tasks.Count;
    }

    public override void Validate(object? result) => Require((int)result! >= Units, "task count");
}

internal sealed class AccessBuildLane : ScalingLane {
    public override string Name => "Access.Build";
    public override string Unit => "row";

    public override object? Run() {
        using AccessDocument database = AccessDocument.Create();
        var table = database.Tables.Add("Items");
        table.Columns.Add("Id", AccessDataType.Int32);
        table.Columns.Add("Name", AccessDataType.ShortText, 255);
        for (int i = 0; i < Units; i++) {
            table.AppendRow(new AccessRowValues { ["Id"] = i, ["Name"] = $"Name {i}" });
        }

        var stream = new MemoryStream();
        database.Save(stream);
        return stream.ToArray();
    }

    public override void Validate(object? result) {
        using AccessDocument loaded = AccessDocument.Load(new MemoryStream((byte[])result!));
        Require(loaded.Tables["Items"].RowCount == Units, "row count");
    }
}

internal sealed class AccessReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Access.Read";
    public override string Unit => "row";

    protected override void Prepare(int units) {
        var lane = new AccessBuildLane();
        lane.Setup(units);
        _bytes = (byte[])lane.Run()!;
    }

    public override object? Run() {
        using AccessDocument loaded = AccessDocument.Load(new MemoryStream(_bytes, writable: false));
        return loaded.Tables["Items"].RowCount;
    }

    public override void Validate(object? result) => Require(Convert.ToInt64(result) == Units, "row count");
}

internal sealed class OneNoteBuildLane : ScalingLane {
    public override string Name => "OneNote.Build";
    public override string Unit => "page";

    public override object? Run() {
        OneNoteSection section = OneNoteSection.Create();
        for (int i = 0; i < Units; i++) {
            var page = new OneNotePage { Title = $"Page {i}" };
            var paragraph = new OneNoteParagraph();
            paragraph.Runs.Add(new OneNoteTextRun { Text = $"Body text for page {i}." });
            page.DirectContent.Add(paragraph);
            section.Pages.Add(page);
        }

        return section.ToBytes();
    }

    public override void Validate(object? result) =>
        Require(OneNoteSection.Load(new MemoryStream((byte[])result!)).Pages.Count == Units, "page count");
}

internal sealed class OneNoteReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "OneNote.Read";
    public override string Unit => "page";

    protected override void Prepare(int units) {
        var lane = new OneNoteBuildLane();
        lane.Setup(units);
        _bytes = (byte[])lane.Run()!;
    }

    public override object? Run() => OneNoteSection.Load(new MemoryStream(_bytes, writable: false)).Pages.Count;
    public override void Validate(object? result) => Require((int)result! == Units, "page count");
}
