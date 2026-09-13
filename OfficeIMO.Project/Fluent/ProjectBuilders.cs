namespace OfficeIMO.Project.Fluent;

/// <summary>Metadata configuration over the document.</summary>
public sealed class ProjectInfoBuilder {
    private readonly ProjectDocument _document;
    internal ProjectInfoBuilder(ProjectDocument document) { _document = document; }
    /// <summary>Sets the project name.</summary>
    public ProjectInfoBuilder Name(string value) { _document.Name = value; return this; }
    /// <summary>Sets the document title.</summary>
    public ProjectInfoBuilder Title(string value) { _document.Title = value; return this; }
    /// <summary>Sets the author.</summary>
    public ProjectInfoBuilder Author(string value) { _document.Author = value; return this; }
    /// <summary>Sets organization metadata.</summary>
    public ProjectInfoBuilder Company(string value) { _document.Company = value; return this; }
}

/// <summary>Fluent task creation inside an existing summary.</summary>
public sealed class ProjectFluentTasks {
    private readonly ProjectFluentDocument _builder;
    private readonly ProjectTaskCollection _tasks;
    internal ProjectFluentTasks(ProjectFluentDocument builder, ProjectTaskCollection tasks) { _builder = builder; _tasks = tasks; }
    /// <summary>Adds a child task and registers its alias.</summary>
    public ProjectFluentTasks Task(string alias, string name, Action<ProjectTaskBuilder>? configure = null) {
        _builder.AddTask(_tasks, alias, name, false, configure); return this;
    }
    /// <summary>Adds a nested summary.</summary>
    public ProjectFluentTasks Summary(string alias, string name, Action<ProjectFluentTasks> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        var task = _builder.AddTask(_tasks, alias, name, true, null);
        configure(new ProjectFluentTasks(_builder, task.Children)); return this;
    }
}

/// <summary>Thin task configuration; dates, work, and costs are inputs rather than calculated results.</summary>
public sealed class ProjectTaskBuilder {
    private readonly ProjectFluentDocument _builder;
    private readonly ProjectTask _task;
    internal ProjectTaskBuilder(ProjectFluentDocument builder, ProjectTask task) { _builder = builder; _task = task; }
    /// <summary>Sets task name.</summary>
    public ProjectTaskBuilder Name(string value) { _task.Name = value; return this; }
    /// <summary>Sets typed duration.</summary>
    public ProjectTaskBuilder Duration(ProjectDuration value) { _task.Duration = value; return this; }
    /// <summary>Sets stored task work.</summary>
    public ProjectTaskBuilder Work(ProjectWork value) { _task.Work = value; return this; }
    /// <summary>Sets stored start/finish dates.</summary>
    public ProjectTaskBuilder Dates(DateTime start, DateTime finish) { _task.Start = start; _task.Finish = finish; return this; }
    /// <summary>Sets manual scheduling mode.</summary>
    public ProjectTaskBuilder Manual(bool value = true) { _task.IsManual = value; return this; }
    /// <summary>Marks a zero-duration milestone.</summary>
    public ProjectTaskBuilder Milestone() { _task.IsMilestone = true; _task.Duration = ProjectDuration.WorkingMinutes(0); return this; }
    /// <summary>Sets task notes.</summary>
    public ProjectTaskBuilder Notes(string value) { _task.Notes = value; return this; }
    /// <summary>Sets an explicit task calendar.</summary>
    public ProjectTaskBuilder Calendar(ProjectCalendar value) { _task.Calendar = value; return this; }
    /// <summary>Stores a dependency; a forward alias is resolved when End is called.</summary>
    public ProjectTaskBuilder After(string alias, ProjectDependencyType type = ProjectDependencyType.FinishToStart, ProjectDuration? lag = null) {
        _builder.Link(_task, alias, type, lag); return this;
    }
    /// <summary>Stores an assignment to a resource alias, resolved when End is called.</summary>
    public ProjectTaskBuilder Assign(string alias, ProjectUnits? units = null) { _builder.Assign(_task, alias, units); return this; }
    /// <summary>Stores a custom field's text by its Microsoft Project field ID.</summary>
    public ProjectTaskBuilder Text(string fieldId, string value) {
        var field = _task.CustomFields.FirstOrDefault(f => f.FieldId == fieldId) ?? _task.CustomFields.Add();
        field.FieldId = fieldId; field.Value = value; return this;
    }
}

/// <summary>Thin resource configuration over a typed resource.</summary>
public sealed class ProjectResourceBuilder {
    private readonly ProjectResource _resource;
    internal ProjectResourceBuilder(ProjectResource resource) { _resource = resource; }
    /// <summary>Sets resource name.</summary>
    public ProjectResourceBuilder Name(string value) { _resource.Name = value; return this; }
    /// <summary>Sets the work resource kind.</summary>
    public ProjectResourceBuilder Work() { _resource.Type = ProjectResourceType.Work; return this; }
    /// <summary>Sets the material resource kind and label.</summary>
    public ProjectResourceBuilder Material(string label) { _resource.Type = ProjectResourceType.Material; _resource.MaterialLabel = label; return this; }
    /// <summary>Sets the cost resource kind.</summary>
    public ProjectResourceBuilder Cost() { _resource.Type = ProjectResourceType.Cost; return this; }
    /// <summary>Sets the standard monetary rate.</summary>
    public ProjectResourceBuilder StandardRate(decimal value) { _resource.StandardRate = value; return this; }
    /// <summary>Sets maximum allocation.</summary>
    public ProjectResourceBuilder MaxUnits(ProjectUnits value) { _resource.MaxUnits = value; return this; }
    /// <summary>Sets an explicit resource calendar.</summary>
    public ProjectResourceBuilder Calendar(ProjectCalendar value) { _resource.Calendar = value; return this; }
}
