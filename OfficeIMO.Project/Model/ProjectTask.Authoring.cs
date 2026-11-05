namespace OfficeIMO.Project;

public sealed partial class ProjectTask {
    /// <summary>Registers an explicit document-local alias for deferred relationships, without persisting it.</summary>
    public ProjectTask SetAlias(string alias) { Document.RegisterTaskAlias(alias, this); return this; }
    /// <summary>Sets the task name.</summary>
    public ProjectTask SetName(string value) { Name = value; return this; }
    /// <summary>Sets stored duration without calculating dates.</summary>
    public ProjectTask SetDuration(ProjectDuration value) { Duration = value; return this; }
    /// <summary>Sets stored work without calculating dates.</summary>
    public ProjectTask SetWork(ProjectWork value) { Work = value; return this; }
    /// <summary>Sets stored start and finish dates without scheduling.</summary>
    public ProjectTask SetDates(DateTime start, DateTime finish) { Start = start; Finish = finish; return this; }
    /// <summary>Sets manual scheduling mode.</summary>
    public ProjectTask SetManual(bool value = true) { IsManual = value; return this; }
    /// <summary>Marks a zero-duration milestone.</summary>
    public ProjectTask SetMilestone() { IsMilestone = true; Duration = ProjectDuration.WorkingMinutes(0); return this; }
    /// <summary>Sets task notes.</summary>
    public ProjectTask SetNotes(string value) { Notes = value; return this; }
    /// <summary>Assigns an explicit task calendar.</summary>
    public ProjectTask SetCalendar(ProjectCalendar value) { Calendar = value; return this; }
    /// <summary>Queues a dependency on a forward task alias; ProjectDocument.End validates and applies the batch.</summary>
    public ProjectTask After(string alias, ProjectDependencyType type = ProjectDependencyType.FinishToStart, ProjectDuration? lag = null) {
        EnsureAttached(); Document.QueueDependency(this, alias, type, lag); return this;
    }
    /// <summary>Queues a resource assignment; ProjectDocument.End validates and applies the batch.</summary>
    public ProjectTask Assign(string alias, ProjectUnits? units = null) { EnsureAttached(); Document.QueueAssignment(this, alias, units); return this; }
    /// <summary>Sets a custom field's text by its Microsoft Project field ID.</summary>
    public ProjectTask SetText(string fieldId, string value) {
        var field = CustomFields.FirstOrDefault(f => f.FieldId == fieldId) ?? CustomFields.Add();
        field.FieldId = fieldId; field.Value = value; return this;
    }
}
