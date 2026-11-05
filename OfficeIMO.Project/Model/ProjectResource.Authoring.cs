namespace OfficeIMO.Project;

public sealed partial class ProjectResource {
    /// <summary>Registers a document-local resource alias for deferred assignments.</summary>
    public ProjectResource SetAlias(string alias) { Document.RegisterResourceAlias(alias, this); return this; }
    /// <summary>Sets the resource name.</summary>
    public ProjectResource SetName(string value) { Name = value; return this; }
    /// <summary>Sets a material resource kind and its label.</summary>
    public ProjectResource SetMaterial(string label) { Type = ProjectResourceType.Material; MaterialLabel = label; return this; }
    /// <summary>Sets the standard monetary rate.</summary>
    public ProjectResource SetStandardRate(decimal value) { StandardRate = value; return this; }
    /// <summary>Sets maximum allocation.</summary>
    public ProjectResource SetMaxUnits(ProjectUnits value) { MaxUnits = value; return this; }
    /// <summary>Assigns an explicit resource calendar.</summary>
    public ProjectResource SetCalendar(ProjectCalendar value) { Calendar = value; return this; }
}
