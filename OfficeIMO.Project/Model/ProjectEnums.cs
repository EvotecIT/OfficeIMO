namespace OfficeIMO.Project;

/// <summary>A dependency's relationship between predecessor and successor dates.</summary>
public enum ProjectDependencyType {
    /// <summary>Finish to finish.</summary>
    FinishToFinish = 0,
    /// <summary>Finish to start.</summary>
    FinishToStart = 1,
    /// <summary>Start to finish.</summary>
    StartToFinish = 2,
    /// <summary>Start to start.</summary>
    StartToStart = 3
}
/// <summary>The task quantity held fixed during later explicit scheduling.</summary>
public enum ProjectTaskType {
    /// <summary>Allocation units stay fixed.</summary>
    FixedUnits = 0,
    /// <summary>Duration stays fixed.</summary>
    FixedDuration = 1,
    /// <summary>Work stays fixed.</summary>
    FixedWork = 2
}
/// <summary>A scheduling constraint; storing a constraint does not calculate it.</summary>
public enum ProjectConstraintType {
    /// <summary>Schedule as soon as possible.</summary>
    AsSoonAsPossible = 0,
    /// <summary>Schedule as late as possible.</summary>
    AsLateAsPossible = 1,
    /// <summary>Must start on the constraint date.</summary>
    MustStartOn = 2,
    /// <summary>Must finish on the constraint date.</summary>
    MustFinishOn = 3,
    /// <summary>Start no earlier than the constraint date.</summary>
    StartNoEarlierThan = 4,
    /// <summary>Start no later than the constraint date.</summary>
    StartNoLaterThan = 5,
    /// <summary>Finish no earlier than the constraint date.</summary>
    FinishNoEarlierThan = 6,
    /// <summary>Finish no later than the constraint date.</summary>
    FinishNoLaterThan = 7
}
/// <summary>The resource accounting kind.</summary>
public enum ProjectResourceType {
    /// <summary>A material resource.</summary>
    Material = 0,
    /// <summary>A work resource.</summary>
    Work = 1,
    /// <summary>A cost resource.</summary>
    Cost = 2
}
/// <summary>How dependent objects are handled when a task or resource is removed.</summary>
public enum ProjectRemovalMode {
    /// <summary>Reject removal when related entities would be orphaned.</summary>
    RejectIfReferenced,
    /// <summary>Remove the selected task subtree and its internal/external assignments and dependencies.</summary>
    Cascade
}
