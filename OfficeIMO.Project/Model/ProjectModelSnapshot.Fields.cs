namespace OfficeIMO.Project;

/// <summary>Typed model inventory for fidelity comparisons, independent of file codecs and runtime reflection.</summary>
internal static partial class ProjectModelSnapshot {
    private static void Fields(ProjectSettings item, string path, ProjectModelValues values) {
        values[path + "/Calendar"] = item.Calendar?.Uid;
        values[path + "/StartDate"] = item.StartDate;
        values[path + "/FinishDate"] = item.FinishDate;
        values[path + "/ScheduleFromStart"] = item.ScheduleFromStart;
        values[path + "/MinutesPerDay"] = item.MinutesPerDay;
        values[path + "/MinutesPerWeek"] = item.MinutesPerWeek;
        values[path + "/DaysPerMonth"] = item.DaysPerMonth;
        values[path + "/DefaultStartTime"] = item.DefaultStartTime;
        values[path + "/DefaultFinishTime"] = item.DefaultFinishTime;
        values[path + "/CurrencyCode"] = item.CurrencyCode;
        values[path + "/CurrencySymbol"] = item.CurrencySymbol;
        values[path + "/CurrencyDigits"] = item.CurrencyDigits;
        values[path + "/StatusDate"] = item.StatusDate;
        values[path + "/DefaultTaskType"] = item.DefaultTaskType;
        values[path + "/NewTasksAreManual"] = item.NewTasksAreManual;
        values[path + "/ExternallyEdited"] = item.ExternallyEdited;
    }
    private static void Fields(ProjectTask item, string path, ProjectModelValues values) {
        values[path + "/IgnoreResourceCalendar"] = item.IgnoreResourceCalendar;
        values[path + "/FixedCostAccrual"] = item.FixedCostAccrual;
        values[path + "/Stop"] = item.Stop; values[path + "/Resume"] = item.Resume; values[path + "/EarnedValueMethod"] = item.EarnedValueMethod;
        values[path + "/Calendar"] = item.Calendar?.Uid;
        values[path + "/DisplayId"] = item.DisplayId;
        values[path + "/Wbs"] = item.Wbs;
        values[path + "/Duration"] = item.Duration;
        values[path + "/ActualDuration"] = item.ActualDuration;
        values[path + "/RemainingDuration"] = item.RemainingDuration;
        values[path + "/Work"] = item.Work;
        values[path + "/ActualWork"] = item.ActualWork;
        values[path + "/RemainingWork"] = item.RemainingWork;
        values[path + "/Start"] = item.Start;
        values[path + "/Finish"] = item.Finish;
        values[path + "/ActualStart"] = item.ActualStart;
        values[path + "/ActualFinish"] = item.ActualFinish;
        values[path + "/Deadline"] = item.Deadline;
        values[path + "/ConstraintDate"] = item.ConstraintDate;
        values[path + "/ConstraintType"] = item.ConstraintType;
        values[path + "/Type"] = item.Type;
        values[path + "/IsManual"] = item.IsManual;
        values[path + "/IsMilestone"] = item.IsMilestone;
        values[path + "/IsRecurring"] = item.IsRecurring;
        values[path + "/LevelingCanSplit"] = item.LevelingCanSplit;
        values[path + "/EffortDriven"] = item.EffortDriven;
        values[path + "/IsActive"] = item.IsActive;
        values[path + "/IsNull"] = item.IsNull;
        values[path + "/IsCritical"] = item.IsCritical;
        values[path + "/PercentComplete"] = item.PercentComplete;
        values[path + "/PercentWorkComplete"] = item.PercentWorkComplete;
        values[path + "/PhysicalPercentComplete"] = item.PhysicalPercentComplete;
        values[path + "/Priority"] = item.Priority;
        values[path + "/LevelingDelay"] = item.LevelingDelay;
        values[path + "/Cost"] = item.Cost;
        values[path + "/ActualCost"] = item.ActualCost;
        values[path + "/RemainingCost"] = item.RemainingCost;
        values[path + "/FixedCost"] = item.FixedCost;
        values[path + "/Notes"] = item.Notes;
        values[path + "/Contact"] = item.Contact;
        values[path + "/Hyperlink"] = item.Hyperlink;
        values[path + "/HyperlinkAddress"] = item.HyperlinkAddress;
        values[path + "/EarlyStart"] = item.EarlyStart;
        values[path + "/EarlyFinish"] = item.EarlyFinish;
        values[path + "/LateStart"] = item.LateStart;
        values[path + "/LateFinish"] = item.LateFinish;
        values[path + "/TotalSlackMinutes"] = item.TotalSlackMinutes;
        values[path + "/FreeSlackMinutes"] = item.FreeSlackMinutes;
    }
    private static void Fields(ProjectResource item, string path, ProjectModelValues values) {
        values[path + "/AccrueAt"] = item.AccrueAt;
        values[path + "/CanLevel"] = item.CanLevel;
        values[path + "/Calendar"] = item.Calendar?.Uid;
        values[path + "/DisplayId"] = item.DisplayId;
        values[path + "/Type"] = item.Type;
        values[path + "/Initials"] = item.Initials;
        values[path + "/Group"] = item.Group;
        values[path + "/EmailAddress"] = item.EmailAddress;
        values[path + "/MaterialLabel"] = item.MaterialLabel;
        values[path + "/MaxUnits"] = item.MaxUnits;
        values[path + "/StandardRate"] = item.StandardRate;
        values[path + "/OvertimeRate"] = item.OvertimeRate;
        values[path + "/CostPerUse"] = item.CostPerUse;
        values[path + "/Cost"] = item.Cost;
        values[path + "/ActualCost"] = item.ActualCost;
        values[path + "/Work"] = item.Work;
        values[path + "/ActualWork"] = item.ActualWork;
        values[path + "/RemainingWork"] = item.RemainingWork;
        values[path + "/Notes"] = item.Notes;
        values[path + "/IsNull"] = item.IsNull;
    }
    private static void Fields(ProjectAssignment item, string path, ProjectModelValues values) {
        values[path + "/Stop"] = item.Stop; values[path + "/Resume"] = item.Resume;
        values[path + "/CostRateTable"] = item.CostRateTable;
        values[path + "/DelayMinutes"] = item.DelayMinutes;
        values[path + "/HasFixedRateUnits"] = item.HasFixedRateUnits;
        values[path + "/MaterialRateScale"] = item.MaterialRateScale;
        values[path + "/WorkContour"] = item.WorkContour;
        values[path + "/Units"] = item.Units;
        values[path + "/Work"] = item.Work;
        values[path + "/ActualWork"] = item.ActualWork;
        values[path + "/RemainingWork"] = item.RemainingWork;
        values[path + "/OvertimeWork"] = item.OvertimeWork;
        values[path + "/ActualOvertimeWork"] = item.ActualOvertimeWork;
        values[path + "/Start"] = item.Start;
        values[path + "/Finish"] = item.Finish;
        values[path + "/ActualStart"] = item.ActualStart;
        values[path + "/ActualFinish"] = item.ActualFinish;
        values[path + "/Cost"] = item.Cost;
        values[path + "/ActualCost"] = item.ActualCost;
        values[path + "/RemainingCost"] = item.RemainingCost;
        values[path + "/PercentWorkComplete"] = item.PercentWorkComplete;
        values[path + "/Notes"] = item.Notes;
    }
    private static void Fields(ProjectCalendar item, string path, ProjectModelValues values) {
        values[path + "/IsBaseCalendar"] = item.IsBaseCalendar;
    }
    private static void Fields(ProjectBaseline item, string path, ProjectModelValues values) {
        values[path + "/Number"] = item.Number;
        values[path + "/Start"] = item.Start;
        values[path + "/Finish"] = item.Finish;
        values[path + "/Duration"] = item.Duration;
        values[path + "/Work"] = item.Work;
        values[path + "/Cost"] = item.Cost;
        values[path + "/FixedCost"] = item.FixedCost;
        values[path + "/Bcws"] = item.Bcws;
        values[path + "/Bcwp"] = item.Bcwp;
    }
    private static void Fields(ProjectCustomFieldValue item, string path, ProjectModelValues values) {
        values[path + "/FieldId"] = item.FieldId;
        values[path + "/Value"] = item.Value;
        values[path + "/ValueId"] = item.ValueId;
        values[path + "/ValueGuid"] = item.ValueGuid;
        values[path + "/DurationFormat"] = item.DurationFormat;
    }
    private static void Fields(ProjectCustomFieldDefinition item, string path, ProjectModelValues values) {
        values[path + "/LookupTableGuid"] = item.LookupTableGuid;
        values[path + "/FieldId"] = item.FieldId;
        values[path + "/FieldName"] = item.FieldName;
        values[path + "/Alias"] = item.Alias;
        values[path + "/FieldType"] = item.FieldType;
        values[path + "/Formula"] = item.Formula;
        values[path + "/SummaryCalculation"] = item.SummaryCalculation;
        values[path + "/RollupType"] = item.RollupType;
        values[path + "/Guid"] = item.Guid;
        values[path + "/RestrictValues"] = item.RestrictValues;
    }
    private static void Fields(ProjectLookupValue item, string path, ProjectModelValues values) {
        values[path + "/Id"] = item.Id;
        values[path + "/Value"] = item.Value;
        values[path + "/Description"] = item.Description;
        values[path + "/Guid"] = item.Guid;
    }
    private static void Fields(ProjectTimephasedValue item, string path, ProjectModelValues values) {
        values[path + "/Type"] = item.Type;
        values[path + "/Uid"] = item.Uid;
        values[path + "/Start"] = item.Start;
        values[path + "/Finish"] = item.Finish;
        values[path + "/Unit"] = item.Unit;
        values[path + "/Value"] = item.Value;
    }
    private static void Fields(ProjectDependency item, string path, ProjectModelValues values) {
        values[path + "/Type"] = item.Type;
        values[path + "/Lag"] = item.Lag;
        values[path + "/LagPercent"] = item.LagPercent;
        values[path + "/CrossProject"] = item.CrossProject;
        values[path + "/CrossProjectName"] = item.CrossProjectName;
    }
    private static void Fields(ProjectWeekDay item, string path, ProjectModelValues values) {
        values[path + "/Day"] = item.Day;
        values[path + "/IsWorking"] = item.IsWorking;
        values[path + "/FromDate"] = item.FromDate;
        values[path + "/ToDate"] = item.ToDate;
    }
    private static void Fields(ProjectCalendarException item, string path, ProjectModelValues values) {
        values[path + "/Name"] = item.Name;
        values[path + "/FromDate"] = item.FromDate;
        values[path + "/ToDate"] = item.ToDate;
        values[path + "/IsWorking"] = item.IsWorking;
    }
    private static void Fields(ProjectWorkingInterval item, string path, ProjectModelValues values) {
        values[path + "/From"] = item.From;
        values[path + "/To"] = item.To;
    }
    private static void Fields(ProjectWorkWeek item, string path, ProjectModelValues values) {
        values[path + "/Name"] = item.Name;
        values[path + "/FromDate"] = item.FromDate;
        values[path + "/ToDate"] = item.ToDate;
    }
}
