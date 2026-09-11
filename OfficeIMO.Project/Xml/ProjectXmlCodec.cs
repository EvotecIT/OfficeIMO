using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    internal const string NamespaceName = "http://schemas.microsoft.com/project";
    internal static readonly string[] RootOrder = "SaveVersion BuildNumber UID Name GUID Title Subject Category Company Manager Author CreationDate Revision LastSaved ScheduleFromStart StartDate FinishDate FYStartDate CriticalSlackLimit CurrencyDigits CurrencySymbol CurrencyCode CurrencySymbolPosition CalendarUID DefaultStartTime DefaultFinishTime MinutesPerDay MinutesPerWeek DaysPerMonth DefaultTaskType DefaultFixedCostAccrual DefaultStandardRate DefaultOvertimeRate DurationFormat WorkFormat StatusDate CurrentDate NewTasksAreManual ProjectExternallyEdited OutlineCodes ExtendedAttributes Calendars Tasks Resources Assignments".Split(' ');
    internal static readonly string[] TaskOrder = "UID GUID ID Name Type IsNull CreateDate Contact WBS WBSLevel OutlineNumber OutlineLevel Priority Start Finish Duration DurationFormat Work Stop Resume ResumeValid EffortDriven Recurring OverAllocated Estimated Milestone Summary Critical IsSubproject IsSubprojectReadOnly SubprojectName ExternalTask ExternalTaskProject EarlyStart EarlyFinish LateStart LateFinish StartVariance FinishVariance WorkVariance FreeSlack TotalSlack FixedCost FixedCostAccrual PercentComplete PercentWorkComplete Cost OvertimeCost OvertimeWork ActualStart ActualFinish ActualDuration ActualCost ActualOvertimeCost ActualWork ActualOvertimeWork RegularWork RemainingDuration RemainingCost RemainingWork RemainingOvertimeCost RemainingOvertimeWork ACWP CV ConstraintType CalendarUID ConstraintDate Deadline LevelAssignments LevelingCanSplit LevelingDelay LevelingDelayFormat PreLeveledStart PreLeveledFinish Hyperlink HyperlinkAddress HyperlinkSubAddress IgnoreResourceCalendar Notes HideBar Rollup BCWS BCWP PhysicalPercentComplete EarnedValueMethod PredecessorLink ExtendedAttribute OutlineCode Baseline Active Manual TimephasedData".Split(' ');
    internal static readonly string[] ResourceOrder = "UID GUID ID Name Type IsNull Initials Phonetics NTAccount MaterialLabel Code Group WorkGroup EmailAddress Hyperlink HyperlinkAddress MaxUnits PeakUnits OverAllocated AvailableFrom AvailableTo Start Finish CanLevel AccrueAt Work RegularWork OvertimeWork ActualWork RemainingWork ActualOvertimeWork RemainingOvertimeWork PercentWorkComplete StandardRate StandardRateFormat Cost OvertimeRate OvertimeRateFormat OvertimeCost CostPerUse ActualCost ActualOvertimeCost RemainingCost RemainingOvertimeCost WorkVariance CostVariance SV CV ACWP CalendarUID Notes BCWS BCWP ExtendedAttribute OutlineCode Baseline IsCostResource AvailabilityPeriods Rates TimephasedData".Split(' ');
    internal static readonly string[] AssignmentOrder = "UID GUID TaskUID ResourceUID PercentWorkComplete ActualCost ActualFinish ActualOvertimeCost ActualOvertimeWork ActualStart ActualWork ACWP Confirmed Cost CostRateTable RateScale CostVariance CV Delay Finish FinishVariance Hyperlink HyperlinkAddress HyperlinkSubAddress WorkVariance HasFixedRateUnits FixedMaterial LevelingDelay LevelingDelayFormat LinkedFields Milestone Notes Overallocated OvertimeCost OvertimeWork PeakUnits RegularWork RemainingCost RemainingOvertimeCost RemainingOvertimeWork RemainingWork ResponsePending Start Stop Resume StartVariance Summary SV Units UpdateNeeded VAC Work WorkContour BCWS BCWP BookingType ExtendedAttribute Baseline TimephasedData".Split(' ');
    internal static readonly string[] CalendarOrder = "UID GUID Name IsBaseCalendar IsBaselineCalendar BaseCalendarUID WeekDays Exceptions WorkWeeks".Split(' ');
    internal static readonly string[] BaselineOrder = "TimephasedData Number Interim Start Finish Duration DurationFormat EstimatedDuration Work Cost BCWS BCWP FixedCost".Split(' ');

    internal static ProjectDocument Read(byte[] bytes, ProjectLoadOptions options, CancellationToken cancellationToken) {
        options.ValidateLimits();
        if (bytes.LongLength > options.MaxInputBytes) throw new InvalidDataException("Project input exceeds MaxInputBytes.");
        if (bytes.Length >= 8 && bytes[0] == 0xd0 && bytes[1] == 0xcf && bytes[2] == 0x11 && bytes[3] == 0xe0)
            throw new NotSupportedException("This is a compound binary file. Native MPP loading is not supported by the XML document codec.");
        using var stream = new MemoryStream(bytes, false);
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = options.MaxCharacters };
        using var rawReader = XmlReader.Create(stream, settings);
        using var reader = new OfficeXmlLimitingReader(rawReader, "MSPDI", options.MaxDepth, options.MaxElements, options.MaxAttributes, cancellationToken);
        var xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        var root = xml.Root;
        if (root == null || root.Name.LocalName != "Project" || (root.Name.NamespaceName != NamespaceName && root.Name.NamespaceName != NamespaceName + "/2007"))
            throw new InvalidDataException("Expected a Project root in a supported Microsoft Project XML namespace.");
        var document = ProjectDocument.CreateForRead();
        document.Source = new ProjectXmlSource(xml, bytes);
        try {
            ReadDocument(document, root, options, cancellationToken);
            CaptureSnapshots(document, cancellationToken);
            InspectPreservedContent(document, options, cancellationToken);
            document.FinishRead(options);
            return document;
        } catch { document.DisposeFailedRead(); throw; }
    }

    internal static byte[] Write(ProjectDocument document, ProjectSaveOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (RetainedBytes(document, options) is byte[] original) {
            if (original.LongLength > options.MaxOutputBytes) throw new InvalidDataException("Project output exceeds MaxOutputBytes.");
            return original;
        }
        var xml = WriteDocument(document, cancellationToken);
        using var buffer = new OfficeBoundedMemoryStream(options.MaxOutputBytes);
        using (var writer = XmlWriter.Create(buffer, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = options.Indent, CloseOutput = false, NewLineHandling = NewLineHandling.Entitize })) {
            xml.Save(writer);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return buffer.ToArray();
    }

    internal static byte[]? RetainedBytes(ProjectDocument document, ProjectSaveOptions options) =>
        !document.IsModified && options.PreserveUnchangedBytes ? document.LastSavedBytes ?? document.Source?.OriginalBytes : null;

    private static XElement NewNode(ProjectDocument document, object model, string name) =>
        document.Source?.CloneOrCreate(model, name) ?? new XElement(XName.Get(name, document.XmlNamespace));

    private static int RequiredUid(XElement element) {
        var uid = element.Element(element.Name.Namespace + "UID");
        if (uid == null || !int.TryParse(uid.Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int result) || result < 0)
            throw new InvalidDataException("A nonnegative UID is required at " + ProjectXmlValue.Location(element));
        return result;
    }
    private static IEnumerable<XElement> Children(XElement parent, string container, string name) {
        var containers = parent.Elements(parent.Name.Namespace + container).ToArray();
        if (containers.Length > 1) throw new InvalidDataException("Duplicate " + container + " container.");
        return containers.FirstOrDefault()?.Elements(parent.Name.Namespace + name) ?? Enumerable.Empty<XElement>();
    }
    private static void Attach(ProjectDocument document, object model, XElement element) {
        string fields = model switch {
            ProjectDocument => "SaveVersion CalendarUID",
            ProjectTask => "UID Summary OutlineLevel CalendarUID DurationFormat",
            ProjectResource => "UID CalendarUID IsCostResource",
            ProjectAssignment => "UID TaskUID ResourceUID",
            ProjectCalendar => "UID Name GUID IsBaseCalendar BaseCalendarUID",
            ProjectWeekDay => "DayType DayWorking",
            ProjectCalendarException => "Name DayWorking",
            ProjectWorkingInterval => "FromTime ToTime",
            ProjectDependency => "PredecessorUID Type CrossProject CrossProjectName LinkLag LagFormat",
            ProjectBaseline => "DurationFormat",
            _ => ""
        };
        foreach (string name in fields.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) CheckScalar(element, name);
        if (model is ProjectWeekDay || model is ProjectCalendarException) {
            var periods = element.Elements(element.Name.Namespace + "TimePeriod").ToArray();
            if (periods.Length > 1) throw new InvalidDataException("Duplicate TimePeriod.");
            if (periods.Length == 1) { CheckScalar(periods[0], "FromDate"); CheckScalar(periods[0], "ToDate"); }
        }
        document.Source!.Attach(model, element);
    }
    private static void CheckScalar(XElement parent, string name) {
        var values = parent.Elements(parent.Name.Namespace + name).ToArray();
        if (values.Length > 1 || (values.Length == 1 && values[0].HasElements))
            throw new InvalidDataException("Expected a single scalar " + name + " at " + ProjectXmlValue.Location(parent));
    }
    private static void ReadFields<T>(T model, XElement element, ProjectDocument document, IEnumerable<ProjectXmlField<T>> fields) where T : class => ProjectXmlFields.Read(model, element, document, fields);

    private static void ReplaceChildren(XElement parent, string childName, IEnumerable<XElement> output, string[]? order = null) {
        var existing = parent.Elements(parent.Name.Namespace + childName).ToArray();
        var replacement = output.ToArray();
        if (existing.Length != 0) {
            // Removing old nodes after inserting the replacements repeatedly searches
            // a singly linked sibling list. Rebuild once to keep large collections linear
            // while retaining comments, whitespace, and foreign extension nodes.
            var selected = new HashSet<XElement>(existing);
            var nodes = new List<XNode>();
            int replacementIndex = 0, existingIndex = 0;
            foreach (var node in parent.Nodes()) {
                if (node is XElement element && selected.Contains(element)) {
                    if (replacementIndex < replacement.Length) nodes.Add(replacement[replacementIndex++]);
                    if (++existingIndex == existing.Length) while (replacementIndex < replacement.Length) nodes.Add(replacement[replacementIndex++]);
                } else nodes.Add(node);
            }
            parent.ReplaceNodes(nodes);
        } else foreach (var child in replacement) ProjectXmlFields.Insert(parent, child, order ?? Array.Empty<string>());
    }
    private static void ReplaceContainer(XElement parent, string containerName, string childName, IEnumerable<XElement> output, string[] order) {
        var replacement = output.ToArray();
        var container = parent.Element(parent.Name.Namespace + containerName);
        if (container == null) {
            if (replacement.Length == 0) return;
            container = new XElement(parent.Name.Namespace + containerName);
            ProjectXmlFields.Insert(parent, container, order);
        }
        ReplaceChildren(container, childName, replacement);
    }
}
