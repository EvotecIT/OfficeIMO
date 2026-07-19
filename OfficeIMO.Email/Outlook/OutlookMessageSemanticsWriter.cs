namespace OfficeIMO.Email;

internal static class OutlookMessageSemanticsWriter {
    internal static void Apply(MsgPropertyBuilder properties, EmailDocument document, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        EmailMessageMetadata metadata = document.MessageMetadata;
        bool meetingRelated = document.OutlookItemKind == OutlookItemKind.Appointment ||
            (document.MessageClass?.StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase) ?? false);
        bool task = document.OutlookItemKind == OutlookItemKind.Task;
        if (!meetingRelated && !task) {
            WriteFollowUp(properties, metadata.FollowUp);
            WriteReminder(properties, metadata.Reminder);
        }
        WriteVoting(properties, metadata.Voting, codePage, diagnostics, location);
    }

    internal static void WriteReminder(MsgPropertyBuilder properties, OutlookReminder reminder) {
        if (!reminder.IsSet.HasValue && !reminder.DeltaMinutes.HasValue && !reminder.Time.HasValue &&
            !reminder.SignalTime.HasValue && !reminder.Override.HasValue && !reminder.PlaySound.HasValue &&
            reminder.SoundFile == null) return;
        properties.Set(MapiKnownProperties.PidLid.ReminderSet, reminder.IsSet);
        properties.Set(MapiKnownProperties.PidLid.ReminderDelta, reminder.DeltaMinutes);
        properties.Set(MapiKnownProperties.PidLid.ReminderTime, reminder.Time);
        properties.Set(MapiKnownProperties.PidLid.ReminderSignalTime, reminder.SignalTime);
        properties.Set(MapiKnownProperties.PidLid.ReminderOverride, reminder.Override);
        properties.Set(MapiKnownProperties.PidLid.ReminderPlaySound, reminder.PlaySound);
        properties.Set(MapiKnownProperties.PidLid.ReminderFileParameter, reminder.SoundFile);
    }

    private static void WriteFollowUp(MsgPropertyBuilder properties, OutlookFollowUp followUp) {
        if (!followUp.Status.HasValue) return;
        int? rawStatus = (int)followUp.Status.Value;
        if (rawStatus == 0) rawStatus = null;
        properties.Set(MapiKnownProperties.PidTag.FlagStatus, rawStatus);
        properties.Set(MapiKnownProperties.PidLid.FlagRequest, rawStatus.HasValue ? followUp.Request : null);
        properties.Set(MapiKnownProperties.PidLid.ToDoTitle, rawStatus.HasValue ? followUp.Title : null);
        properties.Set(MapiKnownProperties.PidLid.TaskStartDate, rawStatus.HasValue ? followUp.Start : null);
        properties.Set(MapiKnownProperties.PidLid.TaskDueDate, rawStatus.HasValue ? followUp.Due : null);
        properties.Set(MapiKnownProperties.PidTag.ToDoItemFlags, rawStatus.HasValue ? followUp.ToDoItemFlags : null);
        properties.Set(MapiKnownProperties.PidLid.FlagString, rawStatus.HasValue ? followUp.FlagString : null);
        properties.Set(MapiKnownProperties.PidLid.ValidFlagStringProof,
            rawStatus.HasValue ? followUp.ValidRequestProof : null);

        bool complete = rawStatus == (int)OutlookFollowUpStatus.Complete;
        bool flagged = rawStatus == (int)OutlookFollowUpStatus.Flagged;
        properties.Set(MapiKnownProperties.PidTag.FlagCompleteTime, complete ? followUp.CompletedAt : null);
        properties.Set(MapiKnownProperties.PidTag.FollowupIcon,
            flagged && followUp.Icon.HasValue && followUp.Icon.Value != OutlookFollowUpIcon.None
                ? (int)followUp.Icon.Value
                : (int?)null);
        if (followUp.Status.HasValue) {
            properties.Set(MapiKnownProperties.PidLid.TaskDateCompleted, complete ? followUp.CompletedAt : null);
            properties.Set(MapiKnownProperties.PidLid.TaskComplete, complete);
            properties.Set(MapiKnownProperties.PidLid.TaskStatus, complete ? 2 : 0);
            properties.Set(MapiKnownProperties.PidLid.PercentComplete, complete ? 1d : 0d);
        }
    }

    private static void WriteVoting(MsgPropertyBuilder properties, OutlookVoting voting, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        properties.Set(MapiKnownProperties.PidLid.VerbResponse, voting.Response);
        if (voting.Options.Count == 0) {
            if (voting.OptionsClearRequested) {
                properties.Set(MapiKnownProperties.PidLid.VerbStream, null);
            } else if (voting.RawVerbStream != null) {
                properties.Set(MapiKnownProperties.PidLid.VerbStream, voting.RawVerbStream);
            }
            return;
        }
        if (OutlookVotingCodec.TryEncode(voting.Options.ToArray(), codePage,
            out byte[]? stream, out string? error)) {
            properties.Set(MapiKnownProperties.PidLid.VerbStream, stream);
            return;
        }
        diagnostics.Add(new EmailDiagnostic(
            "EMAIL_MSG_VOTING_OPTIONS_INVALID",
            string.Concat("The Outlook voting options could not be serialized: ", error),
            EmailDiagnosticSeverity.Error,
            string.Concat(location, "/voting")));
    }
}
