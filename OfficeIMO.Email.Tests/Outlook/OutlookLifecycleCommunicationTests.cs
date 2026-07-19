using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookLifecycleCommunicationTests {
    [Fact]
    public void MeetingCounterProposalRoundTripsTypedEnvelopeThroughMsg() {
        DateTimeOffset start = new DateTimeOffset(2026, 11, 3, 9, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "New time",
            Appointment = new OutlookAppointment { Start = start, End = start.AddMinutes(30) },
            MeetingCommunication = new OutlookMeetingCommunication {
                Kind = OutlookMeetingCommunicationKind.ResponseTentative,
                IsCounterProposal = true,
                ProposedStart = start.AddHours(2),
                ProposedEnd = start.AddHours(3),
                ProposedDurationMinutes = 60,
                AttendeeCriticalChange = start.AddMinutes(-5),
                IsSilent = false,
                ReplyAt = start.AddMinutes(-5),
                ReplyName = "Attendee"
            }
        };

        EmailDocument result = RoundTrip(source, EmailFileFormat.OutlookMsg);

        Assert.Equal("IPM.Schedule.Meeting.Resp.Tent", result.MessageClass);
        Assert.Equal(OutlookItemKind.Appointment, result.OutlookItemKind);
        Assert.NotNull(result.Appointment);
        OutlookMeetingCommunication communication = Assert.IsType<OutlookMeetingCommunication>(result.MeetingCommunication);
        Assert.Equal(OutlookMeetingCommunicationKind.ResponseTentative, communication.Kind);
        Assert.True(communication.IsCounterProposal);
        Assert.Equal(start.AddHours(2), communication.ProposedStart);
        Assert.Equal(start.AddHours(3), communication.ProposedEnd);
        Assert.Equal(60, communication.ProposedDurationMinutes);
        Assert.False(communication.IsSilent);
        Assert.Equal("Attendee", communication.ReplyName);
    }

    [Theory]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void MeetingRequestRoundTripsProtocolTypeAndCriticalChanges(EmailFileFormat format) {
        DateTimeOffset change = new DateTimeOffset(2026, 11, 4, 8, 30, 0, TimeSpan.Zero);
        byte[] globalObjectId = Enumerable.Range(1, 40).Select(value => (byte)value).ToArray();
        byte[] cleanGlobalObjectId = Enumerable.Range(41, 40).Select(value => (byte)value).ToArray();
        var source = new EmailDocument {
            Format = format,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Updated meeting",
            Appointment = new OutlookAppointment {
                Start = change.AddDays(1),
                End = change.AddDays(1).AddHours(1),
                GlobalObjectId = globalObjectId,
                CleanGlobalObjectId = cleanGlobalObjectId
            },
            MeetingCommunication = new OutlookMeetingCommunication {
                Kind = OutlookMeetingCommunicationKind.RequestOrUpdate,
                RequestType = OutlookMeetingRequestType.FullUpdate | OutlookMeetingRequestType.OutOfDate,
                IntendedBusyStatus = 2,
                OwnerCriticalChange = change,
                AttendeeCriticalChange = change.AddMinutes(1)
            }
        };

        EmailDocument result = RoundTrip(source, format);

        Assert.Equal("IPM.Schedule.Meeting.Request", result.MessageClass);
        OutlookMeetingCommunication communication = Assert.IsType<OutlookMeetingCommunication>(result.MeetingCommunication);
        Assert.Equal(OutlookMeetingCommunicationKind.RequestOrUpdate, communication.Kind);
        Assert.Equal(OutlookMeetingRequestType.FullUpdate | OutlookMeetingRequestType.OutOfDate,
            communication.RequestType);
        Assert.Equal(2, communication.IntendedBusyStatus);
        Assert.Equal(change, communication.OwnerCriticalChange);
        Assert.Equal(change.AddMinutes(1), communication.AttendeeCriticalChange);
        Assert.Equal(globalObjectId, result.Appointment!.GlobalObjectId);
        Assert.Equal(cleanGlobalObjectId, result.Appointment.CleanGlobalObjectId);
    }

    [Theory]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void TaskRequestWritesCanonicalFirstHiddenEmbeddedTask(EmailFileFormat format) {
        Guid globalId = new Guid("759228F7-84D3-49D6-865E-F8655ADFC1DC");
        var embedded = new EmailDocument {
            OutlookItemKind = OutlookItemKind.Task,
            Subject = "Assigned work",
            Task = new OutlookTask {
                Status = 1,
                PercentComplete = 0.25,
                GlobalId = globalId,
                CommunicationMode = OutlookTaskCommunicationMode.EmbeddedRequest
            }
        };
        var source = new EmailDocument {
            Format = format,
            OutlookItemKind = OutlookItemKind.Task,
            Subject = "Task request",
            TaskCommunication = OutlookTaskCommunication.Create(OutlookTaskCommunicationKind.Request, embedded)
        };
        source.Attachments.Add(new EmailAttachment { FileName = "ordinary.txt", Content = new byte[] { 1, 2, 3 } });

        EmailDocument result = RoundTrip(source, format);

        Assert.Equal("IPM.TaskRequest", result.MessageClass);
        OutlookTaskCommunication communication = Assert.IsType<OutlookTaskCommunication>(result.TaskCommunication);
        Assert.Equal(OutlookTaskCommunicationKind.Request, communication.Kind);
        Assert.NotNull(communication.EmbeddedTask?.Task);
        Assert.Equal(globalId, communication.EmbeddedTask!.Task!.GlobalId);
        EmailAttachment payload = Assert.IsType<EmailAttachment>(communication.PayloadAttachment);
        Assert.Same(result.Attachments[0], payload);
        Assert.Equal(5, payload.MapiAttachMethod);
        Assert.True(payload.IsHidden);
        Assert.Equal(-1, payload.RenderingPosition);
        Assert.True(communication.Validate().IsValid);
        Assert.Equal("ordinary.txt", result.Attachments[1].FileName);
    }

    [Fact]
    public void TaskLifecyclePropertiesRoundTripWithTypedAliases() {
        DateTimeOffset update = new DateTimeOffset(2026, 11, 5, 16, 0, 0, TimeSpan.Zero);
        Guid globalId = new Guid("0AD10E9F-2FF8-40B1-900A-E7FC939A8F26");
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Subject = "Lifecycle",
            Task = new OutlookTask {
                IsAccepted = true,
                HistoryKind = OutlookTaskHistoryKind.Accepted,
                LastUpdate = update,
                LastUser = "Assignee",
                LastDelegate = "Assigner",
                CommunicationMode = OutlookTaskCommunicationMode.Accepted,
                GlobalId = globalId
            }
        };

        OutlookTask task = RoundTrip(source, EmailFileFormat.OutlookMsg).Task!;

        Assert.True(task.IsAccepted);
        Assert.Equal(OutlookTaskHistoryKind.Accepted, task.HistoryKind);
        Assert.Equal(update, task.LastUpdate);
        Assert.Equal("Assignee", task.LastUser);
        Assert.Equal("Assigner", task.LastDelegate);
        Assert.Equal(OutlookTaskCommunicationMode.Accepted, task.CommunicationMode);
        Assert.Equal(globalId, task.GlobalId);
    }

    [Fact]
    public void TaskCommunicationParticipatesInSemanticComparisonAndReportsEmlLoss() {
        var embedded = new EmailDocument {
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.Task",
            Subject = "Semantic task",
            Task = new OutlookTask { GlobalId = new Guid("090515F8-E8ED-4C8B-AEB5-2562C5889982") }
        };
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.TaskRequest",
            Subject = "Semantic request",
            TaskCommunication = OutlookTaskCommunication.Create(OutlookTaskCommunicationKind.Request, embedded)
        };

        var changedEmbedded = new EmailDocument {
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.Task",
            Subject = "Changed semantic task",
            Task = new OutlookTask { GlobalId = new Guid("61D84FE1-B66C-4B2D-BC31-C33DE6193F3D") }
        };
        var changed = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.TaskRequest",
            Subject = source.Subject,
            TaskCommunication = OutlookTaskCommunication.Create(OutlookTaskCommunicationKind.Request, changedEmbedded)
        };
        EmailSemanticComparisonReport comparison = EmailSemanticComparer.Compare(source, changed);
        EmailConversionReport conversion = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(comparison.IsMatch);
        Assert.Contains(comparison.Differences,
            difference => difference.Path.Contains("attachments/00000000/embedded", StringComparison.Ordinal));
        Assert.Equal(1, comparison.Source.AttachmentCount);
        Assert.True(conversion.HasPotentialDataLoss);
        Assert.Contains(conversion.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_TASK_COMMUNICATION_UNSUPPORTED");
    }

    [Theory]
    [InlineData("IPM.TaskRequest.Accept", OutlookTaskCommunicationKind.Accept)]
    [InlineData("IPM.TaskRequest.Decline", OutlookTaskCommunicationKind.Decline)]
    [InlineData("IPM.TaskRequest.Update", OutlookTaskCommunicationKind.Update)]
    public void TaskCommunicationClassesAreRecognized(string messageClass, OutlookTaskCommunicationKind expected) {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = messageClass,
            Task = new OutlookTask()
        };

        EmailDocument result = RoundTrip(source, EmailFileFormat.OutlookMsg);

        Assert.Equal(expected, result.TaskCommunication!.Kind);
        Assert.False(result.TaskCommunication.Validate().IsValid);
        Assert.Contains(result.TaskCommunication.Validate().Issues,
            issue => issue.Code == "TASK_COMMUNICATION_PAYLOAD_MISSING");
    }

    private static EmailDocument RoundTrip(EmailDocument document, EmailFileFormat format) =>
        new EmailDocumentReader().Read(new EmailDocumentWriter().ToBytes(document, format)).Document;
}
