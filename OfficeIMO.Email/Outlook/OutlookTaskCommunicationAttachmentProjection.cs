namespace OfficeIMO.Email;

internal static class OutlookTaskCommunicationAttachmentProjection {
    internal static EmailAttachment[] GetWritableAttachments(EmailDocument document) {
        EmailAttachment[] attachments = document.Attachments
            .Where(attachment => !attachment.IsProjectedSemanticContent)
            .ToArray();
        OutlookTaskCommunication? communication = document.TaskCommunication;
        if (communication == null || communication.Kind == OutlookTaskCommunicationKind.None) return attachments;

        OutlookTaskCommunicationValidationReport validation = communication.Validate();
        if (!validation.IsValid && communication.EmbeddedTask == null) {
            string detail = string.Join("; ", validation.Issues.Where(issue => issue.IsError)
                .Select(issue => string.Concat(issue.Code, ": ", issue.Message)));
            throw new InvalidOperationException(string.Concat("The task communication is invalid. ", detail));
        }

        EmailDocument embeddedTask = communication.EmbeddedTask!;
        if (embeddedTask.Task == null)
            throw new InvalidOperationException("The task communication payload does not contain an Outlook task.");
        if (!embeddedTask.Task.GlobalId.HasValue) embeddedTask.Task.GlobalId = Guid.NewGuid();

        EmailAttachment? source = communication.PayloadAttachment;
        if (source == null || !attachments.Contains(source)) {
            source = attachments.FirstOrDefault(attachment =>
                ReferenceEquals(attachment.EmbeddedDocument, embeddedTask));
        }
        var result = new List<EmailAttachment>(attachments.Length + (source == null ? 1 : 0)) {
            CreateCanonicalAttachment(source, embeddedTask)
        };
        result.AddRange(attachments.Where(attachment => !ReferenceEquals(attachment, source)));
        return result.ToArray();
    }

    private static EmailAttachment CreateCanonicalAttachment(EmailAttachment? source, EmailDocument embeddedTask) {
        var attachment = new EmailAttachment {
            FileName = source?.FileName,
            ContentType = source?.ContentType,
            ContentId = source?.ContentId,
            ContentLocation = source?.ContentLocation,
            IsInline = false,
            IsMimeRelated = source?.IsMimeRelated ?? false,
            IsHidden = true,
            IsContactPhoto = false,
            RenderingPosition = -1,
            CreatedDate = source?.CreatedDate,
            ModifiedDate = source?.ModifiedDate,
            Length = source?.Length ?? 0,
            EmbeddedDocument = embeddedTask,
            MapiAttachMethod = 5
        };
        if (source == null) return attachment;
        foreach (KeyValuePair<string, string> parameter in source.ContentTypeParameters)
            attachment.ContentTypeParameters[parameter.Key] = parameter.Value;
        foreach (MapiProperty property in source.MapiProperties) attachment.MapiProperties.Add(Clone(property));
        foreach (TnefAttribute attribute in source.TnefAttributes) attachment.TnefAttributes.Add(attribute);
        return attachment;
    }

    private static MapiProperty Clone(MapiProperty property) =>
        new MapiProperty(property.PropertyId, property.PropertyType, property.Value, property.Flags, property.Name) {
            RawData = property.RawData == null ? null : (byte[])property.RawData.Clone()
        };
}
