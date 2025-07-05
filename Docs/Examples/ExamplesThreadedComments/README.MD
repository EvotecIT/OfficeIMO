## Working with threaded comments

OfficeIMO supports threaded comments that mirror the model used in Word. Replies are linked to their parent via unique `ParaId` identifiers.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var paragraph = document.AddParagraph("Main paragraph with comment");
    var top = paragraph.AddComment("Author1", "A1", "Top level comment");
    top.AddReply("Author2", "A2", "First reply");
    var second = top.AddReply("Author3", "A3", "Second reply");
    second.AddReply("Author2", "A2", "Nested reply");
    document.Save();
}
```

Each `WordComment` exposes `ParentComment` and `Replies` collections so you can traverse the thread:

```csharp
foreach (var comment in document.Comments) {
    Console.WriteLine($"{comment.ParaId}: {comment.Text}");
    if (comment.ParentComment != null) {
        Console.WriteLine($"  Parent -> {comment.ParentComment.ParaId}");
    }
    foreach (var reply in comment.Replies) {
        Console.WriteLine($"  Reply -> {reply.ParaId}: {reply.Text}");
    }
}
```
