using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static IEnumerable<ExcelThreadedCommentOptions> RootComments(IWorkTable table, CancellationToken cancellationToken) {
        foreach (IWorkTableCell cell in table.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.Comment is not { } comment) continue;
            yield return new ExcelThreadedCommentOptions {
                Address = A1.CellReference(cell.Row, cell.Column), Text = comment.Text,
                Author = comment.Author, Date = comment.CreationDateUtc
            };
        }
    }
}
