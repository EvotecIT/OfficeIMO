namespace OfficeIMO.Html;

public static partial class HtmlProvenance {
    private static int FindHtmlCommentEnd(string html, int offset) {
        HtmlCommentScanState state = HtmlCommentScanState.Start;
        for (int index = offset; index < html.Length; index++) {
            char value = html[index];
            switch (state) {
                case HtmlCommentScanState.Start:
                    if (value == '>') return index + 1;
                    state = value == '-' ? HtmlCommentScanState.StartDash : HtmlCommentScanState.Comment;
                    break;
                case HtmlCommentScanState.StartDash:
                    if (value == '>') return index + 1;
                    state = value == '-' ? HtmlCommentScanState.End : HtmlCommentScanState.Comment;
                    break;
                case HtmlCommentScanState.Comment:
                    if (value == '-') state = HtmlCommentScanState.EndDash;
                    break;
                case HtmlCommentScanState.EndDash:
                    state = value == '-' ? HtmlCommentScanState.End : HtmlCommentScanState.Comment;
                    break;
                case HtmlCommentScanState.End:
                    if (value == '>') return index + 1;
                    if (value == '!') state = HtmlCommentScanState.EndBang;
                    else if (value != '-') state = HtmlCommentScanState.Comment;
                    break;
                case HtmlCommentScanState.EndBang:
                    if (value == '>') return index + 1;
                    state = value == '-' ? HtmlCommentScanState.EndDash : HtmlCommentScanState.Comment;
                    break;
            }
        }
        return -1;
    }

    private enum HtmlCommentScanState {
        Start,
        StartDash,
        Comment,
        EndDash,
        End,
        EndBang
    }
}
