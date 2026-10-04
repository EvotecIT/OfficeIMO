using System.Threading;

namespace OfficeIMO.Latex;

/// <summary>Lexical control-word whitespace rules shared by binding and literal projection.</summary>
internal static class LatexWhitespaceSyntax {
    internal static int SkipControlWordDelimiter(string source, int cursor, int end, CancellationToken cancellationToken) {
        bool lineEnded = false;
        while (cursor < end) {
            if ((cursor & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = source[cursor];
            if (current == ' ' || current == '\t') { cursor++; continue; }
            if (current == '%') {
                while (cursor < end && source[cursor] != '\r' && source[cursor] != '\n') {
                    if ((cursor & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    cursor++;
                }
                if (cursor < end && source[cursor] == '\r') cursor++;
                if (cursor < end && source[cursor] == '\n') cursor++;
                continue;
            }
            if (current != '\r' && current != '\n' || lineEnded) break;
            lineEnded = true;
            if (current == '\r' && cursor + 1 < end && source[cursor + 1] == '\n') cursor++;
            cursor++;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return cursor;
    }
}
