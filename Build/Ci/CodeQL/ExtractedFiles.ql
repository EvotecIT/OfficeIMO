/** Extracted source inventory, including compiler-generated C# and line counts. */
import csharp

from SourceFile file
select file.getAbsolutePath(), sum(int lines | lines = file.getNumberOfLinesOfCode() | lines)
