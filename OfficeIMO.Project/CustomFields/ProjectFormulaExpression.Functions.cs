namespace OfficeIMO.Project;

internal sealed partial class ProjectFormulaExpression {
    private ProjectFormulaValue Call(string name, List<ProjectFormulaValue> args) {
        Step(); string key = name.ToLowerInvariant();
        void Count(int minimum, int maximum) { if (args.Count < minimum || args.Count > maximum) throw Invalid("Unexpected argument count for " + name); }
        switch (key) {
            case "iif": Count(3, 3); return args[0].FlagIn(_culture) ? args[1] : args[2];
            case "abs": Count(1, 1); return new ProjectFormulaValue(Math.Abs(args[0].NumberIn(_culture)));
            case "round":
                Count(1, 2); int digits = args.Count == 1 ? 0 : args[1].IntegerIn(_culture);
                if (digits < 0 || digits > 28) throw Invalid("Round precision must be between zero and 28");
                return new ProjectFormulaValue(decimal.Round(args[0].NumberIn(_culture), digits, MidpointRounding.ToEven));
            case "int": Count(1, 1); return new ProjectFormulaValue(decimal.Floor(args[0].NumberIn(_culture)));
            case "fix": Count(1, 1); return new ProjectFormulaValue(decimal.Truncate(args[0].NumberIn(_culture)));
            case "cstr": Count(1, 1); return new ProjectFormulaValue(args[0].TextIn(_culture));
            case "cbool": Count(1, 1); return new ProjectFormulaValue(args[0].FlagIn(_culture));
            case "ucase": Count(1, 1); return new ProjectFormulaValue(args[0].TextIn(_culture).ToUpper(_culture));
            case "lcase": Count(1, 1); return new ProjectFormulaValue(args[0].TextIn(_culture).ToLower(_culture));
            case "trim": Count(1, 1); return new ProjectFormulaValue(args[0].TextIn(_culture).Trim(' '));
            case "len": Count(1, 1); return new ProjectFormulaValue((decimal)args[0].TextIn(_culture).Length);
            case "left": case "right":
                Count(2, 2); string text = args[0].TextIn(_culture); int length = args[1].IntegerIn(_culture);
                if (length < 0) throw Invalid("Text length cannot be negative");
                length = Math.Min(length, text.Length);
                return new ProjectFormulaValue(key == "left" ? text.Substring(0, length) : text.Substring(text.Length - length));
            case "mid":
                Count(2, 3); string source = args[0].TextIn(_culture); int from = args[1].IntegerIn(_culture);
                if (from < 1) throw Invalid("Mid uses a positive, one-based start");
                int take = args.Count == 3 ? args[2].IntegerIn(_culture) : source.Length;
                if (take < 0) throw Invalid("Text length cannot be negative");
                return new ProjectFormulaValue(from > source.Length ? "" : source.Substring(from - 1, Math.Min(take, source.Length - from + 1)));
            case "year": case "month": case "day": case "hour": case "minute": case "second":
                Count(1, 1); DateTime date = args[0].Date;
                return new ProjectFormulaValue((decimal)(key switch {
                    "year" => date.Year, "month" => date.Month, "day" => date.Day, "hour" => date.Hour, "minute" => date.Minute, _ => date.Second
                }));
            case "dateadd":
                Count(3, 3); string unit = args[0].TextIn(_culture).ToLower(_culture);
                int amount = checked((int)decimal.Round(args[1].NumberIn(_culture), 0, MidpointRounding.ToEven)); DateTime original = args[2].Date;
                return new ProjectFormulaValue(unit switch {
                    "yyyy" => original.AddYears(amount), "q" => original.AddMonths(checked(amount * 3)), "m" => original.AddMonths(amount),
                    "d" or "y" or "w" => original.AddDays(amount), "ww" => original.AddDays(checked(amount * 7)),
                    "h" => original.AddHours(amount), "n" => original.AddMinutes(amount), "s" => original.AddSeconds(amount),
                    _ => throw new NotSupportedException("DateAdd interval is outside the supported profile.")
                });
            default: throw new NotSupportedException("Formula function '" + name + "' is outside the supported profile.");
        }
    }
}
