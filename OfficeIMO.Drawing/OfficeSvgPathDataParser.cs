using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

internal static class OfficeSvgPathDataParser {
    private const int MaximumCommands = 100000;

    internal static bool TryParse(string? data, out IReadOnlyList<OfficePathCommand> commands) {
        var result = new List<OfficePathCommand>();
        commands = result;
        if (string.IsNullOrWhiteSpace(data)) return false;

        var reader = new PathReader(data!);
        char command = '\0';
        OfficePoint current = default;
        OfficePoint subpathStart = default;
        OfficePoint lastCubicControl = default;
        OfficePoint lastQuadraticControl = default;
        char previousCommand = '\0';
        bool hasCurrent = false;
        bool hasDraw = false;

        while (reader.SkipSeparators()) {
            if (result.Count >= MaximumCommands) return false;
            if (reader.TryReadCommand(out char explicitCommand)) command = explicitCommand;
            else if (command == '\0' || command is 'Z' or 'z') return false;
            if (command is 'A' or 'a') return false;

            bool relative = char.IsLower(command);
            char upper = char.ToUpperInvariant(command);
            if (upper == 'Z') {
                if (!hasCurrent) return false;
                result.Add(OfficePathCommand.Close());
                current = subpathStart;
                hasDraw = true;
                previousCommand = command;
                command = '\0';
                continue;
            }

            int groups = 0;
            while (reader.HasNumberAhead()) {
                if (result.Count >= MaximumCommands) return false;
                switch (upper) {
                    case 'M':
                    case 'L':
                    case 'T':
                        if (!reader.TryReadPoint(out OfficePoint point)) return false;
                        point = Resolve(point, current, relative);
                        if (upper == 'M' && groups == 0) {
                            result.Add(OfficePathCommand.MoveTo(point));
                            subpathStart = point;
                            hasCurrent = true;
                        } else if (upper == 'T') {
                            if (!hasCurrent) return false;
                            OfficePoint control = previousCommand is 'Q' or 'q' or 'T' or 't'
                                ? Reflect(lastQuadraticControl, current)
                                : current;
                            result.Add(OfficePathCommand.QuadraticBezierTo(control, point));
                            lastQuadraticControl = control;
                            hasDraw = true;
                        } else {
                            if (!hasCurrent) return false;
                            result.Add(OfficePathCommand.LineTo(point));
                            hasDraw = true;
                        }
                        current = point;
                        break;
                    case 'H':
                        if (!hasCurrent || !reader.TryReadNumber(out double horizontal)) return false;
                        current = new OfficePoint(relative ? current.X + horizontal : horizontal, current.Y);
                        result.Add(OfficePathCommand.LineTo(current));
                        hasDraw = true;
                        break;
                    case 'V':
                        if (!hasCurrent || !reader.TryReadNumber(out double vertical)) return false;
                        current = new OfficePoint(current.X, relative ? current.Y + vertical : vertical);
                        result.Add(OfficePathCommand.LineTo(current));
                        hasDraw = true;
                        break;
                    case 'C':
                        if (!hasCurrent
                            || !reader.TryReadPoint(out OfficePoint cubic1)
                            || !reader.TryReadPoint(out OfficePoint cubic2)
                            || !reader.TryReadPoint(out OfficePoint cubicEnd)) return false;
                        cubic1 = Resolve(cubic1, current, relative);
                        cubic2 = Resolve(cubic2, current, relative);
                        cubicEnd = Resolve(cubicEnd, current, relative);
                        result.Add(OfficePathCommand.CubicBezierTo(cubic1, cubic2, cubicEnd));
                        lastCubicControl = cubic2;
                        current = cubicEnd;
                        hasDraw = true;
                        break;
                    case 'S':
                        if (!hasCurrent
                            || !reader.TryReadPoint(out OfficePoint smooth2)
                            || !reader.TryReadPoint(out OfficePoint smoothEnd)) return false;
                        smooth2 = Resolve(smooth2, current, relative);
                        smoothEnd = Resolve(smoothEnd, current, relative);
                        OfficePoint smooth1 = previousCommand is 'C' or 'c' or 'S' or 's'
                            ? Reflect(lastCubicControl, current)
                            : current;
                        result.Add(OfficePathCommand.CubicBezierTo(smooth1, smooth2, smoothEnd));
                        lastCubicControl = smooth2;
                        current = smoothEnd;
                        hasDraw = true;
                        break;
                    case 'Q':
                        if (!hasCurrent
                            || !reader.TryReadPoint(out OfficePoint quadraticControl)
                            || !reader.TryReadPoint(out OfficePoint quadraticEnd)) return false;
                        quadraticControl = Resolve(quadraticControl, current, relative);
                        quadraticEnd = Resolve(quadraticEnd, current, relative);
                        result.Add(OfficePathCommand.QuadraticBezierTo(quadraticControl, quadraticEnd));
                        lastQuadraticControl = quadraticControl;
                        current = quadraticEnd;
                        hasDraw = true;
                        break;
                    default:
                        return false;
                }

                groups++;
                previousCommand = command;
                reader.SkipSeparators();
                if (reader.HasCommandAhead()) break;
                if (upper == 'M') upper = 'L';
            }

            if (groups == 0) return false;
        }

        return hasCurrent && hasDraw && result.Count >= 2;
    }

    private static OfficePoint Resolve(OfficePoint point, OfficePoint current, bool relative) =>
        relative ? new OfficePoint(current.X + point.X, current.Y + point.Y) : point;

    private static OfficePoint Reflect(OfficePoint control, OfficePoint around) =>
        new OfficePoint((around.X * 2D) - control.X, (around.Y * 2D) - control.Y);

    private sealed class PathReader {
        private readonly string _value;
        private int _index;

        internal PathReader(string value) {
            _value = value;
        }

        internal bool SkipSeparators() {
            while (_index < _value.Length && (char.IsWhiteSpace(_value[_index]) || _value[_index] == ',')) _index++;
            return _index < _value.Length;
        }

        internal bool HasCommandAhead() => _index < _value.Length && IsCommand(_value[_index]);

        internal bool HasNumberAhead() {
            SkipSeparators();
            return _index < _value.Length && (_value[_index] == '+' || _value[_index] == '-' || _value[_index] == '.' || char.IsDigit(_value[_index]));
        }

        internal bool TryReadCommand(out char command) {
            command = '\0';
            if (!HasCommandAhead()) return false;
            command = _value[_index++];
            return true;
        }

        internal bool TryReadPoint(out OfficePoint point) {
            point = default;
            if (!TryReadNumber(out double x) || !TryReadNumber(out double y)) return false;
            point = new OfficePoint(x, y);
            return true;
        }

        internal bool TryReadNumber(out double number) {
            number = 0D;
            SkipSeparators();
            if (_index >= _value.Length) return false;
            int start = _index;
            if (_value[_index] is '+' or '-') _index++;
            bool digits = false;
            while (_index < _value.Length && char.IsDigit(_value[_index])) {
                digits = true;
                _index++;
            }
            if (_index < _value.Length && _value[_index] == '.') {
                _index++;
                while (_index < _value.Length && char.IsDigit(_value[_index])) {
                    digits = true;
                    _index++;
                }
            }
            if (!digits) {
                _index = start;
                return false;
            }
            if (_index < _value.Length && (_value[_index] is 'e' or 'E')) {
                int exponent = _index++;
                if (_index < _value.Length && (_value[_index] is '+' or '-')) _index++;
                int exponentDigits = _index;
                while (_index < _value.Length && char.IsDigit(_value[_index])) _index++;
                if (_index == exponentDigits) _index = exponent;
            }
            return double.TryParse(_value.Substring(start, _index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out number)
                && !double.IsNaN(number)
                && !double.IsInfinity(number);
        }

        private static bool IsCommand(char value) => value is 'M' or 'm' or 'L' or 'l' or 'H' or 'h' or 'V' or 'v'
            or 'C' or 'c' or 'S' or 's' or 'Q' or 'q' or 'T' or 't' or 'A' or 'a' or 'Z' or 'z';
    }
}
