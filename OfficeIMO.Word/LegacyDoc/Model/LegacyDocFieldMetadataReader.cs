namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Reads field locks from the qualified native profiles' declared story-relative Plcfld records.</summary>
    internal sealed class LegacyDocFieldMetadataReader {
        private readonly Dictionary<int, FieldState> _fieldsByEnd = new();
        private readonly HashSet<int> _projectedEnds = new();
        private readonly List<(int Start, int End, string Name)> _validatedStories = new();
        private readonly List<LegacyDocUnsupportedFeature> _findings = new();
        private readonly string _tableStreamName;

        private LegacyDocFieldMetadataReader(string tableStreamName) {
            _tableStreamName = tableStreamName;
        }

        internal IReadOnlyList<LegacyDocUnsupportedFeature> Findings => _findings;
        internal int EditedCacheCount { get; private set; }

        internal static LegacyDocFieldMetadataReader Read(
            byte[] tableStream, LegacyDocFib fib, IReadOnlyList<LegacyDocTextCharacter> characters,
            string? tableStreamName = null) {
            var reader = new LegacyDocFieldMetadataReader(tableStreamName ?? (fib.UsesOneTableStream ? "1Table" : "0Table"));
            int footnoteStart = fib.CcpText;
            int headerStart = footnoteStart + fib.CcpFtn;
            int commentStart = headerStart + fib.CcpHdd;
            int endnoteStart = commentStart + fib.CcpAtn;
            reader.ReadStory(tableStream, characters, "Main", 0, fib.CcpText, fib.FcPlcfFldMom, fib.LcbPlcfFldMom);
            reader.ReadStory(tableStream, characters, "Header", headerStart, fib.CcpHdd, fib.FcPlcfFldHdr, fib.LcbPlcfFldHdr);
            reader.ReadStory(tableStream, characters, "Footnote", footnoteStart, fib.CcpFtn, fib.FcPlcfFldFtn, fib.LcbPlcfFldFtn);
            reader.ReadStory(tableStream, characters, "Comment", commentStart, fib.CcpAtn, fib.FcPlcfFldAtn, fib.LcbPlcfFldAtn);
            reader.ReadStory(tableStream, characters, "Endnote", endnoteStart, fib.CcpEdn, fib.FcPlcfFldEdn, fib.LcbPlcfFldEdn);
            return reader;
        }

        /// <summary>Uses the Word 6/95 in-stream directories only for its supported body and header stories.</summary>
        internal static LegacyDocFieldMetadataReader ReadWord69(byte[] wordDocument, int usedEnd,
            IReadOnlyList<LegacyDocTextCharacter> characters, int bodyCount, int footnoteCount, int headerCount) {
            var reader = new LegacyDocFieldMetadataReader("WordDocument");
            ReadOldStory("Main", 0, bodyCount, 0xD8);
            ReadOldStory("Header", bodyCount + footnoteCount, headerCount, 0xE0);
            return reader;

            void ReadOldStory(string story, int start, int count, int directoryOffset) {
                int offset = LegacyDocFib.ReadInt32(wordDocument, directoryOffset);
                int length = LegacyDocFib.ReadInt32(wordDocument, directoryOffset + 4);
                if (length != 0 && (offset < LegacyWord69Reader.FibSize || length < 0 || (long)offset + length > usedEnd)) {
                    reader.AddFinding("doc-field-plc-invalid", story,
                        "The Word 6/95 field PLC points outside its used WordDocument stream. Its field-lock metadata is not used.");
                    return;
                }
                reader.ReadStory(wordDocument, characters, story, start, count, offset, length);
            }
        }

        /// <summary>Matches the parsed cache's end delimiter, without treating an absent PLC as an unlocked declaration.</summary>
        internal bool? ReadLock(int endCharacterPosition) {
            if (_fieldsByEnd.TryGetValue(endCharacterPosition, out FieldState field)) {
                _projectedEnds.Add(endCharacterPosition);
                return field.Locked;
            }
            foreach (var story in _validatedStories) {
                if (endCharacterPosition >= story.Start && endCharacterPosition < story.End
                    && _projectedEnds.Add(endCharacterPosition)) {
                    AddFinding("doc-field-state-unqualified", story.Name,
                        $"The parsed field ending at global character position {endCharacterPosition} has no matching end record in its declared field PLC. Its lock state is unknown.");
                    break;
                }
            }
            return null;
        }

        internal IEnumerable<LegacyDocUnsupportedFeature> UnprojectedLocks(ISet<int>? excludedEndPositions = null) {
            foreach (var pair in _fieldsByEnd) {
                if (pair.Value.Locked && !_projectedEnds.Contains(pair.Key)
                    && excludedEndPositions?.Contains(pair.Key) != true) {
                    yield return new LegacyDocUnsupportedFeature(LegacyDocUnsupportedFeatureKind.Formatting,
                        "doc-field-lock-unprojected",
                        $"The locked field ending at global character position {pair.Key} in the {pair.Value.Story} story is not projected as an editable field. Its no-recalculation state cannot be preserved.",
                        _tableStreamName, "Plcfld:" + pair.Value.Story + ":Lock");
                }
            }
        }

        private void ReadStory(byte[] table, IReadOnlyList<LegacyDocTextCharacter> characters,
            string story, int storyStart, int storyCount, int offset, int length) {
            // With no declared PLC, retain the established text-carrier compatibility path.
            // MS-DOC also requires that an undefined fc be ignored when lcb is zero.
            if (length == 0) return;
            if (offset < 0 || length < 4 || (length - 4) % 6 != 0 || (long)offset + length > table.Length) {
                Invalid("The declared field PLC has an invalid byte range or record length.");
                return;
            }

            int count = (length - 4) / 6;
            int propertyStart = offset + (count + 1) * 4;
            int previousCp = -1;
            var frames = new Stack<FieldFrame>();
            var fields = new List<(int End, FieldState State)>();
            var unsupportedFlags = new List<(int End, byte Flags)>();
            int editedCount = 0;
            for (int index = 0; index < count; index++) {
                int cp = LegacyDocFib.ReadInt32(table, offset + index * 4);
                // fldch's high three bits are reserved and MUST be ignored.
                byte character = (byte)(table[propertyStart + index * 2] & 0x1F);
                byte flags = table[propertyStart + index * 2 + 1];
                if (cp <= previousCp || cp < 0 || cp >= storyCount
                    || !MatchesCharacter(characters, storyStart + cp, character)) {
                    Invalid("The declared field PLC has duplicate, unordered, out-of-story, or mismatched delimiter positions.");
                    return;
                }
                previousCp = cp;
                if (character == LegacyDocField.Begin) {
                    frames.Push(new FieldFrame(frames.Count != 0));
                } else if (character == LegacyDocField.Separator) {
                    // The separator's second FLD byte is unused and MUST be ignored.
                    if (frames.Count == 0 || frames.Peek().HasSeparator) {
                        Invalid("The declared field PLC has an unmatched or repeated separator.");
                        return;
                    }
                    frames.Peek().HasSeparator = true;
                } else if (character == LegacyDocField.End) {
                    if (frames.Count == 0) {
                        Invalid("The declared field PLC has an unmatched end delimiter.");
                        return;
                    }
                    FieldFrame frame = frames.Pop();
                    if (((flags & 0x80) != 0) != frame.HasSeparator || ((flags & 0x40) != 0) != frame.Nested) {
                        Invalid("The declared field PLC's separator or nesting flags disagree with its delimiter structure.");
                        return;
                    }
                    fields.Add((storyStart + cp, new FieldState((flags & 0x10) != 0, story)));
                    if ((flags & 0x0C) != 0) editedCount++;
                    if ((flags & 0x23) != 0) unsupportedFlags.Add((storyStart + cp, flags));
                } else {
                    Invalid("The declared field PLC has a character other than begin, separator, or end.");
                    return;
                }
            }

            // The terminal CP must be largest, but its value otherwise is undefined.
            // Word can write a terminal beyond this story's character count.
            int terminal = LegacyDocFib.ReadInt32(table, offset + count * 4);
            if (frames.Count != 0 || (count > 0 && terminal <= previousCp)) {
                Invalid("The declared field PLC has an unterminated field or unordered terminal position.");
                return;
            }
            _validatedStories.Add((storyStart, storyStart + storyCount, story));
            foreach (var field in fields) _fieldsByEnd.Add(field.End, field.State);
            EditedCacheCount += editedCount;
            foreach (var state in unsupportedFlags) {
                if ((state.Flags & 0x01) != 0) {
                    AddFinding("doc-field-display-approximation", story,
                        $"The field ending at global character position {state.End} in the {story} story sets fDiffer (end flags 0x{state.Flags:X2}). Its preference to invert the document's instruction/result display mode is not represented. Retaining its source instruction and cached result does not preserve that source view; the editable field view is an approximation.");
                }
                var names = new List<string>();
                if ((state.Flags & 0x02) != 0) names.Add("an OLE-bearing result (fZombieEmbed)");
                if ((state.Flags & 0x20) != 0) names.Add("a private result (fPrivateResult)");
                if (names.Count > 0) {
                    AddFinding("doc-field-state-unqualified", story,
                        $"The field ending at global character position {state.End} in the {story} story declares {string.Join(", ", names)} in end flags 0x{state.Flags:X2}. These behaviors are not qualified by the editable field projection.");
                }
            }

            void Invalid(string reason) => AddFinding("doc-field-plc-invalid", story, reason + " Its field-lock metadata is not used.");
        }

        private void AddFinding(string code, string story, string description) {
            _findings.Add(new LegacyDocUnsupportedFeature(LegacyDocUnsupportedFeatureKind.Formatting,
                code, description, _tableStreamName, "Plcfld:" + story));
        }

        private static bool MatchesCharacter(IReadOnlyList<LegacyDocTextCharacter> characters, int cp, byte value) {
            int first = 0;
            int last = characters.Count;
            while (first < last) {
                int middle = first + (last - first) / 2;
                if (characters[middle].CharacterPosition < cp) first = middle + 1;
                else last = middle;
            }
            return first < characters.Count && characters[first].CharacterPosition == cp && characters[first].Character == value;
        }

        private readonly struct FieldState {
            internal FieldState(bool locked, string story) {
                Locked = locked;
                Story = story;
            }
            internal bool Locked { get; }
            internal string Story { get; }
        }

        private sealed class FieldFrame {
            internal FieldFrame(bool nested) { Nested = nested; }
            internal bool Nested { get; }
            internal bool HasSeparator { get; set; }
        }
    }
}
