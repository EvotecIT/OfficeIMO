using System.Runtime.CompilerServices;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;

[assembly: InternalsVisibleTo("OfficeIMo.Tests")]

namespace OfficeIMO.Word {
    internal class WordFieldParser {
        private WordFieldType _wordFieldType;
        public WordFieldType WordFieldType {
            get { return _wordFieldType; }
        }

        private readonly List<WordFieldFormat> _formatSwitches = new();
        public List<WordFieldFormat> FormatSwitches {
            get { return _formatSwitches; }
        }

        private readonly List<String> _switches = new();
        public List<String> Switches {
            get { return _switches; }
        }

        private readonly List<String> _instructions = new();
        public List<String> Instructions {
            get { return _instructions; }
        }

        /// <summary>
        /// Fieldcodes consist of several components - similar to URLs - but OpenXML only handles the whole string.
        /// This class provides an type-based interpretation of those strings.
        /// The general syntax is:
        /// <code>FIELDCODE [Instructions] [switches] [format_switches]</code>
        ///
        /// The fieldcode is a string without whitespace
        ///
        /// The instructions are optionally, though some field codes (like ask) require them.
        /// Instructions are positional statements.
        ///
        /// Switches can be found alone (like <code>\h</code>) or with a
        /// parameter (like <code>\h 1223</code> or
        /// like <code>\t "what so ever"</code>)
        ///
        /// Format switches starts with <code>\*</code> or <code>\#</code>.
        /// The latter is not implemented yet. It appears only the last format
        /// switch will be interpreted by Word - apart from <code>MERGEFORMAT</code>.
        /// However, in the effort to keep all information, all switches will be exported.
        /// </summary>
        /// <see href="https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51"/>
        /// <see href="https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fieldcode?view=openxml-2.8.1"/>
        /// <see href="https://office-watch.com/2022/all-named-format-switches-word-field-codes/"/>
        /// <see href="https://office-watch.com/2022/word-number-field-code-format-explained/"/>
        /// <param name="fieldCodeDeclaration"></param>
        public WordFieldParser(String fieldCodeDeclaration) {
            this.ParseFieldCodeDeclaration(fieldCodeDeclaration);
        }

        private void ParseFieldCodeDeclaration(String fieldCodeDeclaration) {

            fieldCodeDeclaration = fieldCodeDeclaration.Trim();


            // get format switches
            string formatSwitches = @"(\\\*) *([A-Za-z-_]+ *)";

            Regex rgx = new Regex(formatSwitches);
            var matches = rgx.Matches(fieldCodeDeclaration);
            foreach (Match m in matches) {
                var success = Enum.TryParse(m.Groups[2].ToString(), true, out WordFieldFormat fieldFormat);
                if (success) {
                    this._formatSwitches.Add(fieldFormat);

                    fieldCodeDeclaration = fieldCodeDeclaration.Replace(m.ToString(), "").Trim();
                }
            }

            // get normal switches
            string switches = "(\\\\)([A-Za-z@]{1} *([A-Za-z0-9/]+|\".+\")?)";

            rgx = new Regex(switches);
            matches = rgx.Matches(fieldCodeDeclaration);
            foreach (Match m in matches) {
                var normalSwitch = m.ToString().Trim();
                this._switches.Add(normalSwitch);

                fieldCodeDeclaration = fieldCodeDeclaration.Replace(normalSwitch.ToString(), "").Trim();
            }

            // get instructions
            string instructions = " +([A-Za-z0-9_-]+|\"[^\"]+\")";

            rgx = new Regex(instructions);
            matches = rgx.Matches(fieldCodeDeclaration);
            foreach (Match m in matches) {
                var instruction = m.ToString().Trim();
                this._instructions.Add(instruction);

                fieldCodeDeclaration = fieldCodeDeclaration.Replace(instruction.ToString(), "").Trim();
            }


            // get field type
            string fieldType = "^[^ ]+";

            rgx = new Regex(fieldType);
            String match = rgx.Match(fieldCodeDeclaration).ToString().Trim();


            var parsed = Enum.TryParse(match, true, out WordFieldType wordFieldType);
            if (parsed) {
                this._wordFieldType = wordFieldType;
                fieldCodeDeclaration = fieldCodeDeclaration.Replace(match.ToString(), "").Trim();
            }

            // No more leftovers
            if (fieldCodeDeclaration.Length > 0) {
                throw new NotImplementedException("The missing parts of the field code \"" + fieldCodeDeclaration + "\" couldn't be processed by the Parser");
            }
        }
    }
}