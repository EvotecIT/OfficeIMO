using System;
using System.Collections.Generic;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Word
    {
        public class FieldParser_Should
        {

            [Theory]
            [InlineData("ASK bookmark question", 0)]
            [InlineData(@"BIBLIOGRAPHY \* roman", 1)]
            [InlineData(@"BIBLIOGRAPHY \*arabic", 1)]
            [InlineData(@"Page \* FIRSTCAP \* MERGEFORMAT", 2)]
            public void Test_IdentifyFormatSwitches(String FieldCodeString, int expected_amount_of_format_switches)
            {
                var parser = new WordFieldParser(FieldCodeString);

                Assert.Equal(expected_amount_of_format_switches, parser.FormatSwitches.Count);
            }

            [Theory]
            [InlineData(@"BIBLIOGRAPHY \* roman", WordFieldFormat.Roman)]
            [InlineData(@"BIBLIOGRAPHY \*arabic", WordFieldFormat.Arabic)]
            [InlineData(@"Page \* FIRSTCAP \* MERGEFORMAT", WordFieldFormat.FirstCap)]
            public void Test_CastFormatSwitches(String FieldCodeString, WordFieldFormat expected_field_format)
            {
                var parser = new WordFieldParser(FieldCodeString);

                Assert.Contains(expected_field_format, parser.FormatSwitches);
            }

            [Theory]
            [InlineData("BIBLIOGRAPHY \\@ \"roman\"", 1)]
            [InlineData("BIBLIOGRAPHY \\h \"some senctence\"", 1)]
            [InlineData(@"BIBLIOGRAPHY \t test \*arabic", 1)]
            [InlineData(@"Page \h \d something \* MERGEFORMAT", 2)]
            public void Test_IdentifySwitches(String FieldCodeString, int expected_amount_of_switches)
            {
                var parser = new WordFieldParser(FieldCodeString);

                Assert.Equal(expected_amount_of_switches, parser.Switches.Count);
            }

            [Theory]
            [InlineData("BIBLIOGRAPHY \\@ \"roman\"", new [] { "\\@ \"roman\""})]
            [InlineData("BIBLIOGRAPHY \\h \"some senctence\"", new[] {"\\h \"some senctence\"" })]
            [InlineData(@"BIBLIOGRAPHY \t test \*arabic", new[] {"\\t test" })]
            [InlineData(@"Page \h \d something \* MERGEFORMAT", new[] {"\\h","\\d something"})]
            public void Test_ParseSwitches(String FieldCodeString, String[] expected_switches)
            {
                var parser = new WordFieldParser(FieldCodeString);

                Assert.Equal(expected_switches, parser.Switches);
            }

            [Theory]
            [InlineData("BIBLIOGRAPHY \\@ \"roman\"", 0)]
            [InlineData("BIBLIOGRAPHY \\h \"some senctence\"", 0)]
            [InlineData(@"BIBLIOGRAPHY \t test \*arabic", 0)]
            [InlineData(@"Page \h \d something \* MERGEFORMAT", 0)]
            [InlineData(@"Page this are several instructions \d something \* MERGEFORMAT", 4)]
            [InlineData("Page \"this is a single instruction \" \\h \\d something \\* MERGEFORMAT", 1)]
            [InlineData("PAGE \"this is a single instruction \" another_one \\h \\d something \\* MERGEFORMAT", 2)]
            public void Test_IdentifyInstructions(String FieldCodeString, int expected_amount_of_instructions)
            {
                var parser = new WordFieldParser(FieldCodeString);

                Assert.Equal(expected_amount_of_instructions, parser.Instructions.Count);
            }

            [Fact]
            public void Test_InformAboutUnprocessableStrings()
            {
                Assert.Throws<NotImplementedException>(() => new WordFieldParser("SillyField not known"));
            }

        }
    }
}

