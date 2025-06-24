using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_CustomList1(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Custom Lists 2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("This is 1st list");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                // create list with bulleted style
                // we could have used CustomStyle, but we are using Bulleted to show removal of levels
                WordList wordList1 = document.AddList(WordListStyle.Bulleted);
                wordList1.AddItem("Text 1 - List1");
                wordList1.AddItem("Text 2 - List1", 1);
                wordList1.AddItem("Text 3 - List1", 2);
                wordList1.AddItem("Text 4 - List1", 3);
                wordList1.AddItem("Text 5 - List1", 4);
                wordList1.AddItem("Text 6 - List1", 5);
                wordList1.AddItem("Text 7 - List1", 6);
                wordList1.AddItem("Text 8 - List1", 7);
                wordList1.AddItem("Text 9 - List1", 8);

                // let's display some properties of the list
                Console.WriteLine(wordList1.Numbering.Levels[0]._level.LevelIndex.ToString());
                Console.WriteLine(wordList1.Numbering.Levels[1]._level.LevelIndex.ToString());
                Console.WriteLine(wordList1.Numbering.Levels[2]._level.LevelIndex.ToString());
                Console.WriteLine(wordList1.Numbering.Levels[0].IndentationHanging);
                Console.WriteLine(wordList1.Numbering.Levels[0].IndentationLeft);
                Console.WriteLine(wordList1.Numbering.Levels[0].IndentationLeftCentimeters);
                Console.WriteLine(wordList1.Numbering.Levels[1].IndentationLeftCentimeters);
                Console.WriteLine(wordList1.Numbering.Levels[2].IndentationLeftCentimeters);
                Console.WriteLine(wordList1.Numbering.Levels[3].IndentationLeftCentimeters);
                Console.WriteLine(wordList1.Numbering.Levels[1].LevelJustification);
                Console.WriteLine(wordList1.Numbering.Levels[1].StartNumberingValue);

                Console.WriteLine("Current levels count: " + wordList1.Numbering.Levels.Count);

                // remove single level
                wordList1.Numbering.Levels[0].Remove();
                // remove all levels
                wordList1.Numbering.RemoveAllLevels();

                Console.WriteLine("Current levels count: " + wordList1.Numbering.Levels.Count);

                // add custom levels
                var level = new WordListLevel(SimplifiedListNumbers.BulletOpenCircle);
                wordList1.Numbering.AddLevel(level);

                var level1 = new WordListLevel(SimplifiedListNumbers.BulletSolidRound);
                wordList1.Numbering.AddLevel(level1);

                var level2 = new WordListLevel(SimplifiedListNumbers.BulletSquare);
                wordList1.Numbering.AddLevel(level2);

                var level3 = new WordListLevel(SimplifiedListNumbers.BulletSquare2);
                wordList1.Numbering.AddLevel(level3);

                var level4 = new WordListLevel(SimplifiedListNumbers.BulletClubs);
                wordList1.Numbering.AddLevel(level4);

                var level5 = new WordListLevel(SimplifiedListNumbers.BulletDiamond);
                wordList1.Numbering.AddLevel(level5);

                var level6 = new WordListLevel(SimplifiedListNumbers.BulletCheckmark);
                wordList1.Numbering.AddLevel(level6);

                var level7 = new WordListLevel(SimplifiedListNumbers.BulletArrow);
                wordList1.Numbering.AddLevel(level7);

                Console.WriteLine("Current levels count: " + wordList1.Numbering.Levels.Count);

                document.AddParagraph("This is 2nd list");

                // define list
                // prefer AddCustomList over AddList(WordListStyle.Custom)
                WordList wordList2 = document.AddCustomList();
                // add levels
                var level21 = new WordListLevel(SimplifiedListNumbers.Decimal);
                wordList2.Numbering.AddLevel(level21);
                var level22 = new WordListLevel(SimplifiedListNumbers.DecimalBracket);
                wordList2.Numbering.AddLevel(level22);
                var level23 = new WordListLevel(SimplifiedListNumbers.DecimalDot);
                wordList2.Numbering.AddLevel(level23);
                var level24 = new WordListLevel(SimplifiedListNumbers.LowerLetter);
                wordList2.Numbering.AddLevel(level24);
                var level25 = new WordListLevel(SimplifiedListNumbers.LowerLetterBracket);
                wordList2.Numbering.AddLevel(level25);
                var level26 = new WordListLevel(SimplifiedListNumbers.LowerLetterDot);
                wordList2.Numbering.AddLevel(level26);
                var level27 = new WordListLevel(SimplifiedListNumbers.UpperLetter);
                wordList2.Numbering.AddLevel(level27);
                var level28 = new WordListLevel(SimplifiedListNumbers.UpperLetterBracket);
                wordList2.Numbering.AddLevel(level28);
                var level29 = new WordListLevel(SimplifiedListNumbers.UpperLetterDot);
                wordList2.Numbering.AddLevel(level29);

                // add items to the list
                wordList2.AddItem("Text 1 - Decimal");
                wordList2.AddItem("Text 2 - DecimalBracket", 1);
                wordList2.AddItem("Text 3 - DecimalDot", 2);
                wordList2.AddItem("Text 4 - LowerLetter", 3);
                wordList2.AddItem("Text 4 - LowerLetter", 3);
                wordList2.AddItem("Text 4 - LowerLetter", 3);
                wordList2.AddItem("Text 5.1 - LowerLetterBracket", 4);
                wordList2.AddItem("Text 5.2 - LowerLetterBracket", 4);
                wordList2.AddItem("Text 5.3 - LowerLetterBracket", 4);
                wordList2.AddItem("Text 6 - LowerLetterDot", 5);
                wordList2.AddItem("Text 7 - UpperLetter", 6);
                wordList2.AddItem("Text 8 - UpperLetterBracket", 7);
                wordList2.AddItem("Text 9 - UpperLetterDot", 8);

                document.AddParagraph("This is 3rd list");

                // another custom list
                WordList wordList3 = document.AddCustomList();
                var level31 = new WordListLevel(SimplifiedListNumbers.UpperRoman);
                wordList3.Numbering.AddLevel(level31);
                var level32 = new WordListLevel(SimplifiedListNumbers.UpperRomanBracket);
                wordList3.Numbering.AddLevel(level32);
                var level33 = new WordListLevel(SimplifiedListNumbers.UpperRomanDot);
                wordList3.Numbering.AddLevel(level33);
                var level34 = new WordListLevel(SimplifiedListNumbers.LowerRoman);
                level34.StartNumberingValue = 4;
                level34.LevelJustification = LevelJustificationValues.Right;
                wordList3.Numbering.AddLevel(level34);
                var level35 = new WordListLevel(SimplifiedListNumbers.LowerRomanBracket);
                wordList3.Numbering.AddLevel(level35);
                var level36 = new WordListLevel(SimplifiedListNumbers.LowerRomanDot);
                wordList3.Numbering.AddLevel(level36);
                var level37 = new WordListLevel(SimplifiedListNumbers.DecimalBracket);
                wordList3.Numbering.AddLevel(level37);
                var level38 = new WordListLevel(SimplifiedListNumbers.DecimalDot);
                wordList3.Numbering.AddLevel(level38);
                var level39 = new WordListLevel(SimplifiedListNumbers.Decimal);
                wordList3.Numbering.AddLevel(level39);


                wordList3.AddItem("Text 1 - UpperRoman");
                wordList3.AddItem("Text 2 - UpperRomanBracket", 1);
                wordList3.AddItem("Text 3 - UpperRomanDot", 2);
                wordList3.AddItem("Text 4 - LowerRoman", 3);
                wordList3.AddItem("Text 5 - LowerRomanBracket", 4);
                wordList3.AddItem("Text 6 - LowerRomanDot", 5);
                wordList3.AddItem("Text 7 - DecimalBracket", 6);
                wordList3.AddItem("Text 8 - DecimalDot", 7);
                wordList3.AddItem("Text 9 - Decimal", 8);

                document.Save(openWord);
            }
        }
    }
}
