using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class SmartArt {
        // Exact-parity custom template 1: full text edit flow
        internal static void Example_EditCustomSmartArt1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Editing Custom SmartArt 1 (parity) with text + formatting");
            string filePath = Path.Combine(folderPath, "SmartArtCustom1_Edit.docx");
            using WordDocument document = WordDocument.Create(filePath);

            var sa = document.AddSmartArt(SmartArtType.CustomSmartArt1);
            // Replace all node texts in order
            sa.ReplaceTexts("Plan", "Build", "Test", "Deploy", "Monitor");
            // Emphasize a node with newline + formatting (use a safe index)
            if (sa.NodeCount > 0) {
                var idx = Math.Min(0, sa.NodeCount - 1);
                sa.SetNodeText(idx, "Test\nQA", bold: true, italic: false, underline: true, colorHex: "#FF0000", sizePt: 12);
            }

            document.Save(openWord);
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }

        // Exact-parity custom template 2: full text edit flow
        internal static void Example_EditCustomSmartArt2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Editing Custom SmartArt 2 (parity) with text + formatting");
            string filePath = Path.Combine(folderPath, "SmartArtCustom2_Edit.docx");
            using WordDocument document = WordDocument.Create(filePath);

            var sa = document.AddSmartArt(SmartArtType.CustomSmartArt2);
            sa.ReplaceTexts("Discover", "Design", "Develop", "Deliver", "Delight");
            if (sa.NodeCount > 0) {
                var last = sa.NodeCount - 1;
                sa.SetNodeText(last, "Delight\nFeedback", bold: false, italic: true, underline: false, colorHex: "#0066CC", sizePt: 11);
            }

            document.Save(openWord);
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }

        // Flexible algorithmic layout: BasicProcess with add/insert/remove and text edits
        internal static void Example_FlexibleBasicSmartArt_FullFlow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Flexible BasicProcess SmartArt: add/insert/remove + text edits");
            string filePath = Path.Combine(folderPath, "SmartArtBasic_Flexible.docx");
            using WordDocument document = WordDocument.Create(filePath);

            var sa = document.AddSmartArt(SmartArtType.BasicProcess);
            // Start with extra nodes
            sa.AddNode("Step 2");
            sa.AddNode("Step 3");
            sa.AddNode("Step 4");
            // Insert a review step at position 2
            sa.InsertNodeAt(2, "Review");
            // Remove the last one
            sa.RemoveNodeAt(sa.NodeCount - 1);
            // Set final texts
            sa.ReplaceTexts("Plan", "Design", "Review", "Build");
            // Emphasize first
            sa.SetNodeText(0, "Plan", bold: true, italic: false);

            document.Save(openWord);
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }

        // Flexible algorithmic layout: Cycle with add/insert/remove and text edits
        internal static void Example_FlexibleCycleSmartArt_FullFlow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Flexible Cycle SmartArt: add/insert/remove + text edits");
            string filePath = Path.Combine(folderPath, "SmartArtCycle_Flexible.docx");
            using WordDocument document = WordDocument.Create(filePath);

            var sa = document.AddSmartArt(SmartArtType.Cycle);
            // Add a few nodes to make a cycle
            sa.AddNode("A");
            sa.AddNode("B");
            sa.AddNode("C");
            sa.AddNode("D");
            // Insert at second position
            sa.InsertNodeAt(1, "AB");
            // Remove a later one
            if (sa.NodeCount > 3) sa.RemoveNodeAt(3);
            // Replace remaining texts in cycle order
            sa.ReplaceTexts("Start", "AB", "C", "D");
            // Add style on one node
            sa.SetNodeText(1, "AB", bold: false, italic: true, underline: false, colorHex: "#008000", sizePt: 12);

            document.Save(openWord);
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }
    }
}
