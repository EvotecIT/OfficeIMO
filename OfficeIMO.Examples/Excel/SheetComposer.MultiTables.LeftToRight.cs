using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    internal static class SheetComposerMultiTables
    {
        public static void Example_LeftToRight(string folderPath, bool openExcel)
        {
            string filePath = System.IO.Path.Combine(folderPath, "Excel.ExcelReportBuilder.LeftToRight.Tables.xlsx");
            using var doc = ExcelDocument.Create(filePath);
            // Ensure predictable behavior for mixed operations
            doc.Execution.Mode = ExcelExecutionMode.Sequential;
            var left = new List<object>
            {
                new
                {
                    Name = "Alpha",
                    Value = 1,
                    Note = "short"
                },
                new
                {
                    Name = "Beta",
                    Value = 2,
                    Note = "wrap\nthis"
                },
                new
                {
                    Name = "Gamma",
                    Value = 3,
                    Note = "ok"
                }
            };
            var middle = new List<object>
            {
                new
                {
                    Key = "K1",
                    Count = 12
                },
                new
                {
                    Key = "K2",
                    Count = 3
                },
                new
                {
                    Key = "K3",
                    Count = 8
                }
            };
            var right = new List<object>
            {
                new
                {
                    Title = "Link 1",
                    Url = "https://example.com/1"
                },
                new
                {
                    Title = "Link 2",
                    Url = "https://example.com/2"
                },
            };
            var s = doc.AddWorksheet("Left→Right", ExcelSheetNameValidationMode.Sanitize).Report();
            s.Title("Tables (Left → Right)", "Two rows of 3 tables using fixed grid; per-table sizing only.");
            // Row 1 (fixed grid)
            s.Columns(3, cols =>
            {
                cols[0].Section("A");
                var r1 = cols[0].AddTable(left, title: null, visuals: v =>
                {
                    v.FreezeHeaderRow = false;
                }).Range!;
                s.ApplyColumnSizing(r1, opt =>
                {
                    opt.MediumHeaders.Add("Name");
                    opt.NumericHeaders.Add("Value");
                    opt.LongHeaders.Add("Note");
                    opt.WrapHeaders.Add("Note");
                });
                cols[1].Section("B");
                var r2 = cols[1].AddTable(middle, title: null, visuals: v =>
                {
                    v.FreezeHeaderRow = false;
                }).Range!;
                s.ApplyColumnSizing(r2, opt =>
                {
                    opt.MediumHeaders.Add("Key");
                    opt.NumericHeaders.Add("Count");
                });
                cols[2].Section("C");
                var r3 = cols[2].AddTable(right, title: null, visuals: v =>
                {
                    v.FreezeHeaderRow = false;
                }).Range!;
                s.ApplyColumnSizing(r3, opt =>
                {
                    opt.MediumHeaders.Add("Title");
                    opt.LongHeaders.Add("Url");
                    opt.WrapHeaders.Add("Title");
                });
            }, columnWidth: 3, gutter: 1);
            // Row 2 (adaptive) – avoid overlaps when table widths vary; keep a 1-column gutter
            var small = new List<object>
            {
                new
                {
                    A = "One",
                    B = 100
                },
                new
                {
                    A = "Two",
                    B = 25
                },
                new
                {
                    A = "Three",
                    B = 250
                }
            };
            var medium5 = new List<object>
            {
                new
                {
                    C1 = "x",
                    C2 = 1,
                    C3 = 2,
                    C4 = 3,
                    C5 = "ok"
                },
                new
                {
                    C1 = "y",
                    C2 = 4,
                    C3 = 5,
                    C4 = 6,
                    C5 = "warn"
                }
            };
            var wide8 = new List<object>
            {
                new
                {
                    D1 = 1,
                    D2 = 2,
                    D3 = 3,
                    D4 = 4,
                    D5 = 5,
                    D6 = 6,
                    D7 = 7,
                    D8 = 8
                },
                new
                {
                    D1 = 2,
                    D2 = 3,
                    D3 = 4,
                    D4 = 5,
                    D5 = 6,
                    D6 = 7,
                    D7 = 8,
                    D8 = 9
                }
            };
            s.ColumnsAdaptive(new List<Action<ExcelReportBuilder.ColumnComposer>> { c =>
            {
                c.Section("D: Small");
                var rr = c.AddTable(small, null, visuals: v =>
                {
                    v.FreezeHeaderRow = false;
                    v.NumericColumnDecimals["B"] = 0;
                    v.DataBars["B"] = OfficeIMO.OfficeColor.ParseHex("5B9BD5");
                }).Range!;
                s.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.Add("A");
                    o.NumericHeaders.Add("B");
                });
            }, c =>
            {
                c.Section("E: Medium");
                var rr = c.AddTable(medium5, null, visuals: v =>
                {
                    v.FreezeHeaderRow = false;
                    v.NumericColumnDecimals["C2"] = 0;
                    v.NumericColumnDecimals["C3"] = 0;
                    v.NumericColumnDecimals["C4"] = 0;
                    v.TextBackgrounds["C5"] = new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                    {
                        {
                            "ok",
                            "#D1E7DD"
                        },
                        {
                            "warn",
                            "#FFF4CE"
                        }
                    };
                }).Range!;
                s.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.UnionWith(new[] { "C1", "C5" });
                    o.NumericHeaders.UnionWith(new[] { "C2", "C3", "C4" });
                });
            }, c =>
            {
                c.Section("F: Wider");
                var rr = c.AddTable(wide8, null, visuals: v =>
                {
                    v.FreezeHeaderRow = false;
                    v.AutoFormatDynamicCollections = true;
                }).Range!;
                s.ApplyColumnSizing(rr, o =>
                {
                    o.NumericHeaders.UnionWith(new[] { "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8" });
                });
            } }, gutter: 1);
            s.Finish(autoFitColumns: false);
            // Adaptive sheet where gutter is always 1 column regardless of table width
            var ad = doc.AddWorksheet("Adaptive L→R", ExcelSheetNameValidationMode.Sanitize).Report();
            ad.Title("Adaptive Columns (1 gutter)", "Widths inferred from rendered tables");
            // Re-declare the shapes used for the adaptive sheet to make them available in this scope
            var small2 = new List<object>
            {
                new
                {
                    A = "One",
                    B = 100
                },
                new
                {
                    A = "Two",
                    B = 25
                },
                new
                {
                    A = "Three",
                    B = 250
                }
            };
            var medium52 = new List<object>
            {
                new
                {
                    C1 = "x",
                    C2 = 1,
                    C3 = 2,
                    C4 = 3,
                    C5 = "ok"
                },
                new
                {
                    C1 = "y",
                    C2 = 4,
                    C3 = 5,
                    C4 = 6,
                    C5 = "warn"
                }
            };
            var wide82 = new List<object>
            {
                new
                {
                    D1 = 1,
                    D2 = 2,
                    D3 = 3,
                    D4 = 4,
                    D5 = 5,
                    D6 = 6,
                    D7 = 7,
                    D8 = 8
                },
                new
                {
                    D1 = 2,
                    D2 = 3,
                    D3 = 4,
                    D4 = 5,
                    D5 = 6,
                    D6 = 7,
                    D7 = 8,
                    D8 = 9
                }
            };
            ad.ColumnsAdaptive(new List<Action<ExcelReportBuilder.ColumnComposer>> { col =>
            {
                col.Section("A");
                var rr = col.AddTable(left, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ad.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.Add("Name");
                    o.NumericHeaders.Add("Value");
                    o.LongHeaders.Add("Note");
                    o.WrapHeaders.Add("Note");
                });
            }, col =>
            {
                col.Section("B (5 cols)");
                var rr = col.AddTable(medium52, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ad.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.UnionWith(new[] { "C1", "C5" });
                    o.NumericHeaders.UnionWith(new[] { "C2", "C3", "C4" });
                });
            }, col =>
            {
                col.Section("C (8 cols)");
                var rr = col.AddTable(wide82, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ad.ApplyColumnSizing(rr, o =>
                {
                    o.NumericHeaders.UnionWith(new[] { "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8" });
                });
            } }, gutter: 1);
            // Second adaptive row with different shapes
            ad.ColumnsAdaptive(new List<Action<ExcelReportBuilder.ColumnComposer>> { col =>
            {
                col.Section("D (2 cols)");
                var rr = col.AddTable(small2, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ad.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.Add("A");
                    o.NumericHeaders.Add("B");
                });
            }, col =>
            {
                col.Section("E (links)");
                var rr = col.AddTable(right, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ad.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.Add("Title");
                    o.LongHeaders.Add("Url");
                    o.WrapHeaders.Add("Title");
                });
            }, col =>
            {
                col.Section("F (3 cols)");
                var rr = col.AddTable(new[] { new { X = "a", Y = "b", Z = "c" }, new { X = "d", Y = "e", Z = "f" } }, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ad.ApplyColumnSizing(rr, o =>
                {
                    o.MediumHeaders.UnionWith(new[] { "X", "Y", "Z" });
                });
            } }, gutter: 1);
            ad.Finish(autoFitColumns: false);
            // Overflow demos (fixed grid, with Shrink and Summarize)
            var ov = doc.AddWorksheet("Overflow Demos", ExcelSheetNameValidationMode.Sanitize).Report();
            ov.Title("Fixed Grid Overflow Modes", "Demonstrates Shrink and Summarize without emitting in-sheet notices.");
            // Band 1: Shrink (keep only 3 columns)
            ov.Columns(2, cols =>
            {
                cols[0].Section("Shrink to 3");
                var big = new List<object>
                {
                    new
                    {
                        A = 1,
                        B = 2,
                        C = 3,
                        D = 4,
                        E = 5
                    },
                    new
                    {
                        A = 6,
                        B = 7,
                        C = 8,
                        D = 9,
                        E = 10
                    }
                };
                var rs = cols[0].AddTable(big, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ov.ApplyColumnSizing(rs, o =>
                {
                    o.NumericHeaders.UnionWith(new[] { "A", "B", "C" });
                });
                cols[1].Section("Original (5 cols)");
                var ro = cols[1].AddTable(big, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ov.ApplyColumnSizing(ro, o =>
                {
                    o.NumericHeaders.UnionWith(new[] { "A", "B", "C", "D", "E" });
                });
            }, columnWidth: 3, gutter: 1, overflow: ExcelReportOverflowMode.Shrink);
            // Band 2: Summarize (2 + More)
            ov.Columns(1, cols =>
            {
                cols[0].Section("Summarize to 2 + More");
                var big = new List<object>
                {
                    new
                    {
                        A = 1,
                        B = 2,
                        C = 3,
                        D = 4,
                        E = 5
                    },
                    new
                    {
                        A = 6,
                        B = 7,
                        C = 8,
                        D = 9,
                        E = 10
                    }
                };
                var rs = cols[0].AddTable(big, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                ov.ApplyColumnSizing(rs, o =>
                {
                    o.NumericHeaders.UnionWith(new[] { "A", "B" });
                    o.LongHeaders.Add("More");
                    o.WrapHeaders.Add("More");
                });
            }, columnWidth: 3, gutter: 1, overflow: ExcelReportOverflowMode.Summarize);
            ov.Finish(autoFitColumns: false);
            // Third sheet: multiple adaptive rows (bands) using ColumnsAdaptiveRows
            var ar = doc.AddWorksheet("Adaptive Rows", ExcelSheetNameValidationMode.Sanitize).Report();
            ar.Title("Adaptive Rows (Bands)", "Multiple left→right bands stacked vertically; 1-column gutter everywhere.");
            // Example data for a score table with icon sets
            var grades = new List<object>
            {
                new
                {
                    Name = "Alice",
                    Score = 92
                },
                new
                {
                    Name = "Bob",
                    Score = 68
                },
                new
                {
                    Name = "Cara",
                    Score = 81
                },
                new
                {
                    Name = "Drew",
                    Score = 55
                }
            };
            // Reuse shapes defined above in this method: left, right, small2, medium52, wide82
            var rows = new List<IReadOnlyList<Action<ExcelReportBuilder.ColumnComposer>>>
            {
                // Row 1: two tables
                new List<Action<ExcelReportBuilder.ColumnComposer>>
                {
                    c =>
                    {
                        c.Section("Domains");
                        var rr = c.AddTable(left, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.MediumHeaders.Add("Name");
                            o.NumericHeaders.Add("Value");
                            o.LongHeaders.Add("Note");
                            o.WrapHeaders.Add("Note");
                        });
                    },
                    c =>
                    {
                        c.Section("Links");
                        var rr = c.AddTable(right, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.MediumHeaders.Add("Title");
                            o.LongHeaders.Add("Url");
                            o.WrapHeaders.Add("Title");
                        });
                    }
                },
                // Row 2: three tables (2,5,8 columns)
                new List<Action<ExcelReportBuilder.ColumnComposer>>
                {
                    c =>
                    {
                        c.Section("Small");
                        var rr = c.AddTable(small2, null, visuals: v =>
                        {
                            v.FreezeHeaderRow = false;
                            v.NumericColumnDecimals["B"] = 0;
                            v.DataBars["B"] = OfficeIMO.OfficeColor.ParseHex("5B9BD5");
                        }).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.MediumHeaders.Add("A");
                            o.NumericHeaders.Add("B");
                        });
                    },
                    c =>
                    {
                        c.Section("Medium5");
                        var rr = c.AddTable(medium52, null, visuals: v =>
                        {
                            v.FreezeHeaderRow = false;
                            v.NumericColumnDecimals["C2"] = 0;
                            v.NumericColumnDecimals["C3"] = 0;
                            v.NumericColumnDecimals["C4"] = 0;
                            v.TextBackgrounds["C5"] = new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                            {
                                {
                                    "ok",
                                    "#D1E7DD"
                                },
                                {
                                    "warn",
                                    "#FFF4CE"
                                }
                            };
                        }).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.MediumHeaders.UnionWith(new[] { "C1", "C5" });
                            o.NumericHeaders.UnionWith(new[] { "C2", "C3", "C4" });
                        });
                    },
                    c =>
                    {
                        c.Section("Wide8");
                        var rr = c.AddTable(wide82, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.NumericHeaders.UnionWith(new[] { "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8" });
                        });
                    }
                },
                // Row 3: score table with icon set + another small table
                new List<Action<ExcelReportBuilder.ColumnComposer>>
                {
                    c =>
                    {
                        c.Section("Scores");
                        var rr = c.AddTable(grades, null, visuals: v =>
                        {
                            v.FreezeHeaderRow = false;
                            v.NumericColumnDecimals["Score"] = 0;
                            v.IconSetColumns.Add("Score");
                        }).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.MediumHeaders.Add("Name");
                            o.NumericHeaders.Add("Score");
                        });
                    },
                    c =>
                    {
                        c.Section("KV");
                        var rr = c.AddTable(new[] { new { K = "Key", V = "Value" }, new { K = "Env", V = "Prod" }, new { K = "Region", V = "EU" } }, null, visuals: v => v.FreezeHeaderRow = false).Range!;
                        ar.ApplyColumnSizing(rr, o =>
                        {
                            o.MediumHeaders.UnionWith(new[] { "K", "V" });
                        });
                    }
                }
            };
            ar.ColumnsAdaptiveRows(rows, gutter: 1);
            ar.Finish(autoFitColumns: false);
            doc.Save();
            if (openExcel)
                doc.OpenInApplication();
        }
    }
}