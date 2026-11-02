using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.Data;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates Include/Exclude and Priority/Pinned ordering when generating rows from POCOs.
    /// </summary>
    public static class RowsFromObjectsPriorityProperties
    {
        private class User {
            public string Id { get; set; } = string.Empty;
            public string FirstName { get; set; } = string.Empty;
            public string LastName { get; set; } = string.Empty;
            public string Email { get; set; } = string.Empty;
            public int Age { get; set; }
            public Address? Address { get; set; }
        }
        private class Address {
            public string City { get; set; } = string.Empty;
            public string Country { get; set; } = string.Empty;
        }

        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - RowsFrom objects with Include/Exclude/Priority");
            string filePath = System.IO.Path.Combine(folderPath, "RowsFromObjects.Priority.xlsx");
            var users = new List<User>
            {
                new User
                {
                    Id = "u-001",
                    FirstName = "Alice",
                    LastName = "Doe",
                    Email = "alice@example.com",
                    Age = 30,
                    Address = new Address
                    {
                        City = "NY",
                        Country = "US"
                    }
                },
                new User
                {
                    Id = "u-002",
                    FirstName = "Bob",
                    LastName = "Smith",
                    Email = "bob@example.com",
                    Age = 25,
                    Address = new Address
                    {
                        City = "LA",
                        Country = "US"
                    }
                },
                new User
                {
                    Id = "u-003",
                    FirstName = "Charlie",
                    LastName = "Brown",
                    Email = "charlie@example.com",
                    Age = 40,
                    Address = new Address
                    {
                        City = "London",
                        Country = "UK"
                    }
                },
            };
            using var document = ExcelDocument.Create(filePath);
            // Sheet 1: Default order (discovery order)
            {
                var sheet1 = document.AddSheet("Default");
                int nextRow2 = 1;
                var projection3 = new OfficeIMO.Data.ObjectFlattenerOptions();
                {
                    (projection3).ExpandProperties.Add(nameof(User.Address));
                    (projection3).HeaderCase = HeaderCase.Title;
                }

                var table4 = sheet1.AddTable(users, new ExcelTableOptions { Name = "Users", Style = ExcelTableStyle.TableStyleMedium9, IncludeAutoFilter = true, Projection = projection3 }, row: nextRow2);
                nextRow2 = table4.AsRange().LastRow + 1;
                sheet1.AutoFitColumns();
            }

            // Sheet 2: Include/Exclude by property names
            {
                var sheet5 = document.AddSheet("IncludeExclude");
                int nextRow6 = 1;
                var projection7 = new OfficeIMO.Data.ObjectFlattenerOptions();
                {
                    (projection7).ExpandProperties.Add(nameof(User.Address));
                    (projection7).HeaderCase = HeaderCase.Title;
                    // Only keep Id, Name, and Address.City; hide Email and Age
                    (projection7).IncludeProperties = new[]
                    {
                        nameof(User.Id),
                        nameof(User.FirstName),
                        nameof(User.LastName),
                        "Address.City"
                    };
                    (projection7).ExcludeProperties = new[]
                    {
                        nameof(User.Email),
                        nameof(User.Age)
                    };
                }

                var table8 = sheet5.AddTable(users, new ExcelTableOptions { Name = "Users", Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection7 }, row: nextRow6);
                nextRow6 = table8.AsRange().LastRow + 1;
                sheet5.AutoFitColumns();
            }

            // Sheet 3: Priority ordering and pinning
            {
                var sheet9 = document.AddSheet("Priority");
                int nextRow10 = 1;
                var projection11 = new OfficeIMO.Data.ObjectFlattenerOptions();
                {
                    (projection11).ExpandProperties.Add(nameof(User.Address));
                    (projection11).HeaderCase = HeaderCase.Title;
                    // One-liner chaining: pin first, set priority order, pin last
                    (projection11).PinFirst(nameof(User.Id)).PriorityOrder(nameof(User.LastName), nameof(User.FirstName), "Address.City").PinLast(nameof(User.Email));
                }

                var table12 = sheet9.AddTable(users, new ExcelTableOptions { Name = "Users", Style = ExcelTableStyle.TableStyleMedium9, IncludeAutoFilter = true, Projection = projection11 }, row: nextRow10);
                nextRow10 = table12.AsRange().LastRow + 1;
                sheet9.AutoFitColumns();
            }

            // Sheet 4: Only pin last
            {
                var sheet13 = document.AddSheet("PinnedLastOnly");
                int nextRow14 = 1;
                var projection15 = new OfficeIMO.Data.ObjectFlattenerOptions();
                {
                    (projection15).ExpandProperties.Add(nameof(User.Address));
                    (projection15).HeaderCase = HeaderCase.Title;
                    (projection15).PinLast(nameof(User.Email));
                }

                var table16 = sheet13.AddTable(users, new ExcelTableOptions { Name = "Users", Style = ExcelTableStyle.TableStyleMedium4, IncludeAutoFilter = true, Projection = projection15 }, row: nextRow14);
                nextRow14 = table16.AsRange().LastRow + 1;
                sheet13.AutoFitColumns();
            }

            // Sheet 5: ExcelReportBuilder.TableFrom with the same options
            var composer = document.AddSheet("ComposerPriority", ExcelSheetNameValidationMode.Sanitize).Report();
            composer.Section("Users (Composer)");
            composer.AddTable(users, title: null, configure: o =>
            {
                o.PinFirst(nameof(User.Id)).PriorityOrder(nameof(User.LastName), nameof(User.FirstName), "Address.City").PinLast(nameof(User.Email));
                o.HeaderCase = HeaderCase.Title;
                o.ExpandProperties.Add(nameof(User.Address));
            }, style: ExcelTableStyle.TableStyleMedium9, visuals: v =>
            {
            });
            composer.Finish(autoFitColumns: true);
            // Sheet 6: Single-call Order(pinFirst, priority, pinLast)
            {
                var sheet17 = document.AddSheet("OrderOneCall");
                int nextRow18 = 1;
                var projection19 = new OfficeIMO.Data.ObjectFlattenerOptions();
                {
                    (projection19).HeaderCase = HeaderCase.Title;
                    (projection19).ExpandProperties.Add(nameof(User.Address));
                    (projection19).Order(pinFirst: new[] { nameof(User.Id) }, priority: new[] { nameof(User.LastName), nameof(User.FirstName), "Address.City" }, pinLast: new[] { nameof(User.Email) });
                }

                var table20 = sheet17.AddTable(users, new ExcelTableOptions { Name = "Users", Style = ExcelTableStyle.TableStyleMedium6, IncludeAutoFilter = true, Projection = projection19 }, row: nextRow18);
                nextRow18 = table20.AsRange().LastRow + 1;
                sheet17.AutoFitColumns();
            }

            document.Save();
            if (openExcel)
                document.OpenInApplication();
        }
    }
}