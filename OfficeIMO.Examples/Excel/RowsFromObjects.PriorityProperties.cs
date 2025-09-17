using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates Include/Exclude and Priority/Pinned ordering when generating rows from POCOs.
    /// </summary>
    public static class RowsFromObjectsPriorityProperties {
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

        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - RowsFrom objects with Include/Exclude/Priority");
            string filePath = System.IO.Path.Combine(folderPath, "RowsFromObjects.Priority.xlsx");

            var users = new List<User> {
                new User { Id = "u-001", FirstName = "Alice", LastName = "Doe", Email = "alice@example.com", Age = 30, Address = new Address{ City = "NY", Country = "US" } },
                new User { Id = "u-002", FirstName = "Bob", LastName = "Smith", Email = "bob@example.com", Age = 25, Address = new Address{ City = "LA", Country = "US" } },
                new User { Id = "u-003", FirstName = "Charlie", LastName = "Brown", Email = "charlie@example.com", Age = 40, Address = new Address{ City = "London", Country = "UK" } },
            };

            using var document = ExcelDocument.Create(filePath);

            // Sheet 1: Default order (discovery order)
            document.AsFluent()
                .Sheet("Default", s => s
                    .RowsFrom(users, o => {
                        o.ExpandProperties.Add(nameof(User.Address));
                        o.HeaderCase = HeaderCase.Title;
                    })
                    .Table("Users", t => t.Style(TableStyle.TableStyleMedium9))
                    .AutoFit(columns: true, rows: false))
                .End();

            // Sheet 2: Include/Exclude by property names
            document.AsFluent()
                .Sheet("IncludeExclude", s => s
                    .RowsFrom(users, o => {
                        o.ExpandProperties.Add(nameof(User.Address));
                        o.HeaderCase = HeaderCase.Title;
                        // Only keep Id, Name, and Address.City; hide Email and Age
                        o.IncludeProperties = new[] { nameof(User.Id), nameof(User.FirstName), nameof(User.LastName), "Address.City" };
                        o.ExcludeProperties = new[] { nameof(User.Email), nameof(User.Age) };
                    })
                    .Table("Users", t => t.Style(TableStyle.TableStyleMedium2))
                    .AutoFit(columns: true, rows: false))
                .End();

            // Sheet 3: Priority ordering and pinning
            document.AsFluent()
                .Sheet("Priority", s => s
                    .RowsFrom(users, o => {
                        o.ExpandProperties.Add(nameof(User.Address));
                        o.HeaderCase = HeaderCase.Title;
                        // One-liner chaining: pin first, set priority order, pin last
                        o.PinFirst(nameof(User.Id))
                         .PriorityOrder(nameof(User.LastName), nameof(User.FirstName), "Address.City")
                         .PinLast(nameof(User.Email));
                    })
                    .Table("Users", t => t.Style(TableStyle.TableStyleMedium9))
                    .AutoFit(columns: true, rows: false))
                .End();

            // Sheet 4: Only pin last
            document.AsFluent()
                .Sheet("PinnedLastOnly", s => s
                    .RowsFrom(users, o => {
                        o.ExpandProperties.Add(nameof(User.Address));
                        o.HeaderCase = HeaderCase.Title;
                        o.PinLast(nameof(User.Email));
                    })
                    .Table("Users", t => t.Style(TableStyle.TableStyleMedium4))
                    .AutoFit(columns: true, rows: false))
                .End();

            // Sheet 5: SheetComposer.TableFrom with the same options
            var composer = new SheetComposer(document, "ComposerPriority");
            composer.Section("Users (Composer)");
            composer.TableFrom(users, title: null, configure: o => {
                o.PinFirst(nameof(User.Id))
                 .PriorityOrder(nameof(User.LastName), nameof(User.FirstName), "Address.City")
                 .PinLast(nameof(User.Email));
                o.HeaderCase = HeaderCase.Title;
                o.ExpandProperties.Add(nameof(User.Address));
            }, style: TableStyle.TableStyleMedium9, visuals: v => { });
            composer.Finish(autoFitColumns: true);

            // Sheet 6: Single-call Order(pinFirst, priority, pinLast)
            document.AsFluent()
                .Sheet("OrderOneCall", s => s
                    .RowsFrom(users, o => {
                        o.HeaderCase = HeaderCase.Title;
                        o.ExpandProperties.Add(nameof(User.Address));
                        o.Order(
                            pinFirst: new[] { nameof(User.Id) },
                            priority: new[] { nameof(User.LastName), nameof(User.FirstName), "Address.City" },
                            pinLast: new[] { nameof(User.Email) }
                        );
                    })
                    .Table("Users", t => t.Style(TableStyle.TableStyleMedium6))
                    .AutoFit(columns: true, rows: false))
                .End();

            document.Save(openExcel);
        }
    }
}
