using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvMappingTests
{
    private sealed record Person
    {
        public int Id { get; init; }

        public string Name { get; init; } = string.Empty;

        public int Age { get; init; }

        public string City { get; init; } = string.Empty;
    }

    private sealed record EventRow
    {
        public DateTime Created { get; init; }
    }

    [Fact]
    public void Maps_To_Typed_Record()
    {
        var doc = new CsvDocument()
            .WithHeader("Id", "Name", "Age", "City")
            .AddRow(1, "Przemek", 36, "Mikołów")
            .AddRow(2, "Dominika", 30, "Mikołów");

        var people = doc.Map<Person>(map => map
            .FromColumn<int>("Id", (p, v) => p with { Id = v })
            .FromColumn<string>("Name", (p, v) => p with { Name = v })
            .FromColumn<int>("Age", (p, v) => p with { Age = v })
            .FromColumn<string>("City", (p, v) => p with { City = v })
        ).ToList();

        Assert.Equal(2, people.Count);
        Assert.Equal("Dominika", people[1].Name);
        Assert.Equal(30, people[1].Age);
    }

    [Fact]
    public void Map_Uses_Document_DateTime_Formats()
    {
        var doc = CsvDocument.Parse(
            "Created\n07-Jul-2026\n",
            new CsvLoadOptions { DateTimeFormats = new[] { "dd-MMM-yyyy" } });

        var row = Assert.Single(doc.Map<EventRow>(map => map
            .FromColumn<DateTime>("Created", (item, value) => item with { Created = value })));

        Assert.Equal(new DateTime(2026, 7, 7), row.Created);
    }
}
