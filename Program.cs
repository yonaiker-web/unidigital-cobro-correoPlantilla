using System.Globalization;
using ClosedXML.Excel;
using DotLiquid;
using Microsoft.Extensions.FileSystemGlobbing;
using Microsoft.Extensions.FileSystemGlobbing.Abstractions;
using Unidigital.Cobros;
using Unidigital.Cobros.Interfaces;

System.IO.Directory.CreateDirectory("./financiados");
System.IO.Directory.CreateDirectory("./financiados/pending");
System.IO.Directory.CreateDirectory("./financiados/backup");
System.IO.Directory.CreateDirectory("./financiados/processed");
System.IO.Directory.CreateDirectory("./financiados/errored");

System.IO.Directory.CreateDirectory("./centralized");
System.IO.Directory.CreateDirectory("./centralized/pending");
System.IO.Directory.CreateDirectory("./centralized/backup");
System.IO.Directory.CreateDirectory("./centralized/processed");
System.IO.Directory.CreateDirectory("./centralized/errored");

#region Financiados

Matcher financedMatcher = new();
financedMatcher.AddIncludePatterns(new[] { "*.xlsx" });

string financedDirectory = "./financiados/pending";

PatternMatchingResult financedFiles = financedMatcher.Execute(
    new DirectoryInfoWrapper(
        new DirectoryInfo(financedDirectory)));

var DEBT_USD = 225;

foreach (var file in financedFiles.Files)
{
    System.IO.Directory.CreateDirectory($"./emails/{file.Path}");

    var workbook = new XLWorkbook(Path.Combine("./financiados/pending", file.Path));
    var ws = workbook.Worksheet(1);
    var rows = ws.RangeUsed().RowsUsed();
    var groups = rows.Skip(1).GroupBy(row => row.Cell("B").GetFormattedString());

    var clients = groups.Select(group =>
    {
        var firstRow = group.First();

        var charges = group.Select(row => new PosCharge
        {
            ValueVES = ((decimal)row.Cell("E").GetDouble()),
            ExchangeRate = ((decimal)row.Cell("F").GetDouble()),
            ValueUSD = ((decimal)row.Cell("G").GetDouble()),
            Description = row.Cell("H").GetFormattedString(),
            Date = row.Cell("I").GetFormattedString()
        }).ToList();

        var chargedVES = charges.Select(x => x.ValueVES).Sum();
        var chargedUSD = charges.Select(x => x.ValueUSD).Sum();
        var debt = DEBT_USD - chargedUSD;

        var title = debt > 0 ? "NOTIFICACIÓN DE COBRO POS FINANCIADOS" : "ESTADO DE CUENTA POS FINANCIADOS";

        return new PosClient
        {
            Id = firstRow.Cell("B").GetFormattedString(),
            AfiliationCode = firstRow.Cell("C").GetFormattedString(),
            Terminal = firstRow.Cell("D").GetFormattedString(),
            Serial = firstRow.Cell("J").GetFormattedString(),
            Model = firstRow.Cell("K").GetFormattedString(),
            Name = firstRow.Cell("L").GetFormattedString(),
            Phone = firstRow.Cell("Q").GetFormattedString(),
            RIF = firstRow.Cell("M").GetFormattedString(),
            State = firstRow.Cell("P").GetFormattedString(),
            Email = firstRow.Cell("R").GetFormattedString(),
            Bank = firstRow.Cell("S").GetFormattedString(),
            Charges = charges,
            ChargedUSD = Util.formatNumber(chargedUSD),
            ChargedVES = Util.formatNumber(chargedVES),
            Debt = Util.formatNumber(debt),
            Title = title
        };
    }).ToList();

    foreach (var client in clients)
    {
        var template = Template.Parse(File.ReadAllText("./template.liquid"));

        File.WriteAllText($"./emails/{file.Path}/{client.Id}.html", template.Render(Hash.FromAnonymousObject(client)));
    }
}
#endregion

#region Centralizados

Matcher centralizedMatcher = new();
centralizedMatcher.AddIncludePatterns(new[] { "*.xlsx" });

string centralizedDirectory = "./centralizados/pending";

PatternMatchingResult centralizedFiles = financedMatcher.Execute(
    new DirectoryInfoWrapper(
        new DirectoryInfo(centralizedDirectory)));

foreach (var file in centralizedFiles.Files)
{

    System.IO.Directory.CreateDirectory($"./emails/{file.Path}");

    var workbook = new XLWorkbook(Path.Combine("./centralizados/pending", file.Path));
    var ws = workbook.Worksheet(1);
    var rows = ws.RangeUsed().RowsUsed();
    var groups = rows.Skip(1).GroupBy(row => row.Cell("B").GetFormattedString());

    var clients = groups.Select(group =>
    {
        var firstRow = group.First();

        var charges = group.Select(row => new CentralizedCharge
        {
            ValueVES = ((decimal)row.Cell("E").GetDouble()),
            ExchangeRate = ((decimal)row.Cell("F").GetDouble()),
            ValueUSD = ((decimal)row.Cell("G").GetDouble()),
            Description = row.Cell("H").GetFormattedString(),
            Date = row.Cell("I").GetFormattedString()
        }).ToList();

        var chargedVES = charges.Select(x => x.ValueVES).Sum();
        var chargedUSD = charges.Select(x => x.ValueUSD).Sum();
        var debt = DEBT_USD - chargedUSD;

        var title = debt > 0 ? "NOTIFICACIÓN DE COBRO POS FINANCIADOS" : "ESTADO DE CUENTA POS FINANCIADOS";

        return new CentralizedUser
        {
            Id = firstRow.Cell("B").GetFormattedString(),
            AffiliationCode = firstRow.Cell("C").GetFormattedString(),
            Terminal = firstRow.Cell("D").GetFormattedString(),
            Model = firstRow.Cell("K").GetFormattedString(),
            Name = firstRow.Cell("L").GetFormattedString(),
            RIF = firstRow.Cell("M").GetFormattedString(),
            State = firstRow.Cell("P").GetFormattedString(),
            Bank = firstRow.Cell("S").GetFormattedString(),
            Charges = charges,
            Debt = Util.formatNumber(debt),
        };
    }).ToList();

    foreach (var client in clients)
    {
        var template = Template.Parse(File.ReadAllText("./template.liquid"));

        File.WriteAllText($"./emails/{file.Path}/{client.Id}.html", template.Render(Hash.FromAnonymousObject(client)));
    }
}
#endregion
