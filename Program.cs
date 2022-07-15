using System.Security.AccessControl;
using System.Globalization;
using ClosedXML.Excel;
using DotLiquid;
using Microsoft.Extensions.FileSystemGlobbing;
using Microsoft.Extensions.FileSystemGlobbing.Abstractions;
using Unidigital.Cobros;
using Unidigital.Cobros.Interfaces;
using Microsoft.Extensions.Configuration;
using System.Text.Json;

System.IO.Directory.CreateDirectory("./financed");
System.IO.Directory.CreateDirectory("./centralized");
System.IO.Directory.CreateDirectory("./errors");
System.IO.Directory.CreateDirectory("./sent");

IConfiguration Configuration = new ConfigurationBuilder()
.SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json")
    .Build();

//email
var emailConfig = Configuration.GetSection("EmailConfiguration").Get<EmailConfiguration>();

var emailSender = new EmailSender(emailConfig);

#region Financiados

Matcher financedMatcher = new();
financedMatcher.AddIncludePatterns(new[] { "*.xlsx" });

string financedDirectory = "./financed";

PatternMatchingResult financedFiles = financedMatcher.Execute(
    new DirectoryInfoWrapper(
        new DirectoryInfo(financedDirectory)));

var DEBT_USD = 225;

foreach (var file in financedFiles.Files)
{
    System.IO.Directory.CreateDirectory($"./emails/financed/{file.Path}");

    var workbook = new XLWorkbook(Path.Combine("./financed", file.Path));
    var ws = workbook.Worksheet(1);
    var rows = ws.RangeUsed().RowsUsed();
    var groups = rows.Skip(1).GroupBy(row => row.Cell("B").GetFormattedString());

    var clients = groups.Select(group =>
    {
        var firstRow = group.First();

        var charges = group.Select(row => new PosCharge
        {
            ValueVES = row.Cell('E').IsEmpty() ? 0 : ((decimal)row.Cell("E").GetDouble()),
            ExchangeRate = row.Cell('F').IsEmpty() ? 0 : ((decimal)row.Cell("F").GetDouble()),
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
            Debt2 = debt,
            Title = title
        };
    }).ToList();

    var template = Template.Parse(File.ReadAllText("./financed.liquid"));

    foreach (var client in clients)
    {
        if (File.Exists($"./sent/{client.Id}.txt"))
        {
            continue;
        }

        try
        {
            var message = new Message(new string[] { "silvinoje14@gmail.com" }, "ESTADO DE CUENTA POS FINANCIADOS", template.Render(Hash.FromAnonymousObject(client)));
            emailSender.SendEmail(message);

            File.WriteAllText($"./sent/{client.Id}.txt", message.Content);
        }
        catch (Exception e)
        {
            File.WriteAllText($"./errors/{client.Id}.txt", e.Message);
        }
    }
}
#endregion Financiados

// #region Centralizados

// Matcher centralizedMatcher = new();
// centralizedMatcher.AddIncludePatterns(new[] { "*.xlsx" });

// string centralizedDirectory = "./centralized";

// string[] stringMonths = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

// PatternMatchingResult centralizedFiles = financedMatcher.Execute(
//     new DirectoryInfoWrapper(
//         new DirectoryInfo(centralizedDirectory)));

// foreach (var file in centralizedFiles.Files)
// {
//     System.IO.Directory.CreateDirectory($"./emails/centralized/{file.Path}");

//     var workbook = new XLWorkbook(Path.Combine("./centralized", file.Path));
//     var ws = workbook.Worksheet(1);
//     var rows = ws.RangeUsed().RowsUsed();
//     var groups = rows.Skip(1).GroupBy(row => row.Cell("E").GetFormattedString());

//     var clients = groups.Select(group =>
//     {
//         var firstRow = group.First();

//         decimal chargedUSD = 0;
//         decimal chargedVES = 0;

//         var charges = group.Select(row =>
//         {
//             var valueVES = ((decimal)row.Cell("S").GetDouble());
//             var valueUSD = ((decimal)row.Cell("T").GetDouble());
//             chargedUSD += valueUSD;
//             chargedVES += valueVES;

//             return new CentralizedCharge
//             {
//                 Day = row.Cell("H").IsEmpty() ? 1 : ((int)row.Cell("H").GetDouble()),
//                 Month = row.Cell("I").IsEmpty() ? 1 : ((int)row.Cell("I").GetDouble()),
//                 Year = row.Cell("K").IsEmpty() ? 1 : ((int)row.Cell("K").GetDouble()),
//                 ValueVES = valueVES,
//                 ExchangeRate = ((decimal)row.Cell("N").GetDouble()),
//                 ValueUSD = valueUSD,
//                 Description = row.Cell("AH").GetFormattedString(),
//                 Date = row.Cell("AI").GetFormattedString()
//             };
//         }).ToList();

//         var months = charges.GroupBy(c => c.Month).Select(month =>
//         {
//             var charged = month.Sum(m => m.ValueUSD);

//             return new CentralizedMonth
//             {
//                 Month = month.First().Month,
//                 MonthLetters = stringMonths[month.First().Month - 1],
//                 MonthlyFee = 30,
//                 Charged = charged,
//                 Pending = 30 - charged
//             };
//         }).ToList();

//         var fee = months.Sum(x => x.MonthlyFee);
//         var debt = months.Sum(x => x.Pending);

//         var user = new CentralizedUser
//         {
//             Id = firstRow.Cell("E").GetFormattedString(),
//             AffiliationCode = firstRow.Cell("D").GetFormattedString(),
//             Terminal = firstRow.Cell("G").GetFormattedString(),
//             Model = "FALTA POR COLOCAR", // firstRow.Cell("K").GetFormattedString(),
//             Name = "FALTA POR COLOCAR", //firstRow.Cell("L").GetFormattedString(),
//             RIF = "FALTA POR COLOCAR", // firstRow.Cell("M").GetFormattedString(),
//             State = "FALTA POR COLOCAR", //firstRow.Cell("P").GetFormattedString(),
//             Bank = firstRow.Cell("B").GetFormattedString(),
//             Charges = charges,
//             Months = months,
//             Debt = Util.formatNumber(debt),
//             ChargedUSD = Util.formatNumber(chargedUSD),
//             ChargedVES = Util.formatNumber(chargedVES),
//             Fee = Util.formatNumber(fee)
//         };

//         return user;
//     }).ToList();

//     var template = Template.Parse(File.ReadAllText("./centralized.liquid"));

//     var htmls = clients.AsParallel().Select(client =>
//     {
//         var html = template.Render(Hash.FromAnonymousObject(client));
//         return (client.Id, html);
//     }).ToArray();

//     foreach (var (id, html) in htmls)
//     {
//         File.WriteAllText($"./emails/centralized/{file.Path}/{id}.html", html);
//     }
// }
// #endregion

Environment.Exit(0);
