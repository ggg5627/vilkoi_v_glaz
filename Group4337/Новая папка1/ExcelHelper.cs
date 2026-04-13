using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Group4337.Новая_папка;

namespace Group4337.Новая_папка1
{
    public static class ExcelHelper
    {

        public static List<Client> ImportFromExcel(string filePath)
        {
            var clients = new List<Client>();
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                {
                    clients.Add(new Client
                    {
                        ClientCode = row.Cell(1).GetValue<string>(),
                        FullName = row.Cell(2).GetValue<string>(),
                        Email = row.Cell(3).GetValue<string>(),
                        Street = row.Cell(4).GetValue<string>()
                    });
                }
            }
            return clients;
        }

        public static void ExportToExcel(List<Client> clients, string outputPath)
        {

            var grouped = clients.GroupBy(c => c.Street).OrderBy(g => g.Key);

            using (var wb = new XLWorkbook())
            {
                foreach (var group in grouped)
                {

                    string sheetName = new string(group.Key
                        .Where(ch => !Path.GetInvalidFileNameChars().Contains(ch))
                        .ToArray());
                    if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                    if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "NoStreet";

                    var ws = wb.Worksheets.Add(sheetName);


                    ws.Cell(1, 1).Value = "Код клиента";
                    ws.Cell(1, 2).Value = "ФИО";
                    ws.Cell(1, 3).Value = "E-mail";

                    var sorted = group.OrderBy(c => c.FullName).ToList();

                    int row = 2;
                    foreach (var c in sorted)
                    {
                        ws.Cell(row, 1).Value = c.ClientCode;
                        ws.Cell(row, 2).Value = c.FullName;
                        ws.Cell(row, 3).Value = c.Email;
                        row++;
                    }
                }
                wb.SaveAs(outputPath);
            }
        }
    }
}
