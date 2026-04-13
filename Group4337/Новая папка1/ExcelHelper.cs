using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Group4337.Новая_папка;

namespace Group4337.Новая_папка1
{
    public static class ExcelHelper
    {
        // === ИМПОРТ из 3.xlsx (Вариант 6) ===
        // Ожидаемые колонки: Код клиента | ФИО | E-mail | Улица проживания
        public static List<Client> ImportFromExcel(string filePath)
        {
            var clients = new List<Client>();
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);
                // Пропускаем заголовок (строка 1), читаем данные
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

        // === ЭКСПОРТ с группировкой по улице (Вариант 6) ===
        // Формат листа: Код клиента | ФИО | E-mail
        // ФИО сортируются по алфавиту внутри каждого листа
        public static void ExportToExcel(List<Client> clients, string outputPath)
        {
            // Группировка по улице, сортировка групп по алфавиту
            var grouped = clients.GroupBy(c => c.Street).OrderBy(g => g.Key);

            using (var wb = new XLWorkbook())
            {
                foreach (var group in grouped)
                {
                    // Очищаем имя листа от недопустимых символов
                    string sheetName = new string(group.Key
                        .Where(ch => !Path.GetInvalidFileNameChars().Contains(ch))
                        .ToArray());
                    if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                    if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "NoStreet";

                    var ws = wb.Worksheets.Add(sheetName);

                    // Заголовки согласно Варианту 6
                    ws.Cell(1, 1).Value = "Код клиента";
                    ws.Cell(1, 2).Value = "ФИО";
                    ws.Cell(1, 3).Value = "E-mail";

                    // Сортировка ФИО по алфавиту (требование варианта)
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
