using Xceed.Document.NET;
using Xceed.Words.NET;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Group4337.Новая_папка;

namespace Group4337.Новая_папка1
{
    public static class JsonHelper
    {
        public static List<Client> ImportFromJson(string filePath)
        {
            string json = File.ReadAllText(filePath);
            var clients = JsonConvert.DeserializeObject<List<Client>>(json);
            return clients ?? new List<Client>();
        }

        public static void ExportToWord(List<Client> clients, string outputPath)
        {
            var grouped = clients.GroupBy(c => c.Street).OrderBy(g => g.Key);

            using (var document = DocX.Create(outputPath))
            {
                bool isFirst = true;
                foreach (var group in grouped)
                {
                    // ✅ Гарантированный разрыв страницы (вместо InsertBreak)
                    if (!isFirst)
                    {
                        document.InsertSection();
                    }
                    isFirst = false;

                    // Заголовок
                    var header = document.InsertParagraph($"Улица: {group.Key}");
                    header.FontSize(16).Bold().SpacingAfter(10);

                    // Сортировка ФИО по алфавиту (требование варианта 6)
                    var sorted = group.OrderBy(c => c.FullName).ToList();

                    // Таблица
                    var table = document.AddTable(sorted.Count + 1, 3);

                    // ✅ Полное имя пространства имён для Alignment
                    table.Alignment = Xceed.Document.NET.Alignment.left;

                    // Заголовки таблицы
                    table.Rows[0].Cells[0].Paragraphs.First().Append("Код клиента").Bold();
                    table.Rows[0].Cells[1].Paragraphs.First().Append("ФИО").Bold();
                    table.Rows[0].Cells[2].Paragraphs.First().Append("E-mail").Bold();

                    // Данные
                    for (int i = 0; i < sorted.Count; i++)
                    {
                        table.Rows[i + 1].Cells[0].Paragraphs.First().Append(sorted[i].ClientCode ?? "");
                        table.Rows[i + 1].Cells[1].Paragraphs.First().Append(sorted[i].FullName ?? "");
                        table.Rows[i + 1].Cells[2].Paragraphs.First().Append(sorted[i].Email ?? "");
                    }

                    document.InsertTable(table);
                    document.InsertParagraph("");
                }
                document.Save();
            }
        }
    }
}