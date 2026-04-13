using Group4337.Новая_папка1;
using Microsoft.Win32;
using System.Windows;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DatabaseHelper.Initialize();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Выберите файл 3.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    DatabaseHelper.ClearTable();
                    var clients = ExcelHelper.ImportFromExcel(dlg.FileName);
                    DatabaseHelper.SaveClients(clients);

                    MessageBox.Show($"✅ Успешно импортировано {clients.Count} записей из Excel!",
                        "Импорт завершён",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"❌ Ошибка импорта: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "Export_Variant6_Excel.xlsx",
                Title = "Сохранить результат экспорта"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var clients = DatabaseHelper.GetAllClients();

                    if (clients.Count == 0)
                    {
                        MessageBox.Show("⚠️ База данных пуста. Сначала выполните импорт.",
                            "Внимание",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    ExcelHelper.ExportToExcel(clients, dlg.FileName);

                    MessageBox.Show($"✅ Данные экспортированы в Excel: {dlg.FileName}",
                        "Экспорт завершён",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"❌ Ошибка экспорта: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        private void BtnImportJson_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "JSON Files|*.json",
                Title = "Выберите файл 3.json"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    DatabaseHelper.ClearTable();

                    var clients = JsonHelper.ImportFromJson(dlg.FileName);

                    DatabaseHelper.SaveClients(clients);

                    MessageBox.Show($"✅ Успешно импортировано {clients.Count} записей из JSON!",
                        "Импорт завершён",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"❌ Ошибка импорта: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Word Files|*.docx",
                FileName = "Export_Variant6_Word.docx",
                Title = "Сохранить результат экспорта в Word"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var clients = DatabaseHelper.GetAllClients();

                    if (clients.Count == 0)
                    {
                        MessageBox.Show("⚠️ База данных пуста. Сначала выполните импорт.",
                            "Внимание",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    JsonHelper.ExportToWord(clients, dlg.FileName);

                    MessageBox.Show($"✅ Данные экспортированы в Word: {dlg.FileName}\n\n" +
                        $"• Группировка: по улице проживания\n" +
                        $"• Сортировка: ФИО по алфавиту\n" +
                        $"• Каждая улица на отдельной странице",
                        "Экспорт завершён",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"❌ Ошибка экспорта: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }
    }
}