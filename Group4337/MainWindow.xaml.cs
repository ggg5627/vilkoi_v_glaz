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
            // Инициализация БД при запуске
            DatabaseHelper.Initialize();
        }
       

        // 🔹 Импорт из 3.xlsx (Вариант 6)
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
                    // Очищаем таблицу перед новым импортом (опционально)
                    DatabaseHelper.ClearTable();

                    // Читаем Excel
                    var clients = ExcelHelper.ImportFromExcel(dlg.FileName);

                    // Сохраняем в БД
                    DatabaseHelper.SaveClients(clients);

                    MessageBox.Show($"Успешно импортировано {clients.Count} записей!",
                        "Импорт завершён",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        // 🔹 Экспорт в Excel с группировкой по улице (Вариант 6)
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "Export_Variant6_Gayfullina.xlsx",
                Title = "Сохранить результат экспорта"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    // Читаем данные из БД
                    var clients = DatabaseHelper.GetAllClients();

                    if (clients.Count == 0)
                    {
                        MessageBox.Show("База данных пуста. Сначала выполните импорт.",
                            "Внимание",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    // Экспортируем с группировкой
                    ExcelHelper.ExportToExcel(clients, dlg.FileName);

                    MessageBox.Show($"Данные экспортированы в {dlg.FileName}\n" +
                        $"• Группировка: по улице проживания\n" +
                        $"• Сортировка: ФИО по алфавиту",
                        "Экспорт завершён",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка экспорта: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }
    }
}