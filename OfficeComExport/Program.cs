using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using OfficeComExport.PluginBase;

namespace OfficeComExport
{
    class Program
    {
        private static PluginManager _pluginManager;

        static void Main(string[] args)
        {
            Console.WriteLine("=== COM Technology Demo with Plugins ===");

            // Инициализация менеджера плагинов
            _pluginManager = new PluginManager();

            while (true)
            {
                ShowMainMenu();
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ExportToWordDirect();
                        break;
                    case "2":
                        ExportToExcelDirect();
                        break;
                    case "3":
                        ImportFromExcel();
                        break;
                    case "4":
                        ExportWithPlugins();
                        break;
                    case "5":
                        ShowPluginsInfo();
                        break;
                    case "0":
                        return;
                    default:
                        Console.WriteLine("Неверный выбор!");
                        break;
                }
            }
        }

        static void ShowMainMenu()
        {
            Console.WriteLine("\n=== ГЛАВНОЕ МЕНЮ ===");
            Console.WriteLine("1. Прямой экспорт в Word (COM)");
            Console.WriteLine("2. Прямой экспорт в Excel (COM)");
            Console.WriteLine("3. Импорт из Excel");
            Console.WriteLine("4. Экспорт через плагины");
            Console.WriteLine("5. Информация о плагинах");
            Console.WriteLine("0. Выход");
            Console.Write("\nВыберите опцию: ");
        }

        #region Прямая работа с COM (базовое задание)

        static void ExportToWordDirect()
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                Console.WriteLine("Создание Word документа через COM...");

                wordApp = new Word.Application();
                wordApp.Visible = false;

                doc = wordApp.Documents.Add();

                // Заголовок
                Word.Paragraph titlePara = doc.Paragraphs.Add();
                titlePara.Range.Text = "ОТЧЕТ ПО ПРЕДМЕТНОЙ ОБЛАСТИ\n";
                titlePara.Range.Font.Size = 16;
                titlePara.Range.Font.Bold = 1;
                titlePara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                titlePara.Range.InsertParagraphAfter();

                // Дата создания
                Word.Paragraph datePara = doc.Paragraphs.Add();
                datePara.Range.Text = $"Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm}\n";
                datePara.Range.InsertParagraphAfter();

                // Создание таблицы
                CreateWordTable(doc);

                // Подпись
                Word.Paragraph footerPara = doc.Paragraphs.Add();
                footerPara.Range.Text = "\nДокумент создан с использованием COM-технологии";
                footerPara.Range.Font.Italic = 1;

                // Сохранение
                string fileName = GetSaveFileName("Word", ".docx");
                doc.SaveAs2(fileName);

                Console.WriteLine($"Word документ сохранен: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при работе с Word: {ex.Message}");
            }
            finally
            {
                CleanupWordResources(doc, wordApp);
            }
        }

        static void CreateWordTable(Word.Document doc)
        {
            var data = GetSampleData();
            Word.Table table = doc.Tables.Add(doc.Range(), data.Count + 1, 3);
            table.Borders.Enable = 1;

            // Заголовки
            table.Cell(1, 1).Range.Text = "ID";
            table.Cell(1, 2).Range.Text = "Название";
            table.Cell(1, 3).Range.Text = "Описание";

            // Данные
            for (int i = 0; i < data.Count; i++)
            {
                table.Cell(i + 2, 1).Range.Text = data[i].Id.ToString();
                table.Cell(i + 2, 2).Range.Text = data[i].Name;
                table.Cell(i + 2, 3).Range.Text = data[i].Description;
            }

            // Форматирование заголовка
            Word.Row headerRow = table.Rows[1];
            headerRow.Range.Font.Bold = 1;
            headerRow.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
        }

        static void ExportToExcelDirect()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                Console.WriteLine("Создание Excel документа через COM...");

                excelApp = new Excel.Application();
                excelApp.Visible = false;

                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Заголовок
                worksheet.Cells[1, 1] = "ОТЧЕТ ПО ПРЕДМЕТНОЙ ОБЛАСТИ";
                Excel.Range titleRange = worksheet.Range["A1:C1"];
                titleRange.Merge();
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Дата
                worksheet.Cells[2, 1] = $"Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm}";

                // Заголовки таблицы
                worksheet.Cells[4, 1] = "ID";
                worksheet.Cells[4, 2] = "Название";
                worksheet.Cells[4, 3] = "Описание";

                Excel.Range headerRange = worksheet.Range["A4:C4"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;

                // Данные
                var data = GetSampleData();
                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 5, 1] = data[i].Id;
                    worksheet.Cells[i + 5, 2] = data[i].Name;
                    worksheet.Cells[i + 5, 3] = data[i].Description;
                }

                worksheet.Columns.AutoFit();

                string fileName = GetSaveFileName("Excel", ".xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"Excel документ сохранен: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при работе с Excel: {ex.Message}");
            }
            finally
            {
                CleanupExcelResources(worksheet, workbook, excelApp);
            }
        }

        static void ImportFromExcel()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                Console.Write("Введите путь к Excel файлу: ");
                string filePath = Console.ReadLine();

                if (!File.Exists(filePath))
                {
                    Console.WriteLine("Файл не найден!");
                    return;
                }

                Console.WriteLine("Импорт данных из Excel...");

                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                Excel.Range usedRange = worksheet.UsedRange;

                Console.WriteLine("Импортированные данные:");
                for (int row = 1; row <= usedRange.Rows.Count; row++)
                {
                    string rowData = "";
                    for (int col = 1; col <= usedRange.Columns.Count; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                        rowData += cellValue + "\t";
                    }
                    Console.WriteLine(rowData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при импорте: {ex.Message}");
            }
            finally
            {
                CleanupExcelResources(worksheet, workbook, excelApp);
            }
        }

        #endregion

        #region Работа с плагинами (дополнительное задание)

        static void ExportWithPlugins()
        {
            var plugins = _pluginManager.GetPlugins();

            if (plugins.Count == 0)
            {
                Console.WriteLine("Плагины не найдены! Убедитесь, что DLL файлы плагинов находятся в папке 'Plugins'");
                return;
            }

            Console.WriteLine("\n=== ЭКСПОРТ ЧЕРЕЗ ПЛАГИНЫ ===");
            Console.WriteLine("Доступные плагины:");

            for (int i = 0; i < plugins.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {plugins[i].Name} ({plugins[i].FileExtension})");
            }

            Console.Write("Выберите плагин (номер): ");
            if (int.TryParse(Console.ReadLine(), out int choice) && choice > 0 && choice <= plugins.Count)
            {
                var selectedPlugin = plugins[choice - 1];

                Console.Write($"Введите имя файла (без расширения): ");
                string fileName = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(fileName))
                {
                    fileName = $"Report_{DateTime.Now:yyyyMMdd_HHmmss}";
                }

                string fullFileName = GetSaveFileName(fileName, selectedPlugin.FileExtension);
                var data = GetSampleData();

                bool success = _pluginManager.ExportWithPlugin(choice - 1, data, fullFileName);

                if (success)
                {
                    Console.WriteLine("Экспорт успешно выполнен!");
                }
                else
                {
                    Console.WriteLine("Ошибка при экспорте!");
                }
            }
            else
            {
                Console.WriteLine("Неверный выбор плагина!");
            }
        }

        static void ShowPluginsInfo()
        {
            _pluginManager.ShowPluginsInfo();
        }

        #endregion

        #region Вспомогательные методы

        static List<DataItem> GetSampleData()
        {
            return new List<DataItem>
            {
                new DataItem { Id = 1, Name = "Товар 1", Description = "Описание товара 1" },
                new DataItem { Id = 2, Name = "Товар 2", Description = "Описание товара 2" },
                new DataItem { Id = 3, Name = "Товар 3", Description = "Описание товара 3" },
                new DataItem { Id = 4, Name = "Товар 4", Description = "Описание товара 4" },
                new DataItem { Id = 5, Name = "Товар 5", Description = "Описание товара 5" }
            };
        }

        static string GetSaveFileName(string baseName, string extension)
        {
            string directory = @"C:\temp";

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!extension.StartsWith("."))
            {
                extension = "." + extension;
            }

            return Path.Combine(directory, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
        }

        static void CleanupWordResources(Word.Document doc, Word.Application wordApp)
        {
            try
            {
                if (doc != null)
                {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при очистке Word ресурсов: {ex.Message}");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Console.WriteLine("Word ресурсы очищены");
            }
        }

        static void CleanupExcelResources(Excel.Worksheet worksheet, Excel.Workbook workbook, Excel.Application excelApp)
        {
            try
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }

                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при очистке Excel ресурсов: {ex.Message}");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Console.WriteLine("Excel ресурсы очищены");
            }
        }

        #endregion
    }

    public class DataItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
    }
}