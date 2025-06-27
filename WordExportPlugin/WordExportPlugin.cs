using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using OfficeComExport.PluginBase;

namespace WordExportPlugin
{
    public class WordExporter : IExportPlugin
    {
        public string Name => "Word Exporter";
        public string Description => "Экспорт данных в формат Microsoft Word";
        public string FileExtension => ".docx";

        public bool Export(List<DataItem> data, string fileName)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                Console.WriteLine("Плагин: Создание Word документа...");

                // Создание приложения Word
                wordApp = new Word.Application();
                wordApp.Visible = false; // Скрыть Word при работе плагина

                // Создание документа
                doc = wordApp.Documents.Add();

                // Добавление заголовка
                Word.Paragraph titlePara = doc.Paragraphs.Add();
                titlePara.Range.Text = "ОТЧЕТ ИЗ ПЛАГИНА WORD\n";
                titlePara.Range.Font.Size = 16;
                titlePara.Range.Font.Bold = 1;
                titlePara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                titlePara.Range.InsertParagraphAfter();

                // Добавление даты создания
                Word.Paragraph datePara = doc.Paragraphs.Add();
                datePara.Range.Text = $"Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm}\n";
                datePara.Range.InsertParagraphAfter();

                // Создание таблицы
                if (data != null && data.Count > 0)
                {
                    Word.Table table = doc.Tables.Add(doc.Range(), data.Count + 1, 3);
                    table.Borders.Enable = 1;

                    // Заголовки таблицы
                    table.Cell(1, 1).Range.Text = "ID";
                    table.Cell(1, 2).Range.Text = "Название";
                    table.Cell(1, 3).Range.Text = "Описание";

                    // Заполнение данных
                    for (int i = 0; i < data.Count; i++)
                    {
                        table.Cell(i + 2, 1).Range.Text = data[i].Id.ToString();
                        table.Cell(i + 2, 2).Range.Text = data[i].Name ?? "";
                        table.Cell(i + 2, 3).Range.Text = data[i].Description ?? "";
                    }

                    // Форматирование заголовка
                    Word.Row headerRow = table.Rows[1];
                    headerRow.Range.Font.Bold = 1;
                    headerRow.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                }

                // Добавление подписи
                Word.Paragraph footerPara = doc.Paragraphs.Add();
                footerPara.Range.Text = "\nДокумент создан с помощью плагина Word Export";
                footerPara.Range.Font.Italic = 1;
                footerPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                // Сохранение документа
                doc.SaveAs2(fileName);

                Console.WriteLine($"Плагин: Word документ сохранен: {fileName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Плагин: Ошибка при экспорте в Word: {ex.Message}");
                return false;
            }
            finally
            {
                // Очистка COM-объектов
                CleanupWordResources(doc, wordApp);
            }
        }

        private void CleanupWordResources(Word.Document doc, Word.Application wordApp)
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
                Console.WriteLine($"Плагин: Ошибка при очистке Word ресурсов: {ex.Message}");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}