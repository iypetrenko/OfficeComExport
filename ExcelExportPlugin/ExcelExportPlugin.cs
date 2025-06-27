using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeComExport.PluginBase;

namespace ExcelExportPlugin
{
    public class ExcelExporter : IExportPlugin
    {
        public string Name => "Excel Exporter";
        public string Description => "Экспорт данных в формат Microsoft Excel";
        public string FileExtension => ".xlsx";

        public bool Export(List<DataItem> data, string fileName)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                Console.WriteLine("Плагин: Создание Excel документа...");

                // Создание приложения Excel
                excelApp = new Excel.Application();
                excelApp.Visible = false; // Скрыть Excel при работе плагина

                // Создание книги
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Отчет";

                // Заголовок
                worksheet.Cells[1, 1] = "ОТЧЕТ ИЗ ПЛАГИНА EXCEL";
                Excel.Range titleRange = worksheet.Range["A1:D1"];
                titleRange.Merge();
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                titleRange.Interior.Color = Excel.XlRgbColor.rgbLightBlue;

                // Дата создания
                worksheet.Cells[2, 1] = $"Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm}";
                Excel.Range dateRange = worksheet.Range["A2:D2"];
                dateRange.Merge();
                dateRange.Font.Italic = true;

                // Заголовки таблицы
                worksheet.Cells[4, 1] = "ID";
                worksheet.Cells[4, 2] = "Название";
                worksheet.Cells[4, 3] = "Описание";
                worksheet.Cells[4, 4] = "Дата обработки";

                // Форматирование заголовков
                Excel.Range headerRange = worksheet.Range["A4:D4"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // Заполнение данных
                if (data != null && data.Count > 0)
                {
                    for (int i = 0; i < data.Count; i++)
                    {
                        worksheet.Cells[i + 5, 1] = data[i].Id;
                        worksheet.Cells[i + 5, 2] = data[i].Name ?? "";
                        worksheet.Cells[i + 5, 3] = data[i].Description ?? "";
                        worksheet.Cells[i + 5, 4] = DateTime.Now.ToString("dd.MM.yyyy");
                    }

                    // Обводка таблицы данных
                    Excel.Range dataRange = worksheet.Range[$"A4:D{data.Count + 4}"];
                    dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                // Добавление итоговой строки
                int lastRow = (data?.Count ?? 0) + 5;
                worksheet.Cells[lastRow + 1, 1] = "Итого записей:";
                worksheet.Cells[lastRow + 1, 2] = data?.Count ?? 0;
                Excel.Range totalRange = worksheet.Range[$"A{lastRow + 1}:B{lastRow + 1}"];
                totalRange.Font.Bold = true;

                // Автоподбор ширины колонок
                worksheet.Columns.AutoFit();

                // Добавление диаграммы (если есть данные)
                if (data != null && data.Count > 0)
                {
                    Excel.ChartObjects chartObjs = (Excel.ChartObjects)worksheet.ChartObjects();
                    Excel.ChartObject chartObj = chartObjs.Add(350, 50, 300, 200);
                    Excel.Chart chart = chartObj.Chart;

                    chart.SetSourceData(worksheet.Range[$"A4:B{data.Count + 4}"]);
                    chart.ChartType = Excel.XlChartType.xlColumnClustered;
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = "Распределение данных";
                }

                // Сохранение файла
                workbook.SaveAs(fileName);

                Console.WriteLine($"Плагин: Excel документ сохранен: {fileName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Плагин: Ошибка при экспорте в Excel: {ex.Message}");
                return false;
            }
            finally
            {
                // Очистка COM-объектов
                CleanupExcelResources(worksheet, workbook, excelApp);
            }
        }

        private void CleanupExcelResources(Excel.Worksheet worksheet, Excel.Workbook workbook, Excel.Application excelApp)
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
                Console.WriteLine($"Плагин: Ошибка при очистке Excel ресурсов: {ex.Message}");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}