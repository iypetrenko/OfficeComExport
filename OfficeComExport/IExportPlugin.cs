using System.Collections.Generic;

namespace OfficeComExport.PluginBase
{
    /// <summary>
    /// Интерфейс для плагинов экспорта данных
    /// </summary>
    public interface IExportPlugin
    {
        /// <summary>
        /// Название плагина
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Описание плагина
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Расширение файла, которое создает плагин
        /// </summary>
        string FileExtension { get; }

        /// <summary>
        /// Экспорт данных
        /// </summary>
        /// <param name="data">Данные для экспорта</param>
        /// <param name="fileName">Путь к файлу для сохранения</param>
        /// <returns>True если экспорт успешен</returns>
        bool Export(List<DataItem> data, string fileName);
        bool Export(List<OfficeComExport.DataItem> data, string fileName);
    }

    /// <summary>
    /// Класс данных для экспорта
    /// </summary>
    public class DataItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
    }
}