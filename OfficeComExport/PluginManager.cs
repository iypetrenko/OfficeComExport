using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeComExport.PluginBase;

namespace OfficeComExport
{
    /// <summary>
    /// Менеджер для динамической загрузки плагинов
    /// </summary>
    public class PluginManager
    {
        private readonly List<IExportPlugin> _plugins = new List<IExportPlugin>();
        private readonly string _pluginsDirectory;

        public PluginManager(string pluginsDirectory = "Plugins")
        {
            _pluginsDirectory = pluginsDirectory;
            LoadPlugins();
        }

        /// <summary>
        /// Получить все загруженные плагины
        /// </summary>
        public IReadOnlyList<IExportPlugin> GetPlugins() => _plugins.AsReadOnly();

        /// <summary>
        /// Найти плагин по расширению файла
        /// </summary>
        public IExportPlugin GetPluginByExtension(string extension)
        {
            return _plugins.FirstOrDefault(p => p.FileExtension.Equals(extension, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Найти плагин по имени
        /// </summary>
        public IExportPlugin GetPluginByName(string name)
        {
            return _plugins.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Динамическая загрузка плагинов из папки
        /// </summary>
        private void LoadPlugins()
        {
            try
            {
                // Создаем папку для плагинов, если её нет
                if (!Directory.Exists(_pluginsDirectory))
                {
                    Directory.CreateDirectory(_pluginsDirectory);
                    Console.WriteLine($"Создана папка для плагинов: {_pluginsDirectory}");
                }

                // Ищем все DLL файлы в папке плагинов
                string[] dllFiles = Directory.GetFiles(_pluginsDirectory, "*.dll");

                Console.WriteLine($"Найдено {dllFiles.Length} потенциальных плагинов");

                foreach (string dllFile in dllFiles)
                {
                    try
                    {
                        LoadPluginFromAssembly(dllFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при загрузке плагина {dllFile}: {ex.Message}");
                    }
                }

                Console.WriteLine($"Успешно загружено {_plugins.Count} плагинов");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке плагинов: {ex.Message}");
            }
        }

        /// <summary>
        /// Загрузка плагина из конкретной сборки
        /// </summary>
        private void LoadPluginFromAssembly(string assemblyPath)
        {
            // Загружаем сборку
            Assembly assembly = Assembly.LoadFrom(assemblyPath);

            // Ищем типы, которые реализуют IExportPlugin
            Type[] types = assembly.GetTypes();

            foreach (Type type in types)
            {
                // Проверяем, реализует ли тип интерфейс IExportPlugin
                if (typeof(IExportPlugin).IsAssignableFrom(type) && !type.IsInterface && !type.IsAbstract)
                {
                    try
                    {
                        // Создаем экземпляр плагина
                        IExportPlugin plugin = (IExportPlugin)Activator.CreateInstance(type);
                        _plugins.Add(plugin);

                        Console.WriteLine($"Загружен плагин: {plugin.Name} ({plugin.Description})");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при создании экземпляра плагина {type.Name}: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Показать информацию о всех загруженных плагинах
        /// </summary>
        public void ShowPluginsInfo()
        {
            Console.WriteLine("\n=== ЗАГРУЖЕННЫЕ ПЛАГИНЫ ===");

            if (_plugins.Count == 0)
            {
                Console.WriteLine("Плагины не найдены");
                return;
            }

            for (int i = 0; i < _plugins.Count; i++)
            {
                var plugin = _plugins[i];
                Console.WriteLine($"{i + 1}. {plugin.Name}");
                Console.WriteLine($"   Описание: {plugin.Description}");
                Console.WriteLine($"   Расширение: {plugin.FileExtension}");
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Выполнить экспорт с помощью выбранного плагина
        /// </summary>
        public bool ExportWithPlugin(int pluginIndex, List<DataItem> data, string fileName)
        {
            if (pluginIndex < 0 || pluginIndex >= _plugins.Count)
            {
                Console.WriteLine("Неверный индекс плагина");
                return false;
            }

            var plugin = _plugins[pluginIndex];

            // Добавляем расширение к имени файла, если его нет
            if (!fileName.EndsWith(plugin.FileExtension))
            {
                fileName += plugin.FileExtension;
            }

            Console.WriteLine($"Экспорт с помощью плагина: {plugin.Name}");
            return plugin.Export(data, fileName);
        }

        /// <summary>
        /// Экспорт с автоматическим выбором плагина по расширению
        /// </summary>
        public bool ExportWithAutoPlugin(List<DataItem> data, string fileName)
        {
            string extension = Path.GetExtension(fileName);
            var plugin = GetPluginByExtension(extension);

            if (plugin == null)
            {
                Console.WriteLine($"Плагин для расширения {extension} не найден");
                return false;
            }

            Console.WriteLine($"Автоматически выбран плагин: {plugin.Name}");
            return plugin.Export(data, fileName);
        }
    }
}