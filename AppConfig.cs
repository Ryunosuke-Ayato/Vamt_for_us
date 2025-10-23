using Microsoft.Extensions.Configuration;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace Vamt_for_us.Services
{
    #region Основные классы настроек
    // Обобщающий класс конфига
    public class AppConfig
    {
        public ConnectionStrings ConnectionStrings { get; set; } = new();
        public AppSettings AppSettings { get; set; } = new();
    }

    // Класс строки подключения 
    public class ConnectionStrings
    {
        public string DefaultConnection { get; set; } = string.Empty;
    }

    // Класс настроек
    public class AppSettings
    {
        public string Version { get; set; } = "1.0.0";
        public string AppName { get; set; } = string.Empty;
        public string Company { get; set; } = string.Empty;
        public ExportSettings ExportSettings { get; set; } = new();
    }

    // Класс настроек отчётов
    public class ExportSettings
    {
        public string DefaultPath { get; set; } = "Desktop";
        public bool AutoOpenExport { get; set; } = false;
    }
    #endregion

    // Класс сервиса конфигурации
    public class ConfigurationService
    {
        private static IConfiguration _configuration;
        private static AppConfig _appConfig;
        private static string _configFilePath;

        static ConfigurationService()
        {
            try
            {
                _configFilePath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
                LoadConfiguration();
            }
            catch (Exception ex)
            {
                // Сброс до дефолта в случае невозможности загрузить настройки
                _appConfig = CreateDefaultConfig();
            }
        }

        // Загрузка конфига и настроек
        private static void LoadConfiguration()
        {
            var configurationBuilder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            _configuration = configurationBuilder.Build();
            _appConfig = new AppConfig();
            _configuration.Bind(_appConfig);

            // Если файла не существует, создаем его с дефолт значениями
            if (!File.Exists(_configFilePath))
            {
                SaveConfigurationToFile();
            }
        }

        public static string ConnectionString => _appConfig.ConnectionStrings.DefaultConnection;
        public static string AppVersion => _appConfig.AppSettings.Version;
        public static string AppName => _appConfig.AppSettings.AppName;
        public static string Company => _appConfig.AppSettings.Company;
        public static ExportSettings ExportSettings => _appConfig.AppSettings.ExportSettings;

        // Метод для обновления строки подключения
        public static void UpdateConnectionString(string newConnectionString)
        {
            try
            {
                _appConfig.ConnectionStrings.DefaultConnection = newConnectionString;
                SaveConfigurationToFile();
                LoadConfiguration();
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось обновить строку подключения: {ex.Message}", ex);
            }
        }

        // Метод сохраняем конфиг и настройки
        private static void SaveConfigurationToFile()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };

                var json = JsonSerializer.Serialize(_appConfig, options);
                File.WriteAllText(_configFilePath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось сохранить конфигурацию: {ex.Message}", ex);
            }
        }

        // Метод настроек по дефолту
        private static AppConfig CreateDefaultConfig()
        {
            return new AppConfig
            {
                ConnectionStrings = new ConnectionStrings
                {
                    DefaultConnection = ""
                },
                AppSettings = new AppSettings
                {
                    //Version = "1.12.293.251024",
                    AppName = "VAMT for Us",
                    Company = "69 Team",
                    ExportSettings = new ExportSettings
                    {
                        DefaultPath = "Desktop",
                        AutoOpenExport = false
                    }
                }
            };
        }

        // Метод для получения информации о сборке
        public static string GetAssemblyVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                return assembly.GetName().Version?.ToString() ?? AppVersion;
            }
            catch
            {
                return AppVersion;
            }
        }

        // Метод для отображения полной информации о версии
        public static string GetFullVersionInfo()
        {
            return $"{AppName} v{AppVersion}\n" +
                   $"Сборка: {GetAssemblyVersion()}\n" +
                   $"Компания: {Company}\n" +
                   $"Дата сборки: {File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location):dd.MM.yyyy HH:mm}";
        }
    }
}
