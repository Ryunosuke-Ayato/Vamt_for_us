using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Vamt_for_us.Data;
using Vamt_for_us.Models;
using Vamt_for_us.Services;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
namespace Vamt_for_us.ViewModels
{
    #region Базовый регион паттерна MVVM
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object parameter) => _execute();

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool> _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) => _canExecute?.Invoke((T)parameter) ?? true;

        public void Execute(object parameter) => _execute((T)parameter);

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
    #endregion

    #region Viewmodel регион
    // ViewModel для ключей продуктов
    public class ProductKeyViewModel : ViewModelBase
    {
        private string _keyValue = string.Empty;
        private string _keyDescription = string.Empty;
        private int _remainingActivations;
        private string _userRemarks = string.Empty;

        public string KeyValue
        {
            get => _keyValue;
            set => SetProperty(ref _keyValue, value);
        }

        public string KeyDescription
        {
            get => _keyDescription;
            set => SetProperty(ref _keyDescription, value);
        }

        public int RemainingActivations
        {
            get => _remainingActivations;
            set => SetProperty(ref _remainingActivations, value);
        }

        public string UserRemarks
        {
            get => _userRemarks;
            set => SetProperty(ref _userRemarks, value);
        }
    }

    // ViewModel для активированных продуктов
    public class ActiveProductViewModel : ViewModelBase
    {
        private string _fullyQualifiedDomainName = string.Empty;
        private string _productKeyId = string.Empty;
        private string _licenseStatusText = string.Empty;
        private string _keyTypeName = string.Empty;

        public string FullyQualifiedDomainName
        {
            get => _fullyQualifiedDomainName;
            set => SetProperty(ref _fullyQualifiedDomainName, value);
        }

        public string ProductKeyId
        {
            get => _productKeyId;
            set => SetProperty(ref _productKeyId, value);
        }

        public string LicenseStatusText
        {
            get => _licenseStatusText;
            set => SetProperty(ref _licenseStatusText, value);
        }

        public string KeyTypeName
        {
            get => _keyTypeName;
            set => SetProperty(ref _keyTypeName, value);
        }
    }

    // Главная ViewModel
    public class MainViewModel : ViewModelBase
    {
        // основные свойсвта главной ViewModel
        private ObservableCollection<ProductKeyViewModel> _productKeys = new();
        private ObservableCollection<ActiveProductViewModel> _activeProducts = new();
        private ProductKeyViewModel _selectedProductKey;
        private ActiveProductViewModel _selectedActiveProduct;
        private string _comment = string.Empty;
        private bool _isLoading;

        // Объявление команд
        public ICommand ReloadDataCommand { get; private set; }
        public ICommand UpdateCommentCommand { get; private set; }
        public ICommand ExportLicenseReportCommand { get; private set; }
        public ICommand ExportActivatedReportCommand { get; private set; }
        public ICommand ShowInfoCommand { get; private set; }

        public MainViewModel()
        {
            _productKeys = new ObservableCollection<ProductKeyViewModel>();
            _activeProducts = new ObservableCollection<ActiveProductViewModel>();
            InitializeCommands();
        }

        // Инициализация команд/функций
        private void InitializeCommands()
        {
            ReloadDataCommand = new RelayCommand(async () => await ReloadDataAsync());
            UpdateCommentCommand = new RelayCommand(async () => await UpdateCommentAsync(), () => SelectedProductKey != null);
            ExportLicenseReportCommand = new RelayCommand(async () => await ExportLicenseReportAsync());
            ExportActivatedReportCommand = new RelayCommand(async () => await ExportActivatedReportAsync());
            ShowInfoCommand = new RelayCommand(ShowVersionInfo);
            GetAllComputersInDomainCommand = new RelayCommand(async () => await GetAllComputersInDomain());
            SendMessageCommand = new RelayCommand(async () => await SendMessageToSelected());
        }

        #region регион доменной модели 
        // Свойства доменных утилит
        private ObservableCollection<ComputerInfo> _domainComputers = new ObservableCollection<ComputerInfo>();
        private string _message;
        private string _statusMessage = "Готов";
        private Brush _statusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7D00FF"));

        public ObservableCollection<ComputerInfo> DomainComputers
        {
            get => _domainComputers;
            set => SetProperty(ref _domainComputers, value);
        }

        public string Message
        {
            get => _message;
            set => SetProperty(ref _message, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public Brush StatusColor
        {
            get => _statusColor;
            set => SetProperty(ref _statusColor, value);
        }

        public int OnlineComputersCount => DomainComputers?.Count(c => c.IsOnline) ?? 0;
        public int OfflineComputersCount => DomainComputers?.Count(c => !c.IsOnline) ?? 0;

        public ICommand GetAllComputersInDomainCommand { get; private set; }
        public ICommand SendMessageCommand { get; private set; }

        // Инициализация поулчнеия списка компьюетров в домене
        private async Task GetAllComputersInDomain()
        {
            try
            {
                StatusMessage = "Получение списка компьютеров...";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFA000FF"));

                await GetAllPCDomain(); // Тут происходит получение списка

                StatusMessage = $"Список обновлен. Онлайн: {OnlineComputersCount}, Офлайн: {OfflineComputersCount}";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF5B9C5D"));
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFEA5E54"));
            }
        }

        // Функция получения доменных компьютеров
        private async Task GetAllPCDomain()
        {
            var computers = new List<ComputerInfo>();

            // Получаем компьютеры из домена
            DirectoryEntry rootDSE = new DirectoryEntry("LDAP://RootDSE");
            string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
            DirectorySearcher searcher = new DirectorySearcher(new DirectoryEntry($"LDAP://{defaultNamingContext}"));
            searcher.Filter = "(objectCategory=computer)";

            SearchResultCollection results = searcher.FindAll();

            // Сначала собираем все имена
            foreach (SearchResult result in results)
            {
                string dn = result.Properties["distinguishedName"][0].ToString();
                int indexOfCN = dn.IndexOf(",DC=");
                if (indexOfCN > 0)
                {
                    string cn = dn.Substring(0, indexOfCN);
                    string computerName = cn.Split('=')[1];
                    computerName = computerName.Split(',')[0];

                    computers.Add(new ComputerInfo { Name = computerName });
                }
            }

            // Параллельно проверяем доступность
            await Task.Run(() =>
            {
                Parallel.ForEach(computers, computer =>
                {
                    computer.IsOnline = CheckComputerOnline(computer.Name);
                    computer.IPAddress = GetComputerIP(computer.Name);
                });
            });

            // Обновляем UI в основном потоке
            Application.Current.Dispatcher.Invoke(() =>
            {
                DomainComputers.Clear();
                foreach (var computer in computers.OrderBy(c => c.Name))
                {
                    DomainComputers.Add(computer);
                }
                OnPropertyChanged(nameof(OnlineComputersCount));
                OnPropertyChanged(nameof(OfflineComputersCount));
            });
        }

        // Проверка компьютера online/offline
        private bool CheckComputerOnline(string computerName)
        {
            try
            {
                using (var ping = new Ping())
                {
                    // Пробуем пинговать по имени
                    var reply = ping.Send(computerName, 2000); // Увеличиваем таймаут до 2 секунд

                    // Если пинг неуспешен, пробуем получить IP и пинговать по нему
                    if (reply?.Status != IPStatus.Success)
                    {
                        var ip = GetComputerIP(computerName);
                        if (ip != "N/A" && ip != "127.0.0.1")
                        {
                            reply = ping.Send(ip, 2000);
                        }
                    }

                    return reply?.Status == IPStatus.Success;
                }
            }
            catch (PingException)
            {
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        // Получение ip адреса компьютера
        private string GetComputerIP(string computerName)
        {
            try
            {
                var hostEntry = Dns.GetHostEntry(computerName);

                // Получаем IPv4 адрес
                var ipv4Address = hostEntry.AddressList
                    .FirstOrDefault(addr => addr.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);

                // Если IPv4 нет, берем первый доступный адрес
                return ipv4Address?.ToString() ??
                       hostEntry.AddressList.FirstOrDefault()?.ToString() ??
                       "N/A";
            }
            catch (Exception ex)
            {
                return "N/A";
            }
        }

        // Отправка сообщений выбранным компьютерам
        private async Task SendMessageToSelected()
        {
            var allSelectedComputers = DomainComputers.Where(c => c.IsSelected).ToList();
            var onlineSelectedComputers = allSelectedComputers.Where(c => c.IsOnline).ToList();
            var offlineSelectedComputers = allSelectedComputers.Where(c => !c.IsOnline).ToList();

            if (!allSelectedComputers.Any())
            {
                StatusMessage = "Нет выбранных компьютеров для отправки сообщения";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFD4A053"));
                return;
            }

            if (!onlineSelectedComputers.Any())
            {
                StatusMessage = "Нет выбранных онлайн-компьютеров для отправки сообщения";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFD4A053"));
                return;
            }

            if (string.IsNullOrWhiteSpace(Message))
            {
                StatusMessage = "Введите сообщение для отправки";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFD4A053"));
                return;
            }

            try
            {
                StatusMessage = $"Отправка сообщения на {onlineSelectedComputers.Count} компьютер(ов)...";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFA000FF"));

                int successCount = 0;
                int totalCount = onlineSelectedComputers.Count;

                foreach (var computer in onlineSelectedComputers)
                {
                    if (await SendMessageToComputer(computer.Name, Message))  // Вызываем функию отправки сообщения одному компьютеру
                    {
                        successCount++;
                    }

                    StatusMessage = $"Отправлено {successCount} из {totalCount}...";
                    await Task.Delay(100);
                }

                string resultMessage = $"Сообщение отправлено на {successCount} из {totalCount} компьютеров";
                if (offlineSelectedComputers.Any())
                {
                    resultMessage += $". {offlineSelectedComputers.Count} выбранных компьютеров офлайн";
                }

                StatusMessage = resultMessage;
                StatusColor = successCount > 0 ? new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF5B9C5D")) : new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFEA5E54")); ;
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка отправки: {ex.Message}";
                StatusColor = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFEA5E54")); ;
            }
        }

        // Отправка сообщения ОДНОМУ компьютеру
        private async Task<bool> SendMessageToComputer(string computerName, string message)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using (var process = new Process())
                    {
                        string fullPath = Path.Combine(Environment.SystemDirectory, "msg.exe");
                        if (!File.Exists(fullPath))
                        {
                            return false;
                        }

                        process.StartInfo.FileName = fullPath;
                        process.StartInfo.Arguments = $"* /SERVER:{computerName} \"{message}\"";
                        process.StartInfo.UseShellExecute = false;
                        process.StartInfo.CreateNoWindow = true;
                        process.StartInfo.RedirectStandardError = true;
                        process.StartInfo.RedirectStandardOutput = true;

                        process.Start();
                        string errorOutput = process.StandardError.ReadToEnd().Trim();
                        process.WaitForExit(5000); // Таймаут 5 секунд

                        return string.IsNullOrEmpty(errorOutput) && process.ExitCode == 0;
                    }
                }
                catch
                {
                    return false;
                }
            });
        }
        bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        #endregion
        #region Регион ProgressBar'а
        private int _loadingProgress;
        private string _progressStatus = "Готов";
        private bool _isLoadingWithProgress;

        public int LoadingProgress
        {
            get => _loadingProgress;
            set => SetProperty(ref _loadingProgress, value);
        }

        public string ProgressStatus
        {
            get => _progressStatus;
            set => SetProperty(ref _progressStatus, value);
        }

        public bool IsLoadingWithProgress
        {
            get => _isLoadingWithProgress;
            set => SetProperty(ref _isLoadingWithProgress, value);
        }
        #endregion
        #region Данных

        // Временная затычка
        private void ShowVersionInfo()
        {
            MessageBox.Show(
                ConfigurationService.GetFullVersionInfo(),
                "Информация о версии",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        #region Основные спсики и объекты БД
        public ObservableCollection<ProductKeyViewModel> ProductKeys
        {
            get => _productKeys;
            set => SetProperty(ref _productKeys, value);
        }

        public ObservableCollection<ActiveProductViewModel> ActiveProducts
        {
            get => _activeProducts;
            set => SetProperty(ref _activeProducts, value);
        }
        public ProductKeyViewModel SelectedProductKey
        {
            get => _selectedProductKey;
            set
            {
                if (SetProperty(ref _selectedProductKey, value))
                {
                    Comment = value?.UserRemarks ?? string.Empty;
                    ((RelayCommand)UpdateCommentCommand).RaiseCanExecuteChanged();
                }
            }
        }
        public ActiveProductViewModel SelectedActiveProduct
        {
            get => _selectedActiveProduct;
            set => SetProperty(ref _selectedActiveProduct, value);
        }
        #endregion
        public string Comment
        {
            get => _comment;
            set => SetProperty(ref _comment, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }


        // Основной метод загрузки данных
        public async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                await LoadProductKeysDataAsync();
                await LoadActiveProductsDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // Получение подробностей о ключе
        private async Task<string> GetKeyDescriptionByPartialId(string partialKeyId, VamtDbContext context)
        {
            try
            {
                var keyDescription = await context.ProductKeys
                    .Where(pk => pk.KeyId.Substring(6, pk.KeyId.Length - 25) == partialKeyId)
                    .Select(pk => pk.KeyDescription)
                    .FirstOrDefaultAsync();

                return keyDescription;
            }
            catch
            {
                return null;
            }
        }

        // Заполнение списка ключей
        public async Task LoadProductKeysDataAsync()
        {
            try
            {
                using var context = CreateDbContext();
                var keys = await context.ProductKeys
                    .Select(pk => new ProductKeyViewModel
                    {
                        KeyValue = pk.KeyValue,
                        KeyDescription = pk.KeyDescription,
                        RemainingActivations = pk.RemainingActivations,
                        UserRemarks = pk.UserRemarks
                    })
                    .ToListAsync();

                ProductKeys = new ObservableCollection<ProductKeyViewModel>(keys);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки ключей: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Заполнение списка продуктов на компьютерах 
        public async Task LoadActiveProductsDataAsync()
        {
            try
            {
                using var context = CreateDbContext();

                var query = from ap in context.ActiveProducts
                            join lst in context.LicenseStatusTexts on ap.LicenseStatus equals lst.LicenseStatus
                            join pktn in context.ProductKeyTypeNames on ap.ProductKeyType equals pktn.KeyType
                            select new ActiveProductViewModel
                            {
                                FullyQualifiedDomainName = ap.FullyQualifiedDomainName,
                                ProductKeyId = ap.ProductKeyId,
                                LicenseStatusText = lst.LicenseStatusTextValue,
                                KeyTypeName = pktn.KeyTypeName
                            };

                var products = await query.ToListAsync();

                // Обработка ключей продуктов
                foreach (var product in products)
                {
                    if (!string.IsNullOrEmpty(product.ProductKeyId) && product.ProductKeyId.Length > 25)
                    {
                        var key = product.ProductKeyId.Substring(6, product.ProductKeyId.Length - 25);
                        var keyDescription = await GetKeyDescriptionByPartialId(key, context);
                        product.ProductKeyId = keyDescription ?? "Error Key";
                    }
                    else
                    {
                        product.ProductKeyId = "unavailable";
                    }
                }

                ActiveProducts = new ObservableCollection<ActiveProductViewModel>(products);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки активированных продуктов: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private VamtDbContext CreateDbContext()
        {
            return new VamtDbContext();
        }

        // Переподгрузка данных
        private async Task ReloadDataAsync()
        {
            await LoadDataAsync();
        }

        // Обновление комментария к ключу
        public async Task UpdateCommentAsync()
        {
            if (SelectedProductKey == null) return;

            try
            {
                using var context = CreateDbContext();

                // Находим запись по KeyValue (предполагая, что он уникален)
                var productKey = await context.ProductKeys
                    .FirstOrDefaultAsync(pk => pk.KeyValue == SelectedProductKey.KeyValue);

                if (productKey != null)
                {
                    productKey.UserRemarks = Comment;
                    await context.SaveChangesAsync();

                    // Обновляем локальные данные
                    SelectedProductKey.UserRemarks = Comment;
                    RefreshProductKeysCollection();

                    MessageBox.Show("Комментарий успешно обновлен", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Запись не найдена", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка обновления комментария: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Обновляем коллекцию ключей
        private void RefreshProductKeysCollection()
        {
            var updatedCollection = new ObservableCollection<ProductKeyViewModel>(ProductKeys);
            ProductKeys = updatedCollection;
        }
        #endregion
        #region Регион отчетности

        // Инициализация генерации отчёта по лицензиям (ключам)
        public async Task ExportLicenseReportAsync()
        {
            await ExportToExcelAsync(
                ProductKeys,
                new[] { "Ключ", "Описание продукта", "Осталось активаций", "Комментарий" },
                "License report",
                pk => new object[] { pk.KeyValue, pk.KeyDescription, pk.RemainingActivations, pk.UserRemarks });
        }

        // Инициализация генерации отчёта по продуктам (активным на компьютерах)
        public async Task ExportActivatedReportAsync()
        {
            await ExportToExcelAsync(
                ActiveProducts,
                new[] { "Доменное имя", "Ключ продукта", "Статус лицензии", "Тип ключа" },
                "Activated product report",
                ap => new object[] { ap.FullyQualifiedDomainName, ap.ProductKeyId, ap.LicenseStatusText, ap.KeyTypeName });
        }
        
        // Инициализация создания файла - выбор папки
        private async Task ExportToExcelAsync<T>(
            IEnumerable<T> data,
            string[] headers,
            string reportName,
            Func<T, object[]> dataSelector)
        {
            try
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    FileName = $"{reportName} {DateTime.Now:dd.MM.yyyy}.xlsx",
                    Filter = "Excel Files (*.xlsx)|*.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    await CreateExcelFileAsync(saveFileDialog.FileName, data, headers, dataSelector);
                    MessageBox.Show($"Отчет '{reportName}' успешно экспортирован", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Генерация файла отчётности
        private async Task CreateExcelFileAsync<T>(
    string filePath,
    IEnumerable<T> data,
    string[] headers,
    Func<T, object[]> dataSelector)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (File.Exists(filePath))
                        File.Delete(filePath);

                    using var spreadsheet = SpreadsheetDocument.Create(
                        filePath, SpreadsheetDocumentType.Workbook);

                    var workbookPart = spreadsheet.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    // Добавляем стили
                    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = CreateStylesheet();
                    stylesPart.Stylesheet.Save();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var worksheet = new Worksheet();
                    worksheet.Append(CreatePreciseAutoWidthColumns(headers, data, dataSelector)); // Добавляем столбцы с подстройкой ширины

                    var sheetData = new SheetData();

                    // Добавляем заголовки
                    var headerRow = new Row();
                    foreach (var header in headers)
                    {
                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(header),
                            StyleIndex = 1
                        };
                        headerRow.Append(cell);
                    }
                    sheetData.Append(headerRow);

                    // Добавляем данные
                    foreach (var item in data)
                    {
                        var rowData = dataSelector(item);
                        var row = new Row();

                        foreach (var cellValue in rowData)
                        {
                            var cell = new Cell
                            {
                                DataType = GetCellDataType(cellValue),
                                CellValue = new CellValue(cellValue?.ToString() ?? string.Empty),
                                StyleIndex = 2
                            };
                            row.Append(cell);
                        }
                        sheetData.Append(row);
                    }

                    // Добавляем итоговую строку
                    var totalRow = new Row();
                    var totalCell1 = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue("Всего записей:"),
                        StyleIndex = 1
                    };
                    var totalCell2 = new Cell
                    {
                        DataType = CellValues.Number,
                        CellValue = new CellValue(data.Count().ToString()),
                        StyleIndex = 1
                    };

                    totalRow.Append(totalCell1);
                    totalRow.Append(totalCell2);
                    sheetData.Append(totalRow);

                    // Добавляем SheetData в Worksheet
                    worksheet.Append(sheetData);

                    worksheetPart.Worksheet = worksheet;

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Report"
                    };
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Не удалось создать Excel файл: {ex.Message}", ex);
                }
            });
        }

        // Метод автоподбора ширины столбца в зависимости от наполнения
        private Columns CreatePreciseAutoWidthColumns<T>(string[] headers, IEnumerable<T> data, Func<T, object[]> dataSelector)
        {
            var columns = new Columns();

            for (int colIndex = 0; colIndex < headers.Length; colIndex++)
            {
                // Находим самую длинную строку в столбце
                int maxLength = headers[colIndex].Length;

                foreach (var item in data)
                {
                    var rowData = dataSelector(item);
                    if (colIndex < rowData.Length)
                    {
                        var cellText = rowData[colIndex]?.ToString() ?? "";
                        if (cellText.Length > maxLength)
                        {
                            maxLength = cellText.Length;
                        }
                    }
                }

                double excelWidth = CalculateExcelWidth(maxLength);

                columns.Append(new Column
                {
                    Min = (uint)(colIndex + 1),
                    Max = (uint)(colIndex + 1),
                    Width = excelWidth,
                    CustomWidth = true
                });
            }

            return columns;
        }

        //Вычисление необходимой ширины
        private double CalculateExcelWidth(int textLength)
        {
            const double baseWidth = 2.0;
            const double widthPerChar = 1.3;
            double calculatedWidth = baseWidth + (textLength * widthPerChar);
            return Math.Max(8.0, Math.Min(175.0, calculatedWidth));
        }

        //Содание стиля таблицы
        private Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();
            var fonts = new Fonts();

            // Шрифт 0: Обычный (Calibri 11)
            fonts.Append(new Font(
                new FontSize() { Val = 11 },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, // Черный
                new FontName() { Val = "Calibri" }
            ));

            // Шрифт 1: Жирный для заголовков
            fonts.Append(new Font(
                new Bold(),
                new FontSize() { Val = 11 },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = "Calibri" }
            ));


            // === ЗАЛИВКИ (Fills) ===
            var fills = new Fills();

            // Заливка 0: Пустая
            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None }));


            // === ГРАНИЦЫ (Borders) ===
            var borders = new Borders();

            // Граница 0: Без границ
            borders.Append(new Border());

            // Граница 1: Тонкая черная граница со всех сторон
            borders.Append(new Border(
                new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder()
            ));

            // Граница 2: Толстая граница только снизу
            borders.Append(new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                new DiagonalBorder()
            ));

            // === ФОРМАТЫ ЯЧЕЕК (CellFormats) ===
            var cellFormats = new CellFormats();

            // Формат 0: Обычный стиль (шрифт 0, без заливки, без границ)
            cellFormats.Append(new CellFormat
            {
                FontId = 0,
                BorderId = 0,
                ApplyFont = true
            });

            // Формат 1: Заголовки (жирный шрифт, все границы)
            cellFormats.Append(new CellFormat
            {
                FontId = 1,
                BorderId = 1,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true
            });

            // Формат 2: Данные (обычный шрифт, без заливки, все границы)
            cellFormats.Append(new CellFormat
            {
                FontId = 0,
                BorderId = 1,
                ApplyFont = true,
                ApplyBorder = true
            });

            // Формат 3: Важные ячейки (жирный, толстая граница снизу)
            cellFormats.Append(new CellFormat
            {
                FontId = 1,
                BorderId = 2,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true
            });

            // Собираем все вместе
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellFormats);

            return stylesheet;
        }

        // Получение типа данных ячеек
        private DocumentFormat.OpenXml.Spreadsheet.CellValues GetCellDataType(object value)
        {
            return (value is int || value is long || value is double || value is decimal || value is float)
                ? DocumentFormat.OpenXml.Spreadsheet.CellValues.Number
                : DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
        }
        #endregion
    }
    #endregion
}