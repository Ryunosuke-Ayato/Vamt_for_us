using Microsoft.EntityFrameworkCore;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Vamt_for_us.Data;
using Vamt_for_us.ViewModels;
namespace Vamt_for_us
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            // Инициализируем ViewModel
            _viewModel = new MainViewModel();
            // Подключаем ViewModel к контексту окна
            DataContext = _viewModel;
            Loaded += MainWindow_Loaded;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadDataAsync();
        }

        // Асинхронная подгрузка данных.
        private async Task LoadDataAsync()
        {
            try
            {
                _viewModel.IsLoading = true;
                await _viewModel.LoadProductKeysDataAsync();
                await _viewModel.LoadActiveProductsDataAsync();
                ProductKeysListBox.ItemsSource = _viewModel.ProductKeys;
                ActiveProductsListBox.ItemsSource = _viewModel.ActiveProducts;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _viewModel.IsLoading = false;
            }
        }

        // Метод обработки выбранного ключа
        private void ProductKeysListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProductKeysListBox.SelectedItem is ProductKeyViewModel selectedProduct)
            {
                ProductKeyDetailsGroup.Visibility = Visibility.Visible;
                NoSelectionProductKeyText.Visibility = Visibility.Collapsed;

                SelectedKeyText.Text = selectedProduct.KeyValue;
                SelectedDescriptionText.Text = selectedProduct.KeyDescription;
                SelectedActivationsText.Text = selectedProduct.RemainingActivations.ToString();
                CommentTextBox.Text = selectedProduct.UserRemarks;

                _viewModel.SelectedProductKey = selectedProduct;
                _viewModel.Comment = selectedProduct.UserRemarks;
            }
            else
            {
                ProductKeyDetailsGroup.Visibility = Visibility.Collapsed;
                NoSelectionProductKeyText.Visibility = Visibility.Visible;
                _viewModel.SelectedProductKey = null;
            }
        }

        // Метод обработки выбранного пк с продукцией
        private void ActiveProductsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ActiveProductsListBox.SelectedItem is ActiveProductViewModel selectedProduct)
            {
                ActiveProductDetailsGroup.Visibility = Visibility.Visible;
                NoSelectionActiveProductText.Visibility = Visibility.Collapsed;

                SelectedDomainText.Text = selectedProduct.FullyQualifiedDomainName;
                SelectedProductKeyText.Text = selectedProduct.ProductKeyId;
                SelectedLicenseStatusText.Text = selectedProduct.LicenseStatusText;
                SelectedKeyTypeText.Text = selectedProduct.KeyTypeName;

                _viewModel.SelectedActiveProduct = selectedProduct;
            }
            else
            {
                ActiveProductDetailsGroup.Visibility = Visibility.Collapsed;
                NoSelectionActiveProductText.Visibility = Visibility.Visible;
                _viewModel.SelectedActiveProduct = null;
            }
        }

        // Обновление списков
        private async void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadDataAsync();
        }

        // Метод сохранения комментария к ключу
        private async void SaveCommentButton_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.SelectedProductKey != null)
            {
                _viewModel.Comment = CommentTextBox.Text;
                await _viewModel.UpdateCommentAsync();
            }
        }

        // Метод сохранения комментария к ключу через CTRL+Enter
        private async void CommentTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (_viewModel.SelectedProductKey != null)
                {
                    _viewModel.Comment = CommentTextBox.Text;
                    await _viewModel.UpdateCommentAsync();
                }
            }
        }

        // Предварительное копирвоание 
        private void CopyProductKey_Click(object sender, RoutedEventArgs e)
        {
            CopyTextToClipboard(SelectedKeyText.Text);
        }
        private void CopyDomain_Click(object sender, RoutedEventArgs e)
        {
            CopyTextToClipboard(SelectedDomainText.Text);
        }

        // Копируем полученный текст в буфер обмена
        private void CopyTextToClipboard(string text)
        {
            if (string.IsNullOrEmpty(text)) return;

            try
            {
                Clipboard.SetText(text);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                System.Diagnostics.Debug.WriteLine("Не удалось скопировать в буфер обмена");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка копирования: {ex.Message}");
            }
        }

        //Методы генерации отчётов
        private async void ExportLicenseReport_Click(object sender, RoutedEventArgs e)
        {
            await _viewModel.ExportLicenseReportAsync();
        }
        private async void ExportActivatedReport_Click(object sender, RoutedEventArgs e)
        {
            await _viewModel.ExportActivatedReportAsync();
        }

        // Временная заглушка
        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            // ВНИМАНИЕ, ИНФОМРАЦИЯ - "Спасибо за внимание"!
        }

        #region Регион настроек
        
        // Переключение видимости панели натроек
        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (SettingsGroupBox.Visibility == Visibility.Visible)
            {
                SettingsGroupBox.Visibility = Visibility.Collapsed;
            }
            else
            {
                SettingsGroupBox.Visibility = Visibility.Visible;
                ConnectionStringTextBox.Text = Vamt_for_us.Services.ConfigurationService.ConnectionString;
                ConnectionStringTextBox.Focus();
            }
        }

        // Метод проверки строки подключения
        private async void TestConnection_Click(object sender, RoutedEventArgs e)
        {
            var connectionString = ConnectionStringTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                MessageBox.Show("Введите строку подключения", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _viewModel.IsLoading = true;

                var optionsBuilder = new DbContextOptionsBuilder<VamtDbContext>();
                optionsBuilder.UseSqlServer(connectionString);

                using var testContext = new VamtDbContext(optionsBuilder.Options);
                var canConnect = await testContext.Database.CanConnectAsync();

                if (canConnect)
                {
                    MessageBox.Show("✅ Подключение успешно установлено!", "Тестирование подключения",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("❌ Не удалось подключиться к базе данных", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Ошибка подключения:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _viewModel.IsLoading = false;
            }
        }

        // Сохранение строки подключения в конфиге
        private void SaveConnection_Click(object sender, RoutedEventArgs e)
        {
            var newConnectionString = ConnectionStringTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(newConnectionString))
            {
                MessageBox.Show("Строка подключения не может быть пустой", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Vamt_for_us.Services.ConfigurationService.UpdateConnectionString(newConnectionString);
                MessageBox.Show("✅ Настройки сохранены!\nПрименение изменений требует перезапуска приложения.",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                SettingsGroupBox.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Не удалось сохранить настройки:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Скрытие настроек без сохранения
        private void CancelSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsGroupBox.Visibility = Visibility.Collapsed;
        }
        
        // Сброс строки подключения
        private void ResetConnection_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Сбросить строку подключения к значению по умолчанию?",
                "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                ConnectionStringTextBox.Text = "";
            }
        }
        #endregion
    }
}