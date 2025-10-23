using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Runtime.CompilerServices;
using System.Windows.Media;
namespace Vamt_for_us.Models
{
    #region Модели БД

    // Модель Ключей
    [Table("ProductKey", Schema = "base")]
    public class ProductKey
    {
        [Key]
        [Column("KeyId")] // Явно указываем имя столбца
        public string KeyId { get; set; }

        [Column("KeyValue")]
        public string KeyValue { get; set; } = string.Empty;

        [Column("KeyDescription")]
        public string KeyDescription { get; set; } = string.Empty;

        [Column("RemainingActivations")]
        public int RemainingActivations { get; set; }

        [Column("UserRemarks")]
        public string UserRemarks { get; set; } = string.Empty;


    }

    // Модель Продуктов
    [Table("ActiveProduct", Schema = "base")]
    public class ActiveProduct
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        [Column("FullyQualifiedDomainName")]
        public string FullyQualifiedDomainName { get; set; } = string.Empty;

        [Column("ProductKeyId")]
        public string ProductKeyId { get; set; } = string.Empty;

        [Column("LicenseStatus")]
        public string LicenseStatus { get; set; } = string.Empty;

        [Column("ProductKeyType")]
        public int ProductKeyType { get; set; }
    }

    // Модель статуса лицензии
    [Table("LicenseStatusText", Schema = "base")]
    public class LicenseStatusText
    {
        [Key]
        [Column("LicenseStatus")]
        public string LicenseStatus { get; set; } = string.Empty;

        [Column("LicenseStatusText")]
        public string LicenseStatusTextValue { get; set; } = string.Empty;
    }

    //Модель Типа ключа
    [Table("ProductKeyTypeName", Schema = "base")]
    public class ProductKeyTypeName
    {
        [Key]
        [Column("KeyType")]
        public int KeyType { get; set; }

        [Column("KeyTypeName")]
        public string KeyTypeName { get; set; } = string.Empty;
    }
    #endregion

    // Класс ПК для доменной утилиты
    public class ComputerInfo : INotifyPropertyChanged
    {
        private string _name;
        private string _ipAddress;
        private bool _isOnline;
        private bool _isSelected;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged();
            }
        }

        public string IPAddress
        {
            get => _ipAddress;
            set
            {
                _ipAddress = value;
                OnPropertyChanged();
            }
        }

        public bool IsOnline
        {
            get => _isOnline;
            set
            {
                _isOnline = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(Status));
                OnPropertyChanged(nameof(StatusColor));
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
                Console.WriteLine($"Computer {Name} IsSelected changed to: {value}"); // Отладочное сообщение
            }
        }

        public string Status => IsOnline ? "Online" : "Offline";
        public Brush StatusColor => IsOnline ? Brushes.Green : Brushes.Red;

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}