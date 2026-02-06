using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using WpfBrush = System.Windows.Media.Brush;
using WpfBrushes = System.Windows.Media.Brushes;

namespace lnk_to_EXE
{
    public enum ConversionStatus
    {
        Ready,
        Building,
        Success,
        Failed
    }

    public class ShortcutItem : INotifyPropertyChanged
    {
        private string _targetPath;
        private string _arguments;
        private string _workingDirectory;
        private string? _customIconPath;
        private System.Windows.Media.ImageSource? _iconSource;
        private ConversionStatus _status = ConversionStatus.Ready;
        private string? _errorMessage;
        private bool _hasChanges;

        public string SourcePath { get; }
        public string FileName => Path.GetFileName(SourcePath);
        public ShortcutInfo OriginalInfo { get; }
        public string? OutputPath { get; set; }

        public string TargetPath
        {
            get => _targetPath;
            set { _targetPath = value; OnPropertyChanged(); }
        }

        public string Arguments
        {
            get => _arguments;
            set { _arguments = value; OnPropertyChanged(); }
        }

        public string WorkingDirectory
        {
            get => _workingDirectory;
            set { _workingDirectory = value; OnPropertyChanged(); }
        }

        public string IconPath => _customIconPath ?? OriginalInfo.IconPath;

        public string? CustomIconPath
        {
            get => _customIconPath;
            set { _customIconPath = value; OnPropertyChanged(); OnPropertyChanged(nameof(IconPath)); }
        }

        public System.Windows.Media.ImageSource? IconSource
        {
            get => _iconSource;
            set { _iconSource = value; OnPropertyChanged(); }
        }

        public ConversionStatus Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(StatusText));
                OnPropertyChanged(nameof(StatusColor));
            }
        }

        public string? ErrorMessage
        {
            get => _errorMessage;
            set { _errorMessage = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusText)); }
        }

        public bool HasChanges
        {
            get => _hasChanges;
            set { _hasChanges = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusText)); }
        }

        public string StatusText => Status switch
        {
            ConversionStatus.Building => "? Building...",
            ConversionStatus.Success => "? Success",
            ConversionStatus.Failed => $"? Failed: {ErrorMessage}",
            _ => HasChanges ? "?? Modified" : "? Ready"
        };

        public WpfBrush StatusColor => Status switch
        {
            ConversionStatus.Building => WpfBrushes.Orange,
            ConversionStatus.Success => WpfBrushes.Green,
            ConversionStatus.Failed => WpfBrushes.Red,
            _ => HasChanges ? WpfBrushes.Blue : WpfBrushes.Gray
        };

        public ShortcutItem(string sourcePath, ShortcutInfo info)
        {
            SourcePath = sourcePath;
            OriginalInfo = info;
            _targetPath = info.TargetPath;
            _arguments = info.Arguments;
            _workingDirectory = info.WorkingDirectory;

            // Extract icon - if no icon path specified, try to extract from target executable
            try
            {
                string iconPath = !string.IsNullOrEmpty(info.IconPath) 
                    ? info.IconPath 
                    : info.TargetPath;
                    
                _iconSource = IconExtractor.ExtractIcon(iconPath, info.IconIndex);
            }
            catch
            {
                // Use default icon if extraction fails
            }
        }

        public ShortcutInfo ToShortcutInfo()
        {
            return new ShortcutInfo
            {
                TargetPath = TargetPath,
                Arguments = Arguments,
                WorkingDirectory = WorkingDirectory,
                IconPath = IconPath,
                IconIndex = CustomIconPath != null ? 0 : OriginalInfo.IconIndex
            };
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
