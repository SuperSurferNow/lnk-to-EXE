using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WinForms = System.Windows.Forms;

namespace lnk_to_EXE
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<ShortcutItem> _shortcuts = new();
        private bool _isUpdatingSelection = false;

        public MainWindow()
        {
            InitializeComponent();
            ShortcutListBox.ItemsSource = _shortcuts;
            _shortcuts.CollectionChanged += (s, e) => UpdateUI();

            // Set default output folder
            OutputFolderTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        #region File Management

        private void AddFilesButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Shortcut Files (*.lnk)|*.lnk",
                Title = "Select Shortcut Files",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                AddShortcuts(dialog.FileNames);
            }
        }

        private void AddFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "Select folder containing .lnk files",
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                var lnkFiles = Directory.GetFiles(dialog.SelectedPath, "*.lnk", SearchOption.AllDirectories);
                AddShortcuts(lnkFiles);
            }
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (_shortcuts.Count > 0)
            {
                var result = System.Windows.MessageBox.Show(
                    $"Remove all {_shortcuts.Count} shortcuts from the list?",
                    "Clear All",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    _shortcuts.Clear();
                }
            }
        }

        private void RemoveSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            if (ShortcutListBox.SelectedItem is ShortcutItem item)
            {
                _shortcuts.Remove(item);
            }
        }

        private void AddShortcuts(string[] filePaths)
        {
            int addedCount = 0;
            int failedCount = 0;

            foreach (var path in filePaths)
            {
                try
                {
                    // Check for duplicates
                    if (_shortcuts.Any(s => s.SourcePath.Equals(path, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var info = ShortcutParser.Parse(path);
                    var item = new ShortcutItem(path, info);
                    _shortcuts.Add(item);
                    addedCount++;
                }
                catch (Exception ex)
                {
                    failedCount++;
                    StatusTextBlock.Text = $"Failed to load {Path.GetFileName(path)}: {ex.Message}";
                }
            }

            if (addedCount > 0)
            {
                StatusTextBlock.Text = $"Added {addedCount} shortcut(s)" +
                    (failedCount > 0 ? $", {failedCount} failed" : "");
            }
        }

        #endregion

        #region Drag & Drop

        private void Window_Drop(object sender, System.Windows.DragEventArgs e)
        {
            DropZoneOverlay.Visibility = Visibility.Collapsed;

            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                var lnkFiles = files.Where(f => f.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)).ToArray();

                if (lnkFiles.Length > 0)
                {
                    AddShortcuts(lnkFiles);
                }
                else
                {
                    System.Windows.MessageBox.Show("No .lnk files found in dropped items.", "Invalid Files",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void Window_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                if (files.Any(f => f.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effects = System.Windows.DragDropEffects.Copy;
                    DropZoneOverlay.Visibility = Visibility.Visible;
                }
                else
                {
                    e.Effects = System.Windows.DragDropEffects.None;
                }
            }
        }

        private void Window_DragLeave(object sender, System.Windows.DragEventArgs e)
        {
            DropZoneOverlay.Visibility = Visibility.Collapsed;
        }

        #endregion

        #region Selection & Editing

        private void ShortcutListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingSelection) return;

            RemoveSelectedButton.IsEnabled = ShortcutListBox.SelectedItem != null;

            if (ShortcutListBox.SelectedItem is ShortcutItem item)
            {
                LoadItemForEditing(item);
            }
            else
            {
                ClearEditor();
            }

            UpdateBuildButtons();
        }

        private void LoadItemForEditing(ShortcutItem item)
        {
            _isUpdatingSelection = true;

            EditTargetTextBox.Text = item.TargetPath;
            EditArgumentsTextBox.Text = item.Arguments;
            EditWorkingDirTextBox.Text = item.WorkingDirectory;
            EditIconPathTextBox.Text = item.IconPath;
            IconPreview.Source = item.IconSource;

            _isUpdatingSelection = false;
        }

        private void ClearEditor()
        {
            _isUpdatingSelection = true;

            EditTargetTextBox.Text = string.Empty;
            EditArgumentsTextBox.Text = string.Empty;
            EditWorkingDirTextBox.Text = string.Empty;
            EditIconPathTextBox.Text = string.Empty;
            IconPreview.Source = null;

            _isUpdatingSelection = false;
        }

        private void EditField_Changed(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingSelection || ShortcutListBox.SelectedItem is not ShortcutItem item)
                return;

            item.TargetPath = EditTargetTextBox.Text;
            item.Arguments = EditArgumentsTextBox.Text;
            item.WorkingDirectory = EditWorkingDirTextBox.Text;
            item.HasChanges = true;
        }

        private void ChangeIconButton_Click(object sender, RoutedEventArgs e)
        {
            if (ShortcutListBox.SelectedItem is not ShortcutItem item)
                return;

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Icon Files|*.ico;*.exe;*.dll|All Files|*.*",
                Title = "Select Icon File"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var icon = IconExtractor.ExtractIcon(dialog.FileName, 0);
                    item.CustomIconPath = dialog.FileName;
                    item.IconSource = icon;
                    IconPreview.Source = icon;
                    item.HasChanges = true;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Failed to load icon: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ResetIconButton_Click(object sender, RoutedEventArgs e)
        {
            if (ShortcutListBox.SelectedItem is not ShortcutItem item)
                return;

            item.CustomIconPath = null;

            // Try to extract icon from IconPath or fall back to target executable
            string iconPath = !string.IsNullOrEmpty(item.OriginalInfo.IconPath)
                ? item.OriginalInfo.IconPath
                : item.OriginalInfo.TargetPath;

            var originalIcon = IconExtractor.ExtractIcon(iconPath, item.OriginalInfo.IconIndex);
            item.IconSource = originalIcon;
            IconPreview.Source = originalIcon;
            item.HasChanges = true;
        }

        private void TestLaunchButton_Click(object sender, RoutedEventArgs e)
        {
            if (ShortcutListBox.SelectedItem is not ShortcutItem item)
                return;

            try
            {
                if (!File.Exists(item.TargetPath))
                {
                    System.Windows.MessageBox.Show($"Target file not found:\n{item.TargetPath}", "Target Not Found",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = item.TargetPath,
                    Arguments = item.Arguments,
                    WorkingDirectory = item.WorkingDirectory,
                    UseShellExecute = true
                });

                StatusTextBlock.Text = $"Launched: {Path.GetFileName(item.TargetPath)}";
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Failed to launch target:\n{ex.Message}", "Launch Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Building

        private void BrowseOutputFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "Select output folder for generated executables",
                SelectedPath = OutputFolderTextBox.Text
            };

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                OutputFolderTextBox.Text = dialog.SelectedPath;
            }
        }

        private async void BuildSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            if (ShortcutListBox.SelectedItem is ShortcutItem item)
            {
                await BuildItems(new[] { item });
            }
        }

        private async void BuildAllButton_Click(object sender, RoutedEventArgs e)
        {
            await BuildItems(_shortcuts.ToArray());
        }

        private async Task BuildItems(ShortcutItem[] items)
        {
            if (string.IsNullOrEmpty(OutputFolderTextBox.Text))
            {
                System.Windows.MessageBox.Show("Please select an output folder.", "No Output Folder",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DisableBuildUI();

            int successCount = 0;
            int failCount = 0;

            foreach (var item in items)
            {
                try
                {
                    item.Status = ConversionStatus.Building;
                    StatusTextBlock.Text = $"Building: {item.FileName}...";

                    string outputPath = Path.Combine(OutputFolderTextBox.Text,
                        Path.GetFileNameWithoutExtension(item.FileName) + ".exe");

                    var info = item.ToShortcutInfo();
                    await Task.Run(() => ExeBuilder.Build(info, outputPath, true));

                    item.Status = ConversionStatus.Success;
                    item.OutputPath = outputPath;
                    successCount++;
                }
                catch (Exception ex)
                {
                    item.Status = ConversionStatus.Failed;
                    item.ErrorMessage = ex.Message;
                    failCount++;
                }
            }

            EnableBuildUI();

            StatusTextBlock.Text = $"Build complete: {successCount} succeeded, {failCount} failed";

            if (successCount > 0)
            {
                var result = System.Windows.MessageBox.Show(
                    $"Successfully created {successCount} executable(s).\n\nOpen output folder?",
                    "Build Complete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", OutputFolderTextBox.Text);
                }
            }
        }

        #endregion

        #region UI Updates

        private void UpdateUI()
        {
            EmptyStateText.Visibility = _shortcuts.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            ItemCountText.Text = $"{_shortcuts.Count} shortcut(s) loaded";
            UpdateBuildButtons();
        }

        private void UpdateBuildButtons()
        {
            BuildSelectedButton.IsEnabled = ShortcutListBox.SelectedItem != null && !string.IsNullOrEmpty(OutputFolderTextBox.Text);
            BuildAllButton.IsEnabled = _shortcuts.Count > 0 && !string.IsNullOrEmpty(OutputFolderTextBox.Text);
        }

        private void DisableBuildUI()
        {
            BuildSelectedButton.IsEnabled = false;
            BuildAllButton.IsEnabled = false;
            AddFilesButton.IsEnabled = false;
            AddFolderButton.IsEnabled = false;
            ClearAllButton.IsEnabled = false;
            RemoveSelectedButton.IsEnabled = false;
        }

        private void EnableBuildUI()
        {
            AddFilesButton.IsEnabled = true;
            AddFolderButton.IsEnabled = true;
            ClearAllButton.IsEnabled = true;
            UpdateBuildButtons();
        }

        #endregion

        #region Theme

        private void DarkModeToggle_Click(object sender, RoutedEventArgs e)
        {
            bool isDarkMode = DarkModeToggle.IsChecked == true;
            ApplyTheme(isDarkMode);
            DarkModeToggle.Content = isDarkMode ? "☀️ Light Mode" : "🌙 Dark Mode";
        }

        private void ApplyTheme(bool isDarkMode)
        {
            if (isDarkMode)
            {
                // Dark mode colors
                Resources["WindowBackground"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 30, 30));
                Resources["TextColor"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220));
                Resources["BorderColor"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(60, 60, 60));
                Resources["SecondaryTextColor"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(150, 150, 150));
                Resources["GroupBoxBackground"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(40, 40, 40));
                Resources["ControlBackground"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(45, 45, 45));
                Resources["DisabledBackground"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(55, 55, 55));
                Resources["ListItemBackground"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(35, 35, 35));

                Background = (SolidColorBrush)Resources["WindowBackground"];
            }
            else
            {
                // Light mode colors
                Resources["WindowBackground"] = new SolidColorBrush(Colors.White);
                Resources["TextColor"] = new SolidColorBrush(Colors.Black);
                Resources["BorderColor"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(204, 204, 204));
                Resources["SecondaryTextColor"] = new SolidColorBrush(Colors.Gray);
                Resources["GroupBoxBackground"] = new SolidColorBrush(Colors.White);
                Resources["ControlBackground"] = new SolidColorBrush(Colors.White);
                Resources["DisabledBackground"] = new SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 245, 245));
                Resources["ListItemBackground"] = new SolidColorBrush(Colors.White);

                Background = (SolidColorBrush)Resources["WindowBackground"];
            }
        }

        #endregion
    }
}
