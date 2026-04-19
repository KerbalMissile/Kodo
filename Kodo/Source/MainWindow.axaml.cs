// Licensed under the Kodo Public License v1.0
// April 19th, 2026 - KerbalMissile - Changed "One file at a time" note to "No file open"
// April 19th, 2026 - KerbalMissile - Added proper comments
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Styling;

namespace Kodo;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private string? _currentFilePath;
    private string _editorContent = string.Empty;
    private bool _isDirty;
    private bool _isSettingsPageVisible;
    private string _currentThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private event PropertyChangedEventHandler? ViewModelPropertyChanged;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        ApplyTheme("Dark");
        RefreshState();
    }

    public string EditorContent
    {
        get => _editorContent;
        set
        {
            if (_editorContent == value)
            {
                return;
            }

            _editorContent = value;
            OnPropertyChanged();
        }
    }

    public bool IsSettingsPageVisible
    {
        get => _isSettingsPageVisible;
        set
        {
            if (_isSettingsPageVisible == value)
            {
                return;
            }

            _isSettingsPageVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsEditorPageVisible));
        }
    }

    public bool IsEditorPageVisible => !IsSettingsPageVisible;

    public bool HasFileOpen => _currentFilePath is not null;

    public bool IsEmptyStateVisible => !HasFileOpen;

    public string CurrentThemeName
    {
        get => _currentThemeName;
        private set
        {
            if (_currentThemeName == value)
            {
                return;
            }

            _currentThemeName = value;
            OnPropertyChanged();
        }
    }

    // Displays the file name and unsaved status in the top bar
    public string FileSummaryText => HasFileOpen
        ? $"{Path.GetFileName(_currentFilePath!)}{(_isDirty ? " • unsaved" : string.Empty)}"
        : "No file open";

    public string FilePathText => HasFileOpen ? _currentFilePath! : "No file open";

    public string ThemeStatusText => $"Current theme: {CurrentThemeName}";

    public string EditorStatsText
    {
        get => _editorStatsText;
        private set
        {
            if (_editorStatsText == value)
            {
                return;
            }

            _editorStatsText = value;
            OnPropertyChanged();
        }
    }

    public IBrush WindowBackgroundBrush { get; private set; } = Brush.Parse("#1E1E1E");
    public IBrush TopBarBrush { get; private set; } = Brush.Parse("#181818");
    public IBrush SidebarBrush { get; private set; } = Brush.Parse("#181818");
    public IBrush ButtonBrush { get; private set; } = Brush.Parse("#252526");
    public IBrush ButtonHoverBrush { get; private set; } = Brush.Parse("#313437");
    public IBrush EditorBackgroundBrush { get; private set; } = Brush.Parse("#1E1E1E");
    public IBrush CardBrush { get; private set; } = Brush.Parse("#252526");
    public IBrush PrimaryTextBrush { get; private set; } = Brush.Parse("#F4F4F4");
    public IBrush MutedTextBrush { get; private set; } = Brush.Parse("#A0A0A0");
    public IBrush SurfaceBorderBrush { get; private set; } = Brush.Parse("#2B2B2B");
    public IBrush AccentBrush { get; private set; } = Brush.Parse("#0E639C");

    event PropertyChangedEventHandler? INotifyPropertyChanged.PropertyChanged
    {
        add => ViewModelPropertyChanged += value;
        remove => ViewModelPropertyChanged -= value;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        ViewModelPropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private void RefreshState()
    {
        var lines = EditorContent.Length == 0 ? 1 : EditorContent.Count(static c => c == '\n') + 1;
        EditorStatsText = $"{lines} lines  |  {EditorContent.Length} characters";
        Title = HasFileOpen ? $"{Path.GetFileName(_currentFilePath!)} - Kodo" : "Kodo";
        OnPropertyChanged(nameof(HasFileOpen));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        OnPropertyChanged(nameof(ThemeStatusText));
    }

    // Light Mode and Dark Mode color definitions
    private void ApplyTheme(string themeName)
    {
        CurrentThemeName = themeName == "Light" ? "Light" : "Dark";
        Application.Current!.RequestedThemeVariant = CurrentThemeName == "Light"
            ? ThemeVariant.Light
            : ThemeVariant.Dark;

        if (CurrentThemeName == "Light") // Light mode colours
        {
            WindowBackgroundBrush = Brush.Parse("#F3F3F3");
            TopBarBrush = Brush.Parse("#FFFFFF");
            SidebarBrush = Brush.Parse("#EFF2F7");
            ButtonBrush = Brush.Parse("#E3E8F1");
            ButtonHoverBrush = Brush.Parse("#D5DDE9");
            EditorBackgroundBrush = Brush.Parse("#FFFFFF");
            CardBrush = Brush.Parse("#F7F9FC");
            PrimaryTextBrush = Brush.Parse("#202124");
            MutedTextBrush = Brush.Parse("#5F6B7A");
            SurfaceBorderBrush = Brush.Parse("#D7DCE5");
            AccentBrush = Brush.Parse("#0067C0");
        }
        else // Dark mode colours
        {
            WindowBackgroundBrush = Brush.Parse("#1E1E1E");
            TopBarBrush = Brush.Parse("#181818");
            SidebarBrush = Brush.Parse("#181818");
            ButtonBrush = Brush.Parse("#252526");
            ButtonHoverBrush = Brush.Parse("#313437");
            EditorBackgroundBrush = Brush.Parse("#1E1E1E");
            CardBrush = Brush.Parse("#252526");
            PrimaryTextBrush = Brush.Parse("#F4F4F4");
            MutedTextBrush = Brush.Parse("#A0A0A0");
            SurfaceBorderBrush = Brush.Parse("#2B2B2B");
            AccentBrush = Brush.Parse("#0E639C");
        }

        // Notify the UI that all the brushes have changed
        OnPropertyChanged(nameof(WindowBackgroundBrush));
        OnPropertyChanged(nameof(TopBarBrush));
        OnPropertyChanged(nameof(SidebarBrush));
        OnPropertyChanged(nameof(ButtonBrush));
        OnPropertyChanged(nameof(ButtonHoverBrush));
        OnPropertyChanged(nameof(EditorBackgroundBrush));
        OnPropertyChanged(nameof(CardBrush));
        OnPropertyChanged(nameof(PrimaryTextBrush));
        OnPropertyChanged(nameof(MutedTextBrush));
        OnPropertyChanged(nameof(SurfaceBorderBrush));
        OnPropertyChanged(nameof(AccentBrush));
        RefreshState();
    }
    
    // Opens a file picker dialog and loads the selected file into the editor
    private async Task OpenFileAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open File",
            AllowMultiple = false
        });

        var file = files.Count > 0 ? files[0] : null;
        if (file is null)
        {
            return;
        }

        var path = file.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return;
        }

        _currentFilePath = path;
        EditorContent = await File.ReadAllTextAsync(path);
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
    }

    // Clears the editor and resets the state to allow creating a new file
    // Currently is BROKEN, new file does not work
    private void NewFile()
    {
        _currentFilePath = null;
        EditorContent = string.Empty;
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
        EditorTextBox.Focus();
    }

    // Saves the current editor content to the open file, or prompts for a file path if no file is open
    private async Task SaveAsync()
    {
        if (_currentFilePath is null)
        {
            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save File",
                SuggestedFileName = "untitled.txt"
            });

            var newPath = file?.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(newPath))
            {
                return;
            }

            _currentFilePath = newPath;
        }

        await File.WriteAllTextAsync(_currentFilePath, EditorContent);
        _isDirty = false;
        RefreshState();
    }
    
    // Event handlers for UI interactions
    private void EditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
    }

    private async void OpenFileButton_OnClick(object? sender, RoutedEventArgs e)
    {
        await OpenFileAsync();
    }

    private async void SaveButton_OnClick(object? sender, RoutedEventArgs e)
    {
        await SaveAsync();
    }

    private void NewFileButton_OnClick(object? sender, RoutedEventArgs e)
    {
        NewFile();
    }

    private void SettingsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = true;
    }

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
    }

    // Handles theme selection buttons in the settings page
    private void ThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { Tag: string themeName })
        {
            ApplyTheme(themeName);
        }
    }

    // Marks the editor content as dirty (unsaved) whenever it changes
    private void EditorTextBox_OnTextChanged(object? sender, TextChangedEventArgs e)
    {
        _isDirty = true;
        RefreshState();
    }
}
