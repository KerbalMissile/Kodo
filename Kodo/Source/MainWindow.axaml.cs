// Licensed under the Kodo Public License v1.0
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
    private string _pendingFileName = string.Empty;
    private bool _hasDocumentOpen;
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
            if (_editorContent == value) return;
            _editorContent = value;
            OnPropertyChanged();
        }
    }

    public string PendingFileName
    {
        get => _pendingFileName;
        set
        {
            if (_pendingFileName == value) return;
            _pendingFileName = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(FileSummaryText));
        }
    }

    public bool IsSettingsPageVisible
    {
        get => _isSettingsPageVisible;
        set
        {
            if (_isSettingsPageVisible == value) return;
            _isSettingsPageVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsEditorPageVisible));
        }
    }

    public bool IsEditorPageVisible => !IsSettingsPageVisible;
    public bool HasFileOpen => _hasDocumentOpen;
    public bool IsEmptyStateVisible => !_hasDocumentOpen;
    public bool ShowTitleInput => _hasDocumentOpen && _currentFilePath is null;
    public bool ShowReadOnlyTitle => _hasDocumentOpen && _currentFilePath is not null;

    public string CurrentThemeName
    {
        get => _currentThemeName;
        private set
        {
            if (_currentThemeName == value) return;
            _currentThemeName = value;
            OnPropertyChanged();
        }
    }

    public string FileSummaryText => _hasDocumentOpen
        ? $"{GetDisplayName()}{(_isDirty ? " • unsaved" : string.Empty)}"
        : "One file at a time";

    public string FilePathText => _hasDocumentOpen
        ? _currentFilePath ?? "Unsaved file"
        : "No file open";

    public string ThemeStatusText => $"Current theme: {CurrentThemeName}";

    public string EditorStatsText
    {
        get => _editorStatsText;
        private set
        {
            if (_editorStatsText == value) return;
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
        var lines = _editorContent.Length == 0 ? 1 : _editorContent.Count(static c => c == '\n') + 1;
        EditorStatsText = $"{lines} lines  |  {_editorContent.Length} characters";
        Title = _hasDocumentOpen ? $"{GetDisplayName()} - Kodo" : "Kodo";
        OnPropertyChanged(nameof(HasFileOpen));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(ShowTitleInput));
        OnPropertyChanged(nameof(ShowReadOnlyTitle));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        OnPropertyChanged(nameof(ThemeStatusText));
    }

    private string GetDisplayName()
    {
        if (_currentFilePath is not null)
            return Path.GetFileName(_currentFilePath);
        return string.IsNullOrWhiteSpace(_pendingFileName) ? "Untitled" : _pendingFileName;
    }

    private void ApplyTheme(string themeName)
    {
        CurrentThemeName = themeName == "Light" ? "Light" : "Dark";
        Application.Current!.RequestedThemeVariant = CurrentThemeName == "Light"
            ? ThemeVariant.Light
            : ThemeVariant.Dark;

        if (CurrentThemeName == "Light")
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
        else
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

    private async Task OpenFileAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open File",
            AllowMultiple = false
        });

        var file = files.Count > 0 ? files[0] : null;
        if (file is null) return;

        var path = file.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;

        _currentFilePath = path;
        _pendingFileName = Path.GetFileName(path);
        _hasDocumentOpen = true;
        EditorContent = await File.ReadAllTextAsync(path);
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
    }

    // Opens a blank editor immediately. The header TextBox (FileNameTextBox) is
    // automatically shown when HasFileOpen=true and _currentFilePath=null,
    // and it gets focused so the user names the file before writing.
    private void NewFile()
    {
        _currentFilePath = null;
        _pendingFileName = string.Empty;
        _hasDocumentOpen = true;
        _editorContent = string.Empty;
        _isDirty = false;
        IsSettingsPageVisible = false;
        // Notify EditorContent directly since we set the field to bypass the early-return guard
        OnPropertyChanged(nameof(EditorContent));
        RefreshState();
        // FileNameTextBox is in the always-visible header, so Focus() is safe here
        FileNameTextBox.Focus();
    }

    private async Task SaveAsync()
    {
        if (_currentFilePath is null)
        {
            var suggestedName = string.IsNullOrWhiteSpace(_pendingFileName)
                ? "untitled.txt"
                : _pendingFileName.Trim();

            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save File",
                SuggestedFileName = suggestedName
            });

            var newPath = file?.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(newPath)) return;

            _currentFilePath = newPath;
            _pendingFileName = Path.GetFileName(newPath);
        }

        await File.WriteAllTextAsync(_currentFilePath, _editorContent);
        _hasDocumentOpen = true;
        _isDirty = false;
        RefreshState();
    }

    private void EditorButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsSettingsPageVisible = false;

    private async void OpenFileButton_OnClick(object? sender, RoutedEventArgs e) =>
        await OpenFileAsync();

    private async void SaveButton_OnClick(object? sender, RoutedEventArgs e) =>
        await SaveAsync();

    private void NewFileButton_OnClick(object? sender, RoutedEventArgs e) =>
        NewFile();

    private void SettingsButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsSettingsPageVisible = true;

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsSettingsPageVisible = false;

    private void ThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { Tag: string themeName })
            ApplyTheme(themeName);
    }

    private void EditorTextBox_OnTextChanged(object? sender, TextChangedEventArgs e)
    {
        _isDirty = true;
        RefreshState();
    }
}
