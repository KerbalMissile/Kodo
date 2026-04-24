// Licensed under the Kodo Public License v1.0
// April 19th, 2026 - KerbalMissile - Added proper comments
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;

namespace Kodo;

// One of the main entry points for the application, responsible for initializing and starting the app. This class is mostly boilerplate for Avalonia applications.
public partial class App : Application

    // Initializes the application by loading the XAML defined in App.axaml, which sets up resources and styles.
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    // Called when the application has finished initializing. Here we check if we're running in a desktop environment and if so, we create and show the main window of the application.
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new MainWindow();
        }

        base.OnFrameworkInitializationCompleted();
    }
}