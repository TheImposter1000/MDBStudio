using System.Windows;

namespace MDBStudio;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        // Windows Explorer's "open" shell command passes the clicked MDB
        // path in quotes as one argument. Pass that path to the normal MDB
        // opening flow rather than launching an empty application window.
        string? path = WindowsFileAssociation.DatabasePathFromArguments(e.Args);
        var window = new MainWindow(path);
        MainWindow = window;
        window.Show();
    }
}
