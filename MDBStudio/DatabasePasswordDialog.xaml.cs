using System.Windows;

namespace MDBStudio;

public partial class DatabasePasswordDialog : Window
{
    // Password is kept only in the current application process, never stored on disk.
    public string EnteredPassword { get; private set; } = string.Empty;

    public DatabasePasswordDialog(string fileName)
    {
        InitializeComponent();
        Title = "Database password | " + fileName;
        Loaded += (_, _) => PasswordInput.Focus();
    }

    private void Submit_Click(object sender, RoutedEventArgs e)
    {
        if (PasswordInput.Password.Length == 0)
        {
            MessageBox.Show(this, "Enter the database password, or select Cancel.",
                "Password required", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        EnteredPassword = PasswordInput.Password;
        PasswordInput.Clear();
        DialogResult = true;
    }
}
