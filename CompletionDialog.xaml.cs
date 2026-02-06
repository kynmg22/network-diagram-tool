using System.Windows;

namespace NetworkDiagramApp
{
    public enum CompletionChoice
    {
        OpenFolder,
        OpenFile,
        Close
    }

    public partial class CompletionDialog : Window
    {
        public CompletionChoice UserChoice { get; private set; }
        private readonly string _filePath;

        public CompletionDialog(string filePath)
        {
            InitializeComponent();
            _filePath = filePath;
            TxtFilePath.Text = filePath;
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            UserChoice = CompletionChoice.OpenFolder;
            DialogResult = true;
            Close();
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            UserChoice = CompletionChoice.OpenFile;
            DialogResult = true;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            UserChoice = CompletionChoice.Close;
            DialogResult = true;
            Close();
        }
    }
}
