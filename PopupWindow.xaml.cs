using Microsoft.Win32;
using System.IO;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

// https://www.wpf-tutorial.com/rich-text-controls/how-to-creating-a-rich-text-editor/

namespace myWPF
{
    /// <summary>
    /// Interaction logic for PopupWindow.xaml
    /// </summary>
    public partial class PopupWindow : Window
    {
        public MainWindow parent;

        public PopupWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new();
            saveFileDialog.Filter = "Microsoft Word dokumentum (*.docx)|*.docx";
            saveFileDialog.InitialDirectory = Directory.GetCurrentDirectory().ToString();
            if (saveFileDialog.ShowDialog() == true)
            {
                string WordDocPath = saveFileDialog.FileName;

                Word.Application wordApp = new();
                wordApp.ShowAnimation = false;
                wordApp.Visible = false;
                Word.Document wordDoc = new();

                // popupDoc -> wordDoc
                Rtb.SelectAll();
                Rtb.Copy();
                Word.Paragraph para = wordDoc.Paragraphs.Add();
                para.Range.Paste();

                wordDoc.SaveAs2(WordDocPath);
                wordDoc.Close(0, 0, 0);
                wordDoc = null;
                wordApp.Quit(0, 0, 0);
                wordApp = null;

                if (parent != null)
                {
                    parent.SaveTemplate(WordDocPath.Remove(WordDocPath.Length - 5) + ".txt");
                }

                MessageBox.Show("Szöveg lementve!");
            }
        }
    }
}
