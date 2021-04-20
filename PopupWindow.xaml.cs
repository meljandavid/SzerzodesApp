using System.IO;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace myWPF
{
    /// <summary>
    /// Interaction logic for PopupWindow.xaml
    /// </summary>
    public partial class PopupWindow : Window
    {
        public PopupWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Word.Application wordApp = new();
            wordApp.ShowAnimation = false;
            wordApp.Visible = false;
            Word.Document wordDoc = new();

            // popupDoc -> wordDoc
            Rtb.SelectAll();
            Rtb.Copy();
            Word.Paragraph para = wordDoc.Paragraphs.Add();
            para.Range.Paste();

            wordDoc.SaveAs2(Directory.GetCurrentDirectory().ToString() + "\\elkeszult.docx");
            wordDoc.Close(0, 0, 0);
            wordDoc = null;
            wordApp.Quit(0, 0, 0);
            wordApp = null;

            MessageBox.Show("Szöveg lementve!");
        }
    }
}
