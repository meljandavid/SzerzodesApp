using System.Windows;

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
            Clipboard.SetText(popupTextbox.Text);
            MessageBox.Show("Szöveg kimásolva!");
        }
    }
}
