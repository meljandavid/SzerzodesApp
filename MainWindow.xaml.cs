using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace myWPF
{
    public partial class MainWindow : Window
    {
        List<Szempont> szempontok = new();
        PopupWindow popup;

        private static Action EmptyDelegate = delegate () { };

        void Load()
        {
            try
            {
                panel_szempontok.Children.Clear();
                panel_opciok.Children.Clear();
                szempontok.Clear();

                Excel.Application xlApp = new();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(
                    Directory.GetCurrentDirectory().ToString() + "\\szerzodes_szoveg_sablon.xlsx");
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];
                Excel.Range excelRange = xlWorksheet.UsedRange;

                object[,] valueArray = (object[,])excelRange.get_Value(
                            Excel.XlRangeValueDataType.xlRangeValueDefault);

                // excel is not zero based
                for (int row = 2; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
                {
                    String rsz = "";

                    // Ha  van uj szempont
                    if (valueArray[row, 1] != null)
                    {
                        rsz = valueArray[row, 1].ToString();
                    }

                    if (rsz != "")
                    {
                        szempontok.Add(new Szempont(rsz));
                        szempontok[^1].opciok = new List<Opcio>();
                    }

                    if(valueArray[row, 2].ToString().Contains("Adatbázis") )
                    {

                    }
                    else
                    {
                        // Mindig van uj opcio, csinalunk egy ujat
                        Opcio opcio = new(valueArray[row, 2].ToString());
                        String str = "";
                        if (xlWorksheet.Range["C" + row.ToString()] != null)
                        {
                            xlWorksheet.Range["C" + row.ToString()].Copy();
                            str = Clipboard.GetText(TextDataFormat.Rtf);
                        }
                        opcio.kifejtve = str;
                        szempontok[^1].opciok.Add(opcio);

                        Clipboard.Clear();
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //close and release
                xlWorkbook.Close();
                xlApp.Quit();
            }
            catch (Exception e)
            {
                MessageBox.Show("EXCEL ERROR:\n" + e.Message);
            }

            foreach (Szempont sz in szempontok)
            {
                panel_szempontok.Children.Add(sz.rb);
                sz.rb.Click += CheckRadioSzempont_Click;
                foreach (Opcio o in sz.opciok)
                {
                    o.cb.Click += CheckRadioOpcio_Click;
                }
            }

        }

        public MainWindow()
        {
            LoadingWindow loadwnd = new();
            loadwnd.Show();

            InitializeComponent();
            myrichtext.IsReadOnly = true;

            Load();

            loadwnd.Close();
        }

        class Opcio
        {
            public Opcio(String _id)
            {
                cb = new()
                {
                    Content = _id,
                    FontSize = 14
                };
                Thickness margin = cb.Margin;
                margin.Left = 10;
                margin.Bottom = 2;
                cb.Margin = margin;
                id = _id;
            }

            public string id, kifejtve;
            public CheckBox cb;
        }

        class Szempont
        {
            public Szempont(String _id)
            {
                rb = new RadioButton
                {
                    Content = _id,
                    FontSize = 14
                };
                id = _id;
                isChecked = false;
            }

            public bool isChecked;
            public string id;
            public RadioButton rb;
            public List<Opcio> opciok;
            public void SzChecked(bool state)
            {
                if (state)
                {
                    rb.Foreground = new SolidColorBrush(Colors.Green);
                    rb.FontWeight = FontWeights.Bold;
                    rb.Background = new SolidColorBrush(Colors.LightGreen);
                    isChecked = true;
                }
                else
                {
                    rb.Foreground = new SolidColorBrush(Colors.Black);
                    rb.FontWeight = FontWeights.Normal;
                    rb.Background = new SolidColorBrush(Colors.White);
                    isChecked = false;
                }
            }
        }

        private void CheckRadioOpcio_Click(object sender, RoutedEventArgs e)
        {
            foreach (Szempont sz in szempontok)
            {
                if (sz.rb.IsChecked == true)
                {
                    int ctr = 0;
                    foreach (Opcio o in sz.opciok)
                    {
                        if (o.cb.IsChecked == true)
                        {
                            ctr++;
                            TextRange tr = new(MyDoc.ContentStart, MyDoc.ContentEnd);
                            byte[] byteArray = Encoding.ASCII.GetBytes(o.kifejtve);
                            MemoryStream stream = new(byteArray);
                            tr.Load(stream, DataFormats.Rtf);
                        }
                    }
                    if(ctr == 0)
                    {
                        sz.SzChecked(false);
                    }
                    else
                    {
                        sz.SzChecked(true);
                    }
                }
            }
        }

        private void CheckRadioSzempont_Click(object sender, RoutedEventArgs e)
        {
            panel_opciok.Children.Clear();
            foreach (Szempont sz in szempontok)
            {
                if (sz.rb.IsChecked == true)
                {
                    bool found = false;
                    foreach (Opcio o in sz.opciok)
                    {
                        panel_opciok.Children.Add(o.cb);
                        if (o.cb.IsChecked == true)
                        {
                            found = true;
                            TextRange tr = new(MyDoc.ContentStart, MyDoc.ContentEnd);
                            byte[] byteArray = Encoding.ASCII.GetBytes(o.kifejtve);
                            MemoryStream stream = new(byteArray);
                            tr.Load(stream, DataFormats.Rtf);
                        }
                    }

                    if (!found)
                    {
                        TextRange tr = new(MyDoc.ContentStart, MyDoc.ContentEnd);
                        tr.Text = "";
                    }

                    return;
                }
            }
        }

        private void MenuItem_openfileClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new();
            openFileDialog.Filter = "Text file (*.txt)|*.txt";
            openFileDialog.InitialDirectory = Directory.GetCurrentDirectory().ToString();
            if (openFileDialog.ShowDialog() == true)
            {
                StringReader reader = new(File.ReadAllText(openFileDialog.FileName));
                int szCount = Int32.Parse(reader.ReadLine());
                for (int i = 0; i < szCount; i++)
                {
                    String fejlec = reader.ReadLine();
                    int cut = fejlec.IndexOf(":");
                    String szempontnev = fejlec.Substring(0, cut);
                    int nszempontok = Int32.Parse( fejlec[(cut + 1)..] );

                    int found = -1;
                    for(int j=0; j<szempontok.Count; j++)
                    {
                        if(szempontok[j].id == szempontnev)
                        {
                            found = j;
                        }
                    }
                    if (found == -1) continue;

                    szempontok[found].SzChecked(true);

                    for(int k=0; k<nszempontok; k++)
                    {
                        String opcionev = reader.ReadLine();

                        for(int j=0; j<szempontok[found].opciok.Count; j++)
                        {
                            if(szempontok[found].opciok[j].id == opcionev)
                            {
                                szempontok[found].opciok[j].cb.IsChecked = true;
                            }
                        }
                    }
                }
            }
        }

        private void MenuItem_savefileClick(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new();
            saveFileDialog.Filter = "Text file (*.txt)|*.txt";
            saveFileDialog.InitialDirectory = Directory.GetCurrentDirectory().ToString();
            if (saveFileDialog.ShowDialog() == true)
            {
                SaveTemplate(saveFileDialog.FileName);
            }
        }

        private void B_preview_Click(object sender, RoutedEventArgs e)
        {
            if (popup==null || popup.IsLoaded == false) popup = new();

            popup.parent = this;

            popup.PopupDoc.Blocks.Clear();

            foreach (Szempont sz in szempontok)
            {
                if (sz.isChecked)
                {
                    Paragraph szp = new();
                    szp.Inlines.Add(sz.id);
                    szp.FontWeight = FontWeights.Bold;
                    popup.PopupDoc.Blocks.Add(szp);
                }

                foreach (Opcio o in sz.opciok)
                {
                    if (o.cb.IsChecked == true)
                    {
                        Paragraph para = new();
                        popup.PopupDoc.Blocks.Add(para);
                        TextRange tr = new(para.ContentStart, para.ContentEnd);
                        byte[] byteArray = Encoding.ASCII.GetBytes(o.kifejtve);
                        MemoryStream stream = new(byteArray);
                        tr.Load(stream, DataFormats.Rtf);
                    }
                }

                Paragraph blank = new();
                popup.PopupDoc.Blocks.Add(blank);
            }

            popup.Show();
        }

        private void B_reload_Click(object sender, RoutedEventArgs e)
        {
            b_reload.IsEnabled = false;
            b_reload.Content = new Run("frissítés...");
            b_reload.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);

            Load();
            
            b_reload.IsEnabled = true;
            b_reload.Content = new Run("Frissítés");

            MessageBox.Show("Frissítve!");
        }

        public void SaveTemplate(String path)
        {
            String str = "";
            int ctr = 0;
            foreach (Szempont sz in szempontok)
            {
                if (sz.isChecked)
                {
                    ctr++;
                    String valasztott = "";
                    int nopciok = 0;
                    foreach (Opcio o in sz.opciok)
                    {
                        if (o.cb.IsChecked == true)
                        {
                            nopciok++;
                            valasztott += o.id + "\n";
                        }
                    }

                    str += sz.id + ":" + nopciok.ToString() + "\n" + valasztott;
                }
            }

            File.WriteAllText(path, ctr.ToString() + "\n" + str);
        }
    }
}
