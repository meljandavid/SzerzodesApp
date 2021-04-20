using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace myWPF
{
    public partial class MainWindow : Window
    {
        List<Szempont> szempontok = new();
        PopupWindow popup;

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
                    o.rb.Click += CheckRadioOpcio_Click;
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
                rb = new RadioButton
                {
                    Content = _id,
                    FontSize = 14
                };
                Thickness margin = rb.Margin;
                margin.Left = 10;
                margin.Bottom = 2;
                rb.Margin = margin;
                id = _id;
            }

            public string id, kifejtve;
            public RadioButton rb;
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
            public void SzChecked()
            {
                rb.Foreground = new SolidColorBrush(Colors.Green);
                rb.FontWeight = FontWeights.Bold;
                rb.Background = new SolidColorBrush(Colors.LightGreen);
                isChecked = true;
            }
        }

        private void CheckRadioOpcio_Click(object sender, RoutedEventArgs e)
        {
            foreach (Szempont sz in szempontok)
            {
                if (sz.rb.IsChecked == true)
                {
                    foreach (Opcio o in sz.opciok)
                    {
                        if (o.rb.IsChecked == true)
                        {
                            TextRange tr = new(MyDoc.ContentStart, MyDoc.ContentEnd);
                            byte[] byteArray = Encoding.ASCII.GetBytes(o.kifejtve);
                            MemoryStream stream = new(byteArray);
                            tr.Load(stream, DataFormats.Rtf);

                            sz.SzChecked();
                        }
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
                        panel_opciok.Children.Add(o.rb);
                        if (o.rb.IsChecked == true)
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
                    String str = reader.ReadLine();
                    int cut = str.IndexOf(":");
                    String ssz = str.Substring(0, cut), so = str[(cut + 1)..];

                    foreach (Szempont sz in szempontok)
                    {
                        if (sz.id == ssz)
                        {
                            foreach (Opcio o in sz.opciok)
                            {
                                if (o.id == so)
                                {
                                    o.rb.IsChecked = true;
                                }
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
                String str = "";
                int ctr = 0;
                foreach (Szempont sz in szempontok)
                {
                    if (sz.isChecked)
                    {
                        String valasztott = "";

                        foreach (Opcio o in sz.opciok)
                        {
                            if (o.rb.IsChecked == true)
                            {
                                ctr++;
                                valasztott = o.id;
                                break;
                            }
                        }

                        str += sz.id + ":" + valasztott + "\n";
                    }
                }

                File.WriteAllText(saveFileDialog.FileName, ctr.ToString() + "\n" + str);
            }
        }

        private void B_preview_Click(object sender, RoutedEventArgs e)
        {
            if (popup==null || popup.IsLoaded == false) popup = new();

            popup.PopupDoc.Blocks.Clear();

            foreach (Szempont sz in szempontok)
            {
                foreach (Opcio o in sz.opciok)
                {
                    if (o.rb.IsChecked == true)
                    {
                        Paragraph szp = new();
                        szp.Inlines.Add(sz.id);
                        szp.FontWeight = FontWeights.Bold;
                        popup.PopupDoc.Blocks.Add(szp);

                        Paragraph para = new();
                        popup.PopupDoc.Blocks.Add(para);
                        TextRange tr = new(para.ContentStart, para.ContentEnd);
                        byte[] byteArray = Encoding.ASCII.GetBytes(o.kifejtve);
                        MemoryStream stream = new(byteArray);
                        tr.Load(stream, DataFormats.Rtf);
                    }
                }
            }

            popup.Show();
        }

        private void B_reload_Click(object sender, RoutedEventArgs e)
        {
            Load();
            MessageBox.Show("Frissítve!");
        }
    }
}
