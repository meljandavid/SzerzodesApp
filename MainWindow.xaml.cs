using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Windows.Media;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace myWPF
{
    public partial class MainWindow : Window
    {
        List<Szempont> szempontok = new();

        void Load()
        {
            panel_szempontok.Children.Clear();
            panel_opciok.Children.Clear();
            szempontok.Clear();
            Excel.Application xlApp = new();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Directory.GetCurrentDirectory().ToString()+"\\szerzodes_szoveg_sablon.xlsx");
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Excel.Range excelRange = xlWorksheet.UsedRange;

            //get an object array of all of the cells in the worksheet (their values)
            object[,] valueArray = (object[,])excelRange.get_Value(
                        Excel.XlRangeValueDataType.xlRangeValueDefault);

            //excel is not zero based!!
            for (int row = 2; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
            {
                String rsz = "";
                if(valueArray[row, 1] != null)
                {
                    rsz = valueArray[row, 1].ToString();
                }

                if (rsz != "")
                {
                    szempontok.Add(new Szempont(rsz));
                    szempontok[^1].opciok = new List<Opcio>();
                }
                Opcio opcio = new(valueArray[row, 2].ToString());
                String str = "";
                if (valueArray[row, 3] != null)
                {
                    str = valueArray[row, 3].ToString();
                }
                opcio.kifejtve = str;
                szempontok[^1].opciok.Add(opcio);
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

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //close and release
            xlWorkbook.Close();
        }

        public void openFile(String str)
        {

        }

        public MainWindow()
        {
            InitializeComponent();
            Load();
            /*
            String[] data = Environment.GetCommandLineArgs();
            szempontok.Add(new Szempont(data[0]));
            */
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
            public void szChecked()
            {
                rb.Foreground = new SolidColorBrush(Colors.Green);
                rb.FontWeight = FontWeights.Bold;
                rb.Background = new SolidColorBrush(Colors.LightGreen);
                isChecked = true;
            }
        }

        private void CheckRadioOpcio_Click(object sender, RoutedEventArgs e)
        {
            foreach(Szempont sz in szempontok)
            {
                if(sz.rb.IsChecked == true)
                {
                    foreach(Opcio o in sz.opciok)
                    {
                        if(o.rb.IsChecked == true)
                        {
                            mytextblock.Text = o.kifejtve;
                            sz.szChecked();
                        }
                    }
                }
            }
        }

        private void CheckRadioSzempont_Click(object sender, RoutedEventArgs e)
        {
            panel_opciok.Children.Clear();
            foreach(Szempont sz in szempontok)
            {
                if(sz.rb.IsChecked == true)
                {
                    foreach(Opcio o in sz.opciok)
                    {
                        panel_opciok.Children.Add(o.rb);
                        if(o.rb.IsChecked == true)
                        {
                            mytextblock.Text = o.kifejtve;
                        }
                    }
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
                /*
                for(int i=0; i<szCount; i++)
                {
                    szempontok[i].opciok[Int32.Parse(reader.ReadLine())].rb.IsChecked = true;
                }
                */
                for (int i = 0; i < szCount; i++)
                {
                    String str = reader.ReadLine();
                    int cut = str.IndexOf(":");
                    String ssz = str.Substring(0, cut), so=str.Substring(cut+1);

                    foreach(Szempont sz in szempontok)
                    {
                        if(sz.id == ssz)
                        {
                            foreach(Opcio o in sz.opciok)
                            {
                                if(o.id == so)
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
                /*
                String str = szempontok.Count.ToString() + "\n";
                foreach(Szempont sz in szempontok)
                {
                    int idx = 0;
                    for(int i=0; i<sz.opciok.Count;  i++)
                    {
                        if (sz.opciok[i].rb.IsChecked == true) idx = i;
                    }
                    str += idx.ToString()+ "\n" ;
                }*/
                String str = "";
                int ctr = 0;
                foreach(Szempont sz in szempontok)
                {
                    if(sz.isChecked)
                    {
                        String valasztott = "";

                        foreach(Opcio o in sz.opciok)
                        {
                            if(o.rb.IsChecked == true)
                            {
                                ctr++;
                                valasztott = o.id;
                                break;
                            }
                        }

                        str += sz.id + ":" + valasztott + "\n";
                    }
                }

                File.WriteAllText(saveFileDialog.FileName, ctr.ToString()+"\n"+str);
            }
        }

        private void b_copy_Click(object sender, RoutedEventArgs e)
        {
            String res = "";

            foreach (Szempont sz in szempontok)
            {
                int idx = -1;
                for (int i = 0; i < sz.opciok.Count; i++)
                {
                    if (sz.opciok[i].rb.IsChecked == true)
                        idx = i;
                }
                if(idx != -1)
                    res += sz.opciok[idx].kifejtve + "\n";
            }

            PopupWindow popup = new();
            popup.Show();
            popup.popupTextbox.Text = res;
        }

        private void b_reload_Click(object sender, RoutedEventArgs e)
        {
            Load();
        }
    }
}
