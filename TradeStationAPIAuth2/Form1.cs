using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace TradeStationAPIAuth2
{
    public partial class Form1 : Form
    {
        List<string> symbolsList;
        TradeStationWebApi TS;
        public Form1()
        {
            InitializeComponent();
            cmbMode.SelectedIndex = 0;
            cmbUnit.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
            txtSybolList.Text = filePath;
            var symbolfile = File.ReadAllLines(txtSybolList.Text);
            symbolsList = new List<string>(symbolfile);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtStockOutFolder.Text = folderDlg.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtOptionsSearchResultFolder.Text = folderDlg.SelectedPath;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtOptionsOutputFolder.Text = folderDlg.SelectedPath;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (InitialFunc() == false)
            {
                return;
            }
            foreach (var symbol in symbolsList)
            {
                if (symbol != "")
                {
                    var criteria = "r=" + symbol;
                    if (txtSpl.Text != "")
                    {
                        criteria += "&Spl=" + txtSpl.Text; 
                    }
                    if (txtSph.Text != "")
                    {
                        criteria += "&Sph=" + txtSph.Text;
                    }
                    criteria += "&Exd=" + txtExd.Text;
                    if (txtEdl.Text != "MM-DD-YYYY")
                    {
                        criteria += "&Edl=" + txtEdl.Text;
                    }
                    if (txtEdh.Text != "MM-DD-YYYY")
                    {
                        criteria += "&Edh=" + txtEdh.Text;
                    }
                    criteria += "&OT=" + ot;
                    criteria += "&ST=" + st;
                    criteria += "&N=" + symbol;

                    criteria += "&C=" + c;
                    if (c == "Stock")
                    {
                        downloadFolder = txtStockOutFolder.Text;
                    }
                    else
                    {
                        downloadFolder = txtOptionsOutputFolder.Text;
                    }
                    //   "&c=StockOption&OT=Both&Stk=5&Exd=3";
                    var symbols = TS.SearchSymbols(criteria);
                    var ser = new JavaScriptSerializer();
                    string jsonData = ser.Serialize(symbols);
                    //write string to file
                    System.IO.File.WriteAllText(downloadFolder + "\\Options-Symbols-" + DateTime.Now.ToString("MM_dd_yy_hh_mm_ss") + ".txt", jsonData);
                }
            }
            MessageBox.Show("Download Finished");
            
        }

        private bool InitialFunc()
        {
            if (txtKey.Text == "")
            {
                MessageBox.Show("Put api key");
                return false;
            }
            if (txtSecret.Text == "")
            {
                MessageBox.Show("Put api secret");
                return false;
            }
            if (txtSybolList.Text == "")
            {
                MessageBox.Show("Select Symbols List text file");
                return false;
            }
            if (txtStockOutFolder.Text == "")
            {
                MessageBox.Show("Select Stock Out Folder");
                return false;
            }
            if (txtOptionsOutputFolder.Text == "")
            {
                MessageBox.Show("Select StockOptions Out Folder");
                return false;
            }

            var mode = cmbMode.SelectedItem.ToString();
            var host = txtUri.Text;
            if (mode == "SIM") host = "https://sim.api.tradestation.com/v2";
            txtAuthUrl.Text = string.Format("{0}/{1}", host,
                    string.Format(
                        "authorize?client_id={0}&response_type=code&redirect_uri={1}",
                        txtKey.Text,
                        "http://localhost:1125/"));
            TS = new TradeStationWebApi(
                txtKey.Text, // key
                txtSecret.Text, // secret
                host,
                "http://localhost:1125/");
            return true;
        }

        string downloadFolder;
        string ot = "Both";
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            ot = "Both";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            ot = "Call";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            ot = "Put";
        }
        string c = "Stock";
        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            c = "Stock";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            c = "StockOption";
        }
        string st = "Composite";
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            st = "Both";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            st = "Composite";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (InitialFunc() == false)
            {
                return;
            }
            foreach (var symbol in symbolsList)
            {
                if (symbol != "")
                {
                    var barcharts = TS.GetBarChart(
                        symbol,
                        txtInterval.Text,
                        cmbUnit.SelectedItem.ToString(),
                        txtBarsback.Text,
                        dateTimePicker1.Value.ToString("MM-dd-yyyy")
                    );
                    //before your loop
                    var csv = new StringBuilder();
                    
                    if (checkBox1.Checked)
                    {
                        csv.AppendLine("Date,Time,Open,High,Low,Close,UpVol,DownVol");
                        foreach (var item in barcharts)
                        {
                            var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                item.TimeStamp.ToString("MM-dd-yyyy"), 
                                item.TimeStamp.ToString("hh-mm-ss"),
                                item.Open,
                                item.High,
                                item.Low,
                                item.Close,
                                item.UpVolume,
                                item.DownVolume
                                );
                            csv.AppendLine(newLine);
                        }
                    }
                    else
                    {
                        csv.AppendLine("Date,Time,Open,High,Low,Close,Volume");
                        foreach (var item in barcharts)
                        {
                            var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6}",
                                item.TimeStamp.ToString("MM-dd-yyyy"),
                                item.TimeStamp.ToString("hh-mm-ss"),
                                item.Open,
                                item.High,
                                item.Low,
                                item.Close,
                                item.TotalVolume
                                );
                            csv.AppendLine(newLine);
                        }
                    }
                    if (checkBox2.Checked)
                    {
                        csv.AppendLine(",,,,,,,");
                    }
                    //after your loop
                    File.WriteAllText(txtStockOutFolder.Text + "\\BarChart-" + DateTime.Now.ToString("MM_dd_yy_hh_mm_ss") + ".csv", csv.ToString());
                }
            }
            MessageBox.Show("Barchart download finished");
            
        }

        private void cmbUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbUnit.SelectedIndex > 0 )
            {
                txtInterval.Text = "1";
            }
        }

        private void txtInterval_TextChanged(object sender, EventArgs e)
        {
            if (txtInterval.Text == "") return;
            var interval = int.Parse(txtInterval.Text);
            if (cmbUnit.SelectedIndex == 0)
            {
                
                if (interval < 1 || interval > 1440)
                {
                    MessageBox.Show("Input Interval between 1 and 1440");
                    return;
                }
            }
            else
            {
                txtInterval.Text = "1";
            }
        }

        private void txtBarsback_TextChanged(object sender, EventArgs e)
        {
            if (txtBarsback.Text == "") return;
            var interval = int.Parse(txtBarsback.Text);
            if (interval < 1 || interval > 57600)
            {
                MessageBox.Show("Input BarsBack between 1 and 57600");
                return;
            }
        }
    }
}
