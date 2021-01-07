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
        List<string> symbolListDownload;
        TradeStationWebApi TS;
        string optionsPath;
        public Form1()
        {
            InitializeComponent();
            cmbMode.SelectedIndex = 0;
            cmbUnit.SelectedIndex = 0;
            cmbC.SelectedIndex = 0;
            cmbSessionTemplate.SelectedIndex = 3;
            cmbDataType.SelectedIndex = 2;
            cmbStrikePriceType.SelectedIndex = 2;
            try
            {
                optionsPath = path + "\\Options.txt";
                var options = File.ReadAllLines(optionsPath);
                txtUri.Text = options[0];
                txtKey.Text = options[1];
                txtSecret.Text = options[2];
                cmbMode.SelectedIndex = int.Parse(options[3]);
                txtSybolList.Text = options[4];
                txtStockOutFolder.Text = options[5];
                txtOptionsSearchResultFolder.Text = options[6];
                txtOptionsOutputFolder.Text = options[7];
                cmbC.SelectedIndex = int.Parse(options[8]);
                txtStk.Text = options[9];
                txtSpl.Text = options[10];
                txtSph.Text = options[11];
                txtExd.Text = options[12];
                txtEdl.Text = options[13];
                txtEdh.Text = options[14];
                ot = options[15];
                if (ot == "Call")
                {
                    radioButton2.Checked = true;
                }
                if (ot == "Put")
                {
                    radioButton3.Checked = true;
                }
                datatype = options[16];
                if (datatype == "StockOption")
                {
                    radioButton4.Checked = true;
                }
                cmbUnit.SelectedIndex = int.Parse(options[17]);
                txtInterval.Text = options[18];
                txtBarsback.Text = options[19];
                if (options[20] == "True")
                {
                    checkBox1.Checked = true;
                }
                if (options[21] == "True")
                {
                    checkBox2.Checked = true;
                }
                dtpLastDate.Value = Convert.ToDateTime(options[22]);
                cmbSessionTemplate.SelectedIndex = int.Parse(options[23]);
                cmbDataType.SelectedIndex = int.Parse(options[24]);
                dtpStartDate.Value = Convert.ToDateTime(options[25]);
                dtpEndDate.Value = Convert.ToDateTime(options[26]);
                txtDaysBack.Text = options[27];
                if (options[28] == "True")
                {
                    ckbMinOfRows.Checked = true;
                }
                txtMinRows.Text = options[29];
                txtMinRows.Enabled = ckbMinOfRows.Checked;
                cmbExpireType.SelectedIndex = int.Parse(options[30]);
                cmbStrikePriceType.SelectedIndex = int.Parse(options[31]);
                txtDownloadSymbolList.Text = options[32];
                if (options[33] == "True")
                {
                    chkUpDwn.Checked = true;
                }


                var symbolfile = File.ReadAllLines(txtSybolList.Text);
                symbolsList = new List<string>(symbolfile);

                symbolfile = File.ReadAllLines(txtDownloadSymbolList.Text);
                symbolListDownload = new List<string>(symbolfile);

            }
            catch(Exception ex)
            {
                ErrorLog(ex);
            }
        }

        private void ErrorLog(Exception ex)
        {
            string logfile = path + "\\errorlog.txt";
            if (!File.Exists(logfile))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(logfile))
                {
                    sw.WriteLine(ex.ToString());
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(logfile))
                {
                    sw.WriteLine(ex.ToString());
                }
            }
        }

        // read symbol list
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
        // set stock output folder
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
        // set option search result
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
        // set option output  folder
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
        // search symbols
        private void button5_Click(object sender, EventArgs e)
        {
            if (InitialFunc() == false)
            {
                return;
            }
            if (symbolsList == null)
            {
                MessageBox.Show("Please check Stocks list file!");
                return;
            }
            progressBar1.Maximum = 100 * symbolsList.Count;
            progressBar1.Step = 100;
            progressBar1.Value = 0;

            var filename = txtOptionsSearchResultFolder.Text + "\\Options-Symbols-" + DateTime.Now.ToString("MM_dd_yy_hh_mm_ss") + ".txt";
            using (StreamWriter outputFile = new StreamWriter(filename))
            {
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

                        criteria += "&C=" + cmbC.SelectedItem.ToString();
                        //   "&c=StockOption&OT=Both&Stk=5&Exd=3";
                        var symbols = TS.SearchSymbols(criteria);
                        if (symbols == null)
                        {
                            continue;
                        }
                        foreach (var item in symbols)
                        {

                            if (string.IsNullOrEmpty(cmbExpireType.Text))
                            {
                                outputFile.WriteLine(item.Name);
                            }
                            else
                            {
                                if (item.ExpirationType == cmbExpireType.SelectedItem.ToString())
                                {
                                    outputFile.WriteLine(item.Name);
                                }
                            }
                                
                        }
                        progressBar1.PerformStep();
                    }
                }
            }
            MessageBox.Show("Symbol Search Download Finished");
            
        }

        private bool InitialFunc()
        {
            if (txtDownloadSymbolList.Text == "")
            {
                MessageBox.Show("Select Download Symbol List text file");
                return false;
            }
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
            try
            {
                TS = new TradeStationWebApi(
                txtKey.Text, // key
                txtSecret.Text, // secret
                host,
                "http://localhost:1125/");
                return true;
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
            }
            return false;
            
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
        string datatype = "Stock";
        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            datatype = "Stock";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            datatype = "StockOption";
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
        // download data
        private void button6_Click(object sender, EventArgs e)
        {
            if (InitialFunc() == false)
            {
                return;
            }
            progressBar1.Maximum = 100* symbolListDownload.Count;
            progressBar1.Step = 100;
            progressBar1.Value = 0;
            List<string> unsymbols = new List<string>();
            foreach (var symbol in symbolListDownload)
            {
                try
                {
                    if (symbol.Trim() != "")
                    {
                        IEnumerable<BarChart> barcharts = null;
                        // starting on date
                        if (cmbDataType.SelectedIndex == 0)
                        {
                            barcharts = TS.GetBarChartStartingOnDate(
                                symbol.Trim(),
                                txtInterval.Text,
                                cmbUnit.SelectedItem.ToString(),
                                dtpStartDate.Value.ToString("MM-dd-yyyy"),
                                cmbSessionTemplate.SelectedItem.ToString()
                            );
                        }
                        // date range
                        else if (cmbDataType.SelectedIndex == 1)
                        {
                            barcharts = TS.GetBarChartDateRange(
                                symbol.Trim(),
                                txtInterval.Text,
                                cmbUnit.SelectedItem.ToString(),
                                dtpStartDate.Value.ToString("MM-dd-yyyy"),
                                dtpEndDate.Value.ToString("MM-dd-yyyy"),
                                cmbSessionTemplate.SelectedItem.ToString()
                            );
                        }
                        // bars back
                        else if (cmbDataType.SelectedIndex == 2)
                        {
                            barcharts = TS.GetBarChart(
                                symbol.Trim(),
                                txtInterval.Text,
                                cmbUnit.SelectedItem.ToString(),
                                txtBarsback.Text,
                                dtpLastDate.Value.ToString("MM-dd-yyyy"),
                                cmbSessionTemplate.SelectedItem.ToString()
                            );
                        }
                        // days back
                        else if (cmbDataType.SelectedIndex == 3)
                        {
                            barcharts = TS.GetBarChartDaysBack(
                                symbol.Trim(),
                                txtInterval.Text,
                                cmbUnit.SelectedItem.ToString(),
                                cmbSessionTemplate.SelectedItem.ToString(),
                                txtDaysBack.Text,
                                dtpLastDate.Value.ToString("MM-dd-yyyy")
                            );
                        }


                        var minOfRows = int.Parse(txtMinRows.Text);
                        if (barcharts == null)
                        {
                            continue;
                        }

                        if (ckbMinOfRows.Checked)
                        {
                            if (barcharts.Count() < minOfRows)
                            {
                                unsymbols.Add(symbol);
                                continue;
                            }
                        }
                        
                        //before your loop
                        var csv = new StringBuilder();
                        if (chkUpDwn.Checked)
                        {
                            csv.AppendLine("Date,Time,Open,High,Low,Close,UpVol,DwVol");
                            foreach (var item in barcharts)
                            {
                                TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                DateTime easternTime = TimeZoneInfo.ConvertTimeFromUtc(item.TimeStamp, easternZone);
                                var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                    easternTime.ToString("MM/dd/yyyy"),
                                    easternTime.ToString("hh:mm:ss"),
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
                            if (checkBox1.Checked)
                            {
                                csv.AppendLine("Date,Time,Open,High,Low,Close,Volume");
                                var ordered = barcharts.OrderBy(b => b.TimeStamp);
                                foreach (var item in ordered)
                                {
                                    if (item.Close > 0)
                                    {
                                        TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                        DateTime easternTime = TimeZoneInfo.ConvertTimeFromUtc(item.TimeStamp, easternZone);
                                        var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6}",
                                            easternTime.ToString("MM/dd/yyyy"),
                                            easternTime.ToString("hh:mm:ss"),
                                            item.Open,
                                            item.High,
                                            item.Low,
                                            item.Close,
                                            item.UpVolume - item.DownVolume
                                            );
                                        csv.AppendLine(newLine);
                                    }

                                }
                            }
                            else
                            {
                                csv.AppendLine("Date,Time,Open,High,Low,Close,Volume");
                                foreach (var item in barcharts)
                                {
                                    if (item.Close > 0)
                                    {
                                        var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6}",
                                        item.TimeStamp.ToString("MM/dd/yyyy"),
                                        item.TimeStamp.ToString("hh:mm:ss"),
                                        item.Open,
                                        item.High,
                                        item.Low,
                                        item.Close,
                                        item.TotalVolume
                                        );
                                        csv.AppendLine(newLine);
                                    }
                                }
                            }
                        }
                        
                        if (checkBox2.Checked)
                        {
                            csv.AppendLine(",,,,,,,");
                        }
                        if (datatype == "Stock")
                        {
                            downloadFolder = txtStockOutFolder.Text;
                        }
                        else
                        {
                            downloadFolder = txtOptionsOutputFolder.Text;
                        }
                        //after your loop
                        File.WriteAllText(downloadFolder + "\\" + symbol + ".csv", csv.ToString());
                    }

                }
                catch (Exception ex)
                {
                    ErrorLog(ex);
                }

                progressBar1.PerformStep();
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
        string path = Directory.GetParent(Application.ExecutablePath).ToString();
        private void button7_Click(object sender, EventArgs e)
        {
            if (InitialFunc() == false)
            {
                return;
            }
            optionsPath = path + "\\Options.txt";
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(optionsPath))
            {
                file.WriteLine(txtUri.Text);
                file.WriteLine(txtKey.Text);
                file.WriteLine(txtSecret.Text);
                file.WriteLine(cmbMode.SelectedIndex);
                file.WriteLine(txtSybolList.Text);
                file.WriteLine(txtStockOutFolder.Text);
                file.WriteLine(txtOptionsSearchResultFolder.Text);
                file.WriteLine(txtOptionsOutputFolder.Text);
                file.WriteLine(cmbC.SelectedIndex);
                file.WriteLine(txtStk.Text);
                file.WriteLine(txtSpl.Text);
                file.WriteLine(txtSph.Text);
                file.WriteLine(txtExd.Text);
                file.WriteLine(txtEdl.Text);
                file.WriteLine(txtEdh.Text);
                file.WriteLine(ot);
                file.WriteLine(datatype);
                file.WriteLine(cmbUnit.SelectedIndex);
                file.WriteLine(txtInterval.Text);
                file.WriteLine(txtBarsback.Text);
                file.WriteLine(checkBox1.Checked);
                file.WriteLine(checkBox2.Checked);
                file.WriteLine(dtpLastDate.Value.ToShortDateString());
                file.WriteLine(cmbSessionTemplate.SelectedIndex);
                file.WriteLine(cmbDataType.SelectedIndex);
                file.WriteLine(dtpStartDate.Value.ToShortDateString());
                file.WriteLine(dtpEndDate.Value.ToShortDateString());
                file.WriteLine(txtDaysBack.Text);
                file.WriteLine(ckbMinOfRows.Checked);
                file.WriteLine(txtMinRows.Text);
                file.WriteLine(cmbExpireType.SelectedIndex);
                file.WriteLine(cmbStrikePriceType.SelectedIndex);
                file.WriteLine(txtDownloadSymbolList.Text);
                file.WriteLine(chkUpDwn.Checked);
            }
            MessageBox.Show("Saved Options");
        }

        private void ckbMinOfRows_CheckedChanged(object sender, EventArgs e)
        {
            txtMinRows.Enabled = ckbMinOfRows.Checked;
        }
        // get quotes
        private void button8_Click(object sender, EventArgs e)
        {
            if (InitialFunc() == false)
            {
                return;
            }
            if (symbolsList == null)
            {
                MessageBox.Show("Please check stocks list");
                return;
            }
            try
            {
                var symbols = string.Join(",", symbolsList);
                var quotes = TS.GetQuote(symbols);
                var csv = new StringBuilder();
                csv.AppendLine(
                    "Symbol," +
                    "Last"
                    );
                foreach (var item in quotes)
                {

                    if (item.Close > 0)
                    {
                        var newLine = string.Format("{0},{1}",
                        item.Symbol,
                        item.Last
                        );
                        csv.AppendLine(newLine);
                    }

                }
                File.WriteAllText(path + "\\Quotes.csv", csv.ToString());
                MessageBox.Show("Finished to get Quotes");
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
            }
        }
        // CleanUp Options List
        private void button9_Click(object sender, EventArgs e)
        {
            if (cmbStrikePriceType.SelectedIndex == 0)
            {

            }
            else if (cmbStrikePriceType.SelectedIndex == 1)
            {

            }
        }
        // download symbol list
        private void button10_Click(object sender, EventArgs e)
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
            txtDownloadSymbolList.Text = filePath;
            var symbolfile = File.ReadAllLines(txtDownloadSymbolList.Text);
            symbolListDownload = new List<string>(symbolfile);
        }
    }
}
