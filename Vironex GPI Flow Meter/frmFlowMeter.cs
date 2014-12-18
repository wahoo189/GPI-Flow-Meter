using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NDde.Client;
//using Microsoft.Office.Interop.Word;
using System.IO;
using System.Xml;
using DevExpress.XtraCharts;
using DevExpress.XtraCharts.Wizard;
using DevExpress.XtraBars;
//using System.Data.SqlServerCe;
using System.Configuration;
using System.Threading;
using Microsoft.Office.Interop;
using DevExpress.XtraReports.UI;
using System.IO.Ports;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;

namespace Vironex_GPI_Flow_Meter
{
    public partial class Main : Form
    {
        private DdeClient _FlowMeterClient;               // This is used to send DDE Commands to WinWedge
        private DdeClient _PressureSwitchClient;
        private DdeClient _DHPressureSwitchClient;
        //StreamWriter _outfile;                   // This is the output file created when data is captured.
        XmlTextWriter _XMLwriter = new XmlTextWriter(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\temp_file.xml", null);
        private const int _timeout = 60000;     // Timeout for requests
        private bool _blnDataCollectionEnabled;
        //FlowMeterReadings readings = new FlowMeterReadings();
        ChartWizard _whiz;
        private bool _doZeroPressure = true;
        System.Data.DataTable _dtChart = new System.Data.DataTable("injection");
        System.Data.DataTable _dtPrintChart = new System.Data.DataTable("PrintInjection");
        //bool _demoMode = true;
        private int _currentRow = 1;
        private string _Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\WinWedge.exe";
        private string _File = "";
        private Microsoft.Office.Interop.Excel.Application _app = new Microsoft.Office.Interop.Excel.Application();
        public Microsoft.Office.Interop.Excel.Workbook _workbook;// = _app.Workbooks.Add(1);
        public Microsoft.Office.Interop.Excel.Worksheet _worksheet;// = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
        private Microsoft.Office.Interop.Excel.Range _range;
        private Microsoft.Office.Interop.Excel.Range _rangePrint;
        private TextAnnotation currentAnnotation = null;
        int chartReadingsAnnotationNum = 0;
        int chart1AnnotationNum = 0;

        public Main()
        {
            InitializeComponent();

            _whiz = new ChartWizard(chartReadings);

            CreateTables();
            InitChart();
            CreatePrintTables();
            InitPrintChart();
            buttonStart.Enabled = false;
            btnSetVoltage.Enabled = false;
            comboBox1.Enabled = false;
            string[] ports = SerialPort.GetPortNames();

            try
            {
                if (ports[0] != "")
                {
                    Array.Sort(ports);
                    //Console.WriteLine("The following serial ports were found:");
                    cbFlowMeterCOMPort.Text = ports[0];
                    cbPressureSwitchCOMPort.Text = ports[2];
                    cbDHPressureSwitchCOMPort.Text = ports[2];
                }
            }
            catch
            {
                MessageBox.Show("Com Ports not hooked up");
            }
        }

        private void CreateTables()
        {
            DataColumn column;
            //DataRow row;
            //Data table for chart:
            for (int i = 1; i < 4; i++)
            {
                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Decimal");
                column.ColumnName = "Field" + i.ToString();
                // Add the Column to the DataColumnCollection.
                _dtChart.Columns.Add(column);
            }
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "DateTime";
            // Add the Column to the DataColumnCollection.
            _dtChart.Columns.Add(column);
        }

        private void InitChart()
        {
            chartReadings.Series.Clear();
            chartReadings.DataSource = _dtChart;

            DevExpress.XtraCharts.Series series1;
            string seriesName = "";
            for (int i = 1; i < 4; i++)
            {
                if (i == 1) { seriesName = "Flow"; }
                if (i == 2) { seriesName = "Uphole Pressure"; }
                if (i == 3) { seriesName = "Downhole Pressure"; }
                series1 = new DevExpress.XtraCharts.Series();
                series1.Name = seriesName;
                series1.ArgumentDataMember = "DateTime";
                series1.ValueDataMembers[0] = "Field" + i.ToString();
                series1.View = new LineSeriesView();
                LineSeriesView lsv = (LineSeriesView)series1.View;
                //lsv.LineMarkerOptions.Visible = false;
                chartReadings.Series.Add(series1);
            }

            ((LineSeriesView)chartReadings.Series[0].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
            ((LineSeriesView)chartReadings.Series[1].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
            ((LineSeriesView)chartReadings.Series[2].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;

            ((LineSeriesView)chartReadings.Series[0].View).LineMarkerOptions.Size = 1;
            ((LineSeriesView)chartReadings.Series[1].View).LineMarkerOptions.Size = 1;
            ((LineSeriesView)chartReadings.Series[2].View).LineMarkerOptions.Size = 1;
        }

        private void CreatePrintTables()
        {
            DataColumn ColumnPrint;

            for (int i = 1; i < 4; i++)
            {
                ColumnPrint = new DataColumn();
                ColumnPrint.DataType = System.Type.GetType("System.Decimal");
                ColumnPrint.ColumnName = "Field" + i.ToString();
                _dtPrintChart.Columns.Add(ColumnPrint);
            }
            ColumnPrint = new DataColumn();
            ColumnPrint.DataType = System.Type.GetType("System.String");
            ColumnPrint.ColumnName = "DateTime";
            // Add the Column to the DataColumnCollection.
            _dtPrintChart.Columns.Add(ColumnPrint);
        }
        private void InitPrintChart()
        {
            chartControl1.Series.Clear();
            chartControl1.DataSource = _dtPrintChart;
            DevExpress.XtraCharts.Series series1;
            string seriesPrintName = "";
            series1 = null;

            for (int i = 1; i < 4; i++)
            {
                if (i == 1) { seriesPrintName = "Flow"; }
                if (i == 2) { seriesPrintName = "Uphole Pressure"; }
                if (i == 3) { seriesPrintName = "Downhole Pressure"; }
                series1 = new DevExpress.XtraCharts.Series();
                series1.Name = seriesPrintName;
                series1.ArgumentDataMember = "DateTime";
                series1.ValueDataMembers[0] = "Field" + i.ToString();
                series1.View = new LineSeriesView();
                LineSeriesView lsv = (LineSeriesView)series1.View;
                chartControl1.Series.Add(series1);
            }

            ((LineSeriesView)chartControl1.Series[0].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
            ((LineSeriesView)chartControl1.Series[1].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
            ((LineSeriesView)chartControl1.Series[2].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;

            ((LineSeriesView)chartControl1.Series[0].View).LineMarkerOptions.Size = 1;
            ((LineSeriesView)chartControl1.Series[1].View).LineMarkerOptions.Size = 1;
            ((LineSeriesView)chartControl1.Series[2].View).LineMarkerOptions.Size = 1;

        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Process p in System.Diagnostics.Process.GetProcessesByName("WinWedge"))
            {
                p.Kill();
                p.WaitForExit();
            }
        }

        private void StartCollecting()
        {
            //if (_blnDataCollectionEnabled) { return; }
            //_blnDataCollectionEnabled = true;

            try
            {
                _FlowMeterClient = new DdeClient("WinWedge", cbFlowMeterCOMPort.Text.ToString());
                _PressureSwitchClient = new DdeClient("WinWedge", cbPressureSwitchCOMPort.Text.ToString());
                _DHPressureSwitchClient = new DdeClient("WinWedge", cbDHPressureSwitchCOMPort.Text.ToString());

                _FlowMeterClient.Connect();
                _PressureSwitchClient.Connect();
                _DHPressureSwitchClient.Connect();
                _FlowMeterClient.Execute("[RESET]", 60000);
                _PressureSwitchClient.Execute("[RESET]", 60000);
                _DHPressureSwitchClient.Execute("[RESET]", 60000);
                InitChart();
            }
            catch (Exception e)
            {
                if (!cbDemoMode.Checked)
                {
                    MessageBox.Show("WinWedge is not running, cannot collect data. Start WinWedge on the Settings tab.", "Start WinWedge", MessageBoxButtons.OK);
                    return;
                }
            }

            int i = 1;
            while (_blnDataCollectionEnabled)
            {
                //timer.Start();
                //if (timer.e)
                //{ }
                getData(_currentRow);
                chartReadings.RefreshData();

                if (rangeControlReset.Checked)
                {
                    //rangeControl1.SelectedRange.Maximum =
                    AxisX axis = ((XYDiagram)chartReadings.Diagram).AxisX;
                    axis.VisualRange.MaxValue = axis.WholeRange.MaxValue;
                    axis.VisualRange.MinValue = axis.WholeRange.MinValue;

                    //rangeControl1.SelectedRange = new DevExpress.XtraEditors.RangeControlRange(axis.VisualRange.MinValueInternal, axis.VisualRange.MaxValueInternal);
                    //rangeControl1.SelectedRange.Reset();
                    //rangeControl1.Refresh(); 
                }
                _currentRow++;
                if (i % 5 == 0)
                {
                    try
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch
                    {
                        //MessageBox.Show("An Error occured");
                    }
                }
                i++;
            }
        }

        private void getData(int rowNum)
        {
            var vDat = "";

            if (!_blnDataCollectionEnabled) { return; }
            if (!cbDemoMode.Checked)
            {
                if (!_FlowMeterClient.IsConnected) { _FlowMeterClient.Connect(); }
                if (!_PressureSwitchClient.IsConnected) { _PressureSwitchClient.Connect(); }
                if (!_DHPressureSwitchClient.IsConnected) { _DHPressureSwitchClient.Connect(); }


                SendSensorCommand();
            }
            Reading r = new Reading();
            Random rnd = new Random();
            DataRow row;
            row = _dtChart.NewRow();

            for (int i = 1; i < 12; i++)
            {
                if (!cbDemoMode.Checked)
                {
                    try
                    {
                        vDat = _FlowMeterClient.Request("Field(" + i + ")", _timeout);
                        if (vDat == "\0") { vDat = "0"; }
                        vDat = vDat.Replace("\0", "");
                        _range = (Microsoft.Office.Interop.Excel.Range)_worksheet.Cells[rowNum + 1, i];
                        if (vDat == "0")
                            _range.Value = "0";
                        else
                            _range.Value = vDat;


                        if (i == 4)
                        {
                            vDat = vDat.TrimStart('0');
                            txtTotalGallons.Text = vDat.ToString();
                        }

                        if (i == 9)
                        {
                            decimal tmp;
                            decimal.TryParse(vDat, out tmp); ;
                            tmp = tmp / 100;
                            row[0] = tmp;
                            txtFlowGPM.Text = tmp.ToString();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Open task manager and kill Excel");
                    }
                }
                else
                {
                    vDat = rnd.Next(1, 100).ToString();
                    _worksheet.Cells[rowNum + 1, i] = vDat;
                    if (i == 9)
                    {
                        row[0] = vDat;
                    }
                }
            }

            for (int i = 1; i < 2; i++)
            {
                double voltageBaseup = new double();
                double voltageBaseupAdj = new double();
                voltageBaseup = Convert.ToDouble(UpHoleVoltage.Text);
                voltageBaseupAdj = 300 / (20 - voltageBaseup);

                if (!cbDemoMode.Checked)
                {
                    vDat = _DHPressureSwitchClient.Request("Field(1)", _timeout);
                    if (vDat == "\0")
                    { vDat = "0"; }
                    vDat = vDat.Replace("*", "");
                    vDat = vDat.Replace("-", "");
                    vDat = vDat.Replace("\0", "");
                    vDat = vDat.Replace("+", "");

                    if (i == 1)
                    {
                        double tempPress = 0;
                        //tempPress = Convert.ToDouble(vDat);
                        tempPress = (Convert.ToDouble(vDat) - voltageBaseup) * voltageBaseupAdj;
                        //tempPress = tempPress * .05066;
                        // = tempPress - Convert.ToDouble(txtAboveGroundPressure.Text);
                        _worksheet.Cells[rowNum + 1, i + 11] = tempPress;
                        //vDat = Convert.ToString(tempPress);

                        row[1] = tempPress;
                        txtPressurePSI.Text = tempPress.ToString().TrimStart('0');
                    }
                }
                else
                {
                    vDat = rnd.Next(1, 100).ToString();
                    _worksheet.Cells[rowNum + 1, i + 12] = vDat;
                    if (i == 1)
                    {
                        row[1] = vDat;
                        txtPressurePSI.Text = vDat.ToString();
                    }
                }
            }

            double decTmp;
            double decTemp2;
            double dhPress;
            double decTmp3;
            double.TryParse(txtPressMinusVoltage.Text, out decTemp2);

            double voltageBasedn = new double();
            double voltageBasednAdj = new double();
            double pressureBasedn = new double();
            voltageBasedn = Convert.ToDouble(DNHoleVoltage.Text);
            pressureBasedn = Convert.ToDouble(txtBelowGroundPressure.Text);
            voltageBasednAdj = 500 / (20 - voltageBasedn);

            for (int i = 1; i == 1; i++)
            {
                if (!cbDemoMode.Checked)
                {
                    vDat = _DHPressureSwitchClient.Request("Field(2)", _timeout);

                    vDat = vDat.Replace("+", "");
                    vDat = vDat.Replace("-", "");
                    vDat = vDat.Replace("*", "");
                    vDat = vDat.Replace("\0", "");
                    vDat = vDat.TrimStart('0');
                    double.TryParse(vDat, out  decTmp);
                    if (vDat == "" || vDat == "0")
                    {
                        vDat = "0";
                        decTmp = 0;
                    }
                    else
                    {

                        decTmp = (Convert.ToDouble(vDat) - voltageBasedn) * voltageBasednAdj;
                        //decTmp = decTmp - pressureBasedn;

                    }

                    row[2] = decTmp;
                    txtDHPressurePSI.Text = Math.Round(Convert.ToDecimal(decTmp), 2).ToString();
                    _worksheet.Cells[rowNum + 1, i + 15] = decTmp;
                    txtDHPressurePSI.Text = Math.Round(Convert.ToDecimal(decTmp), 2).ToString();
                }
                else
                {
                    vDat = rnd.Next(1, 100).ToString();
                    _worksheet.Cells[rowNum + 1, i + 16] = vDat;
                    if (i == 1)
                    {
                        row[2] = vDat;
                        txtDHPressurePSI.Text = vDat.ToString();
                    }
                }
            }

            row["DateTime"] = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssTZD");
            _worksheet.Cells[rowNum + 1, 20] = row["DateTime"];
            _dtChart.Rows.Add(row);

            if (Convert.ToDouble(txtReadingSleep.Text) < 3)
            {
                MessageBox.Show("Must be Greater than 3");
            }
            else
            {
                double sleeping;
                sleeping = Convert.ToDouble(txtReadingSleep.Text) * 1000;
                //System.Threading.Tasks.Task.Delay(1000);

            }
        }

        private void SendSensorCommand()
        {
            if (!cbDemoMode.Checked)
            {
                //MRB: Need to know what the SENDOUT commands do.
                if (!_FlowMeterClient.IsConnected) { _FlowMeterClient.Connect(); }
                if (!_PressureSwitchClient.IsConnected) { _PressureSwitchClient.Connect(); }
                if (!_DHPressureSwitchClient.IsConnected) { _DHPressureSwitchClient.Connect(); }
                //_FlowMeterClient.Execute("[SENDOUT('$1RB',13,10)]", _timeout);
                //_PressureSwitchClient.Execute("[SENDOUT('$1RB',13,10)]", _timeout);
                _DHPressureSwitchClient.Execute("[SENDOUT('$1RB',13,10)]", _timeout);

            }
        }

        private void DurationOver()
        {
            _blnDataCollectionEnabled = false;
        }

        private void StopDataCollection()
        {
            _blnDataCollectionEnabled = false;
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            if (_workbook != null)
            {
                _worksheet = (Microsoft.Office.Interop.Excel.Worksheet)_workbook.Sheets.Add();
                _worksheet.Name = txtBoringName.Text + " - " + txtBeginDepth.Text + txtEndDepth.Text;
                createSheetHeaders();

                buttonPause.Enabled = true;
                buttonStop.Enabled = true;
                btnFinish.Enabled = true;
                buttonStart.Enabled = false;
                _blnDataCollectionEnabled = true;
                Thread.Sleep(3000);
                StartCollecting();
            }
            else
            { MessageBox.Show("No spreadsheet file selected."); }
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            StopDataCollection();
            buttonPause.Enabled = false;
            buttonStop.Enabled = false;
            buttonStart.Enabled = true;
            txtReadingSleep.Enabled = true;
            _currentRow = 1;
            _dtChart.Rows.Clear();
            chartReadings.RefreshData();

            var messageResult = MessageBox.Show("Do you want to add another injection point", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (messageResult == DialogResult.Yes)
            {
                txtBoringName.Text = "";
                txtBeginDepth.Text = "";
                txtEndDepth.Text = "";
                UpHoleVoltage.Enabled = false;
                DNHoleVoltage.Enabled = false;
                tabMain.SelectedIndex = 1;
                MessageBox.Show("To collect another boring Go to settings page:" + Environment.NewLine +
                    "- Enter new Boring Name" + Environment.NewLine +
                    "- Enter new Start and End depth", "Start A New Depth/Boring" + Environment.NewLine +
                    "Turn off pump, open all valves" + Environment.NewLine + "Allow Pressure transducers to equilibrate"
                    + Environment.NewLine + "and zero out pressure switches"
                    , MessageBoxButtons.OK);
            }
            if (messageResult == DialogResult.No)
            {
                //Main_FormClosing();
            }


        }

        private void btnShowChartWiz_Click(object sender, EventArgs e)
        {
            _whiz.ShowDialog();
        }

        private void btnResetChart_Click(object sender, EventArgs e)
        {
            _dtChart.Rows.Clear();
            chartReadings.RefreshData();
        }

        private void btnOpenWinWedge_Click(object sender, EventArgs e)
        {
            if ((cbFlowMeterCOMPort.Text == "") || (cbPressureSwitchCOMPort.Text == "") || (cbDHPressureSwitchCOMPort.Text == ""))
            {
                MessageBox.Show("You must select all COM ports");
            }
            else
            {
                //string temp = "";
                string COMPort = "";
                for (int i = 0; i < 3; i++)
                {
                    StringBuilder newFile = new StringBuilder();
                    switch (i)
                    {
                        case 0:
                            _File = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + Properties.Settings.Default.flowMeterSW3;
                            COMPort = cbFlowMeterCOMPort.Text;
                            char[] ZeroChar = { 'C', 'O', 'M' };
                            COMPort = COMPort.TrimStart(ZeroChar);
                            break;
                        case 1:
                            _File = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + Properties.Settings.Default.downHolePressureSwitchSW3;
                            COMPort = cbPressureSwitchCOMPort.Text;
                            char[] OneChar = { 'C', 'O', 'M' };
                            COMPort = COMPort.TrimStart(OneChar);
                            break;
                        case 2:
                            _File = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + Properties.Settings.Default.downHolePressureSwitchSW3;
                            COMPort = cbDHPressureSwitchCOMPort.Text;
                            char[] TwoChar = { 'C', 'O', 'M' };
                            COMPort = COMPort.TrimStart(TwoChar);
                            break;
                    }

                    string[] file = File.ReadAllLines(@_File);
                    int j = 0;

                    foreach (string line in file)
                    {
                        if (j == 1)
                        {
                            newFile.Append(" " + (Convert.ToInt16(COMPort) - 1) + " " + "\r\n");
                            //continue;
                        }
                        else
                        {
                            newFile.Append(line + "\r\n");
                        }
                        j++;
                    }

                    if (i < 2)
                    {
                        File.WriteAllText(@_File, newFile.ToString());
                        // Use ProcessStartInfo class
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.CreateNoWindow = false;
                        startInfo.UseShellExecute = false;
                        startInfo.FileName = _Path;
                        startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                        startInfo.Arguments = _File;
                        btnSetVoltage.Enabled = true;
                        Process.Start(startInfo);
                    }

                    btnVoltageStart.Enabled = true;
                    btnOpenWinWedge.Enabled = false;
                }

            }
        }

        private void btnSelectExcelSheet_Click(object sender, EventArgs e)
        {
            OFD.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
        }

        private void createSheetHeaders()
        {
            _worksheet.Cells[1, 1] = "Flow Meter";
            _worksheet.Cells[1, 12] = "Pressure Switch";
            _worksheet.Cells[1, 16] = "Down Hole Pressure Switch";
            _worksheet.Cells[1, 20] = "Date Time";
            _worksheet.Cells[1, 21] = "Notes Time";
            _worksheet.Cells[1, 21] = "Notes";
        }

        private bool createExcelWorkbook(string fileName = "")
        {
            _app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;
            _app.Visible = true;
            _workbook = _app.Workbooks.Add(1);
            _worksheet = (Microsoft.Office.Interop.Excel.Worksheet)_workbook.Sheets[1];
            _worksheet.Name = txtBoringName.Text + " - " + txtBeginDepth.Text + txtEndDepth.Text;

            createSheetHeaders();

            _workbook.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return true;
        }

        private void buttonPause_Click(object sender, EventArgs e)
        {
            if (buttonPause.Text == "Suspend")
            {
                buttonPause.Text = "Resume";
                _blnDataCollectionEnabled = false;
                txtReadingSleep.Enabled = true;
                MessageBox.Show("PAUSED", "The DATA IS PAUSED", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                buttonPause.Text = "Suspend";
                _blnDataCollectionEnabled = true;
                txtReadingSleep.Enabled = false;
                StartCollecting();
            }
        }

        private void btnStopVoltage_Click(object sender, EventArgs e)
        {
            _doZeroPressure = false;
            btnVoltageStart.Enabled = true;
            btnStopVoltage.Enabled = false;
            buttonStart.Enabled = true;
        }

        private void btnVoltageStart_Click(object sender, EventArgs e)
        {


            var tmpUP = "";
            var tmpDN = "";
            double tmpAbove1, tmpAbove2, tmpAbove3, tmpBelow1, tmpBelow2, tmpBelow3, tmp3, voltageBaseup, voltageBasedn, voltageBaseupAdj, voltageBasednAdj;
            if (UpHoleVoltage.Text == "" && DNHoleVoltage.Text == "")
            {
                MessageBox.Show("Set Voltage First");
                btnVoltageStart.Enabled = false;
            }

            else
            {
                voltageBaseup = Convert.ToDouble(UpHoleVoltage.Text);
                voltageBasedn = Convert.ToDouble(DNHoleVoltage.Text);

                voltageBaseupAdj = 300 / (20 - voltageBaseup);
                voltageBasednAdj = 500 / (20 - voltageBasedn);

                btnVoltageStart.Enabled = false;
                btnStopVoltage.Enabled = true;
                buttonStart.Enabled = false;

                _DHPressureSwitchClient = new DdeClient("WinWedge", cbDHPressureSwitchCOMPort.Text.ToString());
                _DHPressureSwitchClient.Connect();
                _DHPressureSwitchClient.Execute("[RESET]", 60000);

                //var vDat = "";

                if (!_DHPressureSwitchClient.IsConnected) { _DHPressureSwitchClient.Connect(); }
                _doZeroPressure = true;

                while (_doZeroPressure)
                {
                    _DHPressureSwitchClient.Execute("[SENDOUT('$1RB',13,10)]", _timeout);
                    tmpUP = _DHPressureSwitchClient.Request("Field(1)", _timeout);
                    tmpDN = _DHPressureSwitchClient.Request("Field(2)", _timeout);

                    tmpUP = tmpUP.Replace("+", "");
                    tmpUP = tmpUP.Replace("*", "");
                    tmpUP = tmpUP.Replace("\0", "");
                    tmpUP = tmpUP.TrimStart('0');

                    tmpDN = tmpDN.Replace("+", "");
                    tmpDN = tmpDN.Replace("*", "");
                    tmpDN = tmpDN.Replace("\0", "");
                    tmpDN = tmpDN.TrimStart('0');


                    if (tmpUP == "" || tmpUP == "00.0000")
                    {
                        tmpUP = "0";
                    }

                    else
                    {

                        //tmp3 = (Convert.ToDouble(tmpUP) - 4.8) * 26.6455696;
                        tmp3 = (Convert.ToDouble(tmpUP) - voltageBaseup) * voltageBaseupAdj;
                        tmpUP = Convert.ToString(tmp3);
                    }

                    if (tmpDN == "" || tmpDN == "00.0000")
                    {
                        tmpDN = "0";
                    }

                    else
                    {
                        //tmp3 = (Convert.ToDouble(tmpDN) - 4.46) * 26.6455696;
                        tmp3 = (Convert.ToDouble(tmpDN) - voltageBasedn) * voltageBasednAdj;
                        tmpDN = Convert.ToString(tmp3);
                    }

                    txtVoltageAbove.Text = tmpUP;
                    txtVoltage.Text = tmpDN;

                    double.TryParse(txtBelowGroundPressure.Text, out tmpBelow1);
                    double.TryParse(txtVoltage.Text, out tmpBelow2);
                    tmpBelow3 = tmpBelow1 - tmpBelow2;

                    double.TryParse(txtAboveGroundPressure.Text, out tmpAbove1);
                    double.TryParse(txtVoltageAbove.Text, out tmpAbove2);
                    tmpAbove3 = tmpAbove1 - tmpAbove2;

                    txtPressMinusVoltageAbove.Text = Convert.ToString(tmpAbove3);
                    txtPressMinusVoltage.Text = Convert.ToString(tmpBelow3);

                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(1000);
                }
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            FlowReport report = new FlowReport();
            Stream chartLayout = new MemoryStream();
            chartReadings.SaveToStream(chartLayout);
            XRChart xrChart = new XRChart();

            chartLayout.Seek(0, System.IO.SeekOrigin.Begin);
            xrChart.LoadFromStream(chartLayout);
            xrChart.DataSource = chartReadings.DataSource;
            float height = report.PageHeight - report.Margins.Top - report.Margins.Bottom;
            //height = height - report.Bands.GetBandByType(typeof(PageHeaderBand)).Height;
            height = height - report.Bands.GetBandByType(typeof(PageFooterBand)).HeightF;
            height = height - 1;
            int width = report.PageWidth - report.Margins.Left - report.Margins.Right;
            xrChart.HeightF = height;
            xrChart.WidthF = width;
            //xrChart.HeightF = height;
            //xrChart.WidthF = width;
            //xrChart.TopF = chartY;

            report.Bands.GetBandByType(typeof(DetailBand)).Controls.Add(xrChart);

            report.xrlClientCompany.Text = txtCompany.Text;
            report.xlrVisionJobNo.Text = txtVisionJob.Text;
            report.xlrSiteAddress.Text = textBoxProjectAddress.Text;
            report.xlrBoringName.Text = txtBoringName.Text;
            report.xlrBeginDepth.Text = txtBeginDepth.Text;
            report.xlrEndDepth.Text = txtEndDepth.Text;

            this.Cursor = Cursors.Default;
            report.ShowPreviewDialog();
        }

        private void btnNotes_Click(object sender, EventArgs e)
        {
            lblMainTime.Visible = true;
            lblMainTimeActual.Visible = true;
            txtMainTime.Visible = true;
            btnMainTime.Visible = true;
            lblMainTimeActual.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
            txtMainTime.Text = "INSERT NOTE HERE BONEHEAD";
        }

        private void btnMainTime_Click(object sender, EventArgs e)
        {
            int nextnote = new int();
            string time = lblMainTimeActual.Text;
            string notes = txtMainTime.Text;
            nextnote = _currentRow;
            _worksheet.Cells[nextnote, 21] = time;
            _worksheet.Cells[nextnote, 22] = notes;
            lblMainTime.Visible = false;
            lblMainTimeActual.Visible = false;
            txtMainTime.Visible = false;
            btnMainTime.Visible = false;
            MessageBox.Show("Note Collected");
        }

        private void cbFlowMeterCOMPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbFlowMeterCOMPort.Items.Clear();
            cbPressureSwitchCOMPort.Items.Clear();
            cbDHPressureSwitchCOMPort.Items.Clear();

            // Get a list of serial port names. 
            string[] ports = SerialPort.GetPortNames();

            Array.Sort(ports);
            Console.WriteLine("The following serial ports were found:");
            //cbFlowMeterCOMPort.Text = ports[0];
            //cbPressureSwitchCOMPort.Text = ports[2];
            //cbDHPressureSwitchCOMPort.Text = ports[1];
            //Display each port name to the console. 
            foreach (string port in ports)
            {
                cbFlowMeterCOMPort.Items.Add(port);
                cbPressureSwitchCOMPort.Items.Add(port);
                cbDHPressureSwitchCOMPort.Items.Add(port);
            }

            Console.ReadLine();
        }

        private void txtBelowGroundPressure_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.txtBelowGroundPressure, "Insert base voltage of DownHole pressure transducer (norm should be 890)");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] sheetName = new string[100];
            int sheetCount;
            sheetCount = _workbook.Sheets.Count;
            string[] sheetcountname = new string[0] { };
            for (int i = 0; i < sheetCount; i++)
            {
                sheetName[i] = _workbook.Name[i].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool blnContinue = false;
            txtReadingSleep.Enabled = false;
            OFD.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            DialogResult Result = OFD.ShowDialog();
            //flLabel.Text = OFD.FileName;

            if (Result == DialogResult.OK)
                if (OFD.CheckFileExists == false)
                {
                    createExcelWorkbook();

                }
            if (OFD.CheckFileExists == true)
            {
                _workbook = _app.Workbooks.Open(OFD.FileName,
                    0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
                _worksheet = (Microsoft.Office.Interop.Excel.Worksheet)_workbook.Sheets.Add();
                _worksheet.Name = "ING" + txtBoringName.Text + " - " + txtBeginDepth.Text + txtEndDepth.Text;
                blnContinue = true;

            }
        }

        private void flLabel_TextChanged(object sender, EventArgs e)
        {
            buttonStart.Enabled = true;
        }

        private void rangeControlReset_Click(object sender, EventArgs e)
        {
            //rangeControl1.SelectedRange.Minimum = 1;
            //rangeControl1.SelectedRange.Maximum = 100000;
            rangeControl1.SelectedRange.Minimum.Equals(0);
            rangeControl1.SelectedRange.Maximum.Equals(10000);
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            txtReadingSleep.Enabled = false;
            OFD.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            DialogResult Result = OFD.ShowDialog();
            comboBox1.Enabled = true;
            if (Result == DialogResult.OK)
            {

                if (File.Exists(OFD.FileName))
                {
                    _app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;
                    _app.Visible = true;
                    _workbook = _app.Workbooks.Open(OFD.FileName,
                        0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                        true, false, 0, true, false, false);
                    //_worksheet = (Microsoft.Office.Interop.Excel.Worksheet)_workbook.Sheets.Add();
                    //_worksheet.Name =  txtBoringName.Text + " - " + txtBeginDepth.Text + txtEndDepth.Text;
                    //createSheetHeaders();
                    lblSpreadsheet.Text = "The selected spreadsheet is:" + System.Environment.NewLine + OFD.FileName;
                }
                else
                {
                    createExcelWorkbook(OFD.FileName);
                    lblSpreadsheet.Text = "The selected spreadsheet is:" + System.Environment.NewLine + OFD.FileName;
                }


                //if (OFD.CheckFileExists == false )
                //{
                //    if (createExcelWorkbook(OFD.FileName ))
                //    { lblSpreadsheet.Text = "The selected spreadsheet is:" + System.Environment.NewLine + OFD.FileName; }
                //}
                //else if (OFD.CheckFileExists == True)
                //{
                //    _workbook = _app.Workbooks.Open(OFD.FileName,
                //        0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                //        true, false, 0, true, false, false);

                //    lblSpreadsheet.Text = "The selected spreadsheet is:" + System.Environment.NewLine + OFD.FileName;
                //}
            }
        }

        private void txtBoringName_TextChanged(object sender, EventArgs e)
        {
            if (!txtBoringName.Text.Equals("") && !txtBeginDepth.Text.Equals("") && !txtEndDepth.Text.Equals(""))
            {
                btnOpenExcel.Enabled = true;
            }
        }

        private void btnSetVoltage_Click(object sender, EventArgs e)
        {

            var tmpUP = "";
            var tmpDN = "";
            double tmpAbove1, tmpAbove2, tmpAbove3, tmpBelow1, tmpBelow2, tmpBelow3, tmp3;

            _DHPressureSwitchClient = new DdeClient("WinWedge", cbDHPressureSwitchCOMPort.Text.ToString());
            _DHPressureSwitchClient.Connect();
            // _DHPressureSwitchClient.Execute("[RESET]", 60000);

            if (!_DHPressureSwitchClient.IsConnected) { _DHPressureSwitchClient.Connect(); }

            _DHPressureSwitchClient.Execute("[SENDOUT('$1RB',13,10)]", _timeout);
            tmpUP = _DHPressureSwitchClient.Request("Field(1)", _timeout);
            tmpDN = _DHPressureSwitchClient.Request("Field(2)", _timeout);

            tmpUP = tmpUP.Replace("+", "");
            tmpUP = tmpUP.Replace("*", "");
            tmpUP = tmpUP.Replace("\0", "");
            tmpUP = tmpUP.TrimStart('0');

            tmpDN = tmpDN.Replace("+", "");
            tmpDN = tmpDN.Replace("*", "");
            tmpDN = tmpDN.Replace("\0", "");
            tmpDN = tmpDN.TrimStart('0');

            UpHoleVoltage.Text = Convert.ToString(tmpUP);
            DNHoleVoltage.Text = Convert.ToString(tmpDN);
            Thread.Sleep(1000);

        }


        private void comboBox1_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                comboBox1.Items.Add(sheet.Name);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            Reading r = new Reading();
            Random rnd = new Random();
            DataRow rowPrint;

            _dtPrintChart.Clear();
            rowPrint = _dtPrintChart.NewRow();

            var numRows = 0;
            string worksheetName = comboBox1.Text;
            Microsoft.Office.Interop.Excel.Worksheet sheet;
            //sheet = (Worksheet)_workbook.Sheets[worksheetName];
            sheet = (Microsoft.Office.Interop.Excel.Worksheet)_workbook.Sheets[worksheetName];

            numRows = sheet.Cells.Find("*", Type.Missing, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, false, Type.Missing).Row;

            for (int i = 2; i < numRows; i++)
            {
                rowPrint = _dtPrintChart.NewRow();

                _rangePrint = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i, 4];
                rowPrint[0] = _rangePrint.Value2;
                _rangePrint = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i, 12];
                rowPrint[1] = _rangePrint.Value2;
                _rangePrint = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i, 16];
                rowPrint[2] = _rangePrint.Value2;
                _rangePrint = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i, 20];
                rowPrint["DateTime"] = _rangePrint.Value2;

                _dtPrintChart.Rows.Add(rowPrint);

            }

            chartControl1.RefreshData();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            FlowReport report = new FlowReport();
            Stream chartLayout = new MemoryStream();
            chartControl1.SaveToStream(chartLayout);
            XRChart xrChart = new XRChart();

            chartLayout.Seek(0, System.IO.SeekOrigin.Begin);
            xrChart.LoadFromStream(chartLayout);
            xrChart.DataSource = chartControl1.DataSource;
            float height = report.PageHeight - report.Margins.Top - report.Margins.Bottom;
            //height = height - report.Bands.GetBandByType(typeof(PageHeaderBand)).Height;
            height = height - report.Bands.GetBandByType(typeof(PageFooterBand)).HeightF;
            height = height - 1;
            int width = report.PageWidth - report.Margins.Left - report.Margins.Right;

            xrChart.HeightF = height;
            xrChart.WidthF = width;

            xrChart.Series[0].LabelsVisibility = DevExpress.Utils.DefaultBoolean.False;
            xrChart.Series[1].LabelsVisibility = DevExpress.Utils.DefaultBoolean.False;
            xrChart.Series[2].LabelsVisibility = DevExpress.Utils.DefaultBoolean.False;

            report.Bands.GetBandByType(typeof(DetailBand)).Controls.Add(xrChart);

            report.xrlClientCompany.Text = txtCompany.Text;
            report.xlrVisionJobNo.Text = txtVisionJob.Text;
            report.xlrSiteAddress.Text = textBoxProjectAddress.Text;
            report.xlrBoringName.Text = txtBoringName.Text;
            report.xlrBeginDepth.Text = txtBeginDepth.Text;
            report.xlrEndDepth.Text = txtEndDepth.Text;

            this.Cursor = Cursors.Default;
            report.ShowPreviewDialog();
        }

        private void chartControl1_ObjectSelected(object sender, HotTrackEventArgs e)
        {
            objectSelected(e, ref chart1AnnotationNum, chartControl1, dataGridView1);
        }
        private void chartReadings_ObjectSelected(object sender, HotTrackEventArgs e)
        {
            objectSelected(e, ref chartReadingsAnnotationNum, chartReadings, dgvAnnotations );
        }
        
        private void objectSelected(HotTrackEventArgs e, ref int annotationCounter, ChartControl  chart, DataGridView dgv)
        {
            //existing annotation
            if (e.HitInfo.InAnnotation)
            {
                annatotationExists(e, dgv );
                //currentAnnotation = e.HitInfo.Annotation as TextAnnotation;
                //this.textBox1.Text = currentAnnotation.Text;
            }
            else
                currentAnnotation = null;

            if (e.AdditionalObject is SeriesPoint)
            {
                SeriesPoint sp = e.AdditionalObject as SeriesPoint;
                if (sp != null)
                {
                    if (!e.HitInfo.InAnnotation)
                    {
                        if (!annatotationExists(e, dgv))
                        {
                            TextAnnotation annotation = new TextAnnotation(annotationCounter.ToString() , "");
                            annotation.Name = annotationCounter.ToString();
                            dgv.Rows.Add(((DevExpress.XtraCharts.Series)(e.Object)).Name,
                                ((DevExpress.XtraCharts.SeriesPoint)(e.AdditionalObject)).Values[0],
                                "Enter Annotation...",
                                annotationCounter);
                            annotation.AnchorPoint = new SeriesPointAnchorPoint(sp);
                            annotation.ShapePosition = new RelativePosition();
                            RelativePosition position = annotation.ShapePosition as RelativePosition;
                            position.ConnectorLength = 50;
                            annotation.RuntimeMoving = true;
                            for (int i = 0; i < chart.AnnotationRepository.Count; i++)
                            {
                                Annotation ax = chart.AnnotationRepository[i];
                                SeriesPointAnchorPoint sx = ax.AnchorPoint as SeriesPointAnchorPoint;
                                if (sx.SeriesPoint == sp)
                                    return;
                            }
                            chart.AnnotationRepository.Add(annotation);
                            annotationCounter++;
                        }
                    }
                }
            }
        }

        //HotTrackEventArgs e, ref int annotationCounter, ChartControl  chart, DataGridView dgv
        private bool annatotationExists(DevExpress.XtraCharts.HotTrackEventArgs e, DataGridView dgv)
        {
            try
            {
                string name = "";
                double value = 0;
                int colCompare1 = 0;
                int colCompare2 = 0;

                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                if (e.Object is DevExpress.XtraCharts.Series)
                {
                    name = ((DevExpress.XtraCharts.Series)(e.Object)).Name;
                    value = ((DevExpress.XtraCharts.SeriesPoint)(e.AdditionalObject)).Values[0];
                    colCompare1 = 0;
                    colCompare2 = 1;
                }
                else if (e.Object is DevExpress.XtraCharts.TextAnnotation)
                {
                    name = ((DevExpress.XtraCharts.TextAnnotation)(e.Object)).Name;
                    value = ((DevExpress.XtraCharts.SeriesPoint)(e.AdditionalObject)).Values[0];
                    colCompare1 = 3;
                    colCompare2 = 0;
                }
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.Cells[colCompare1].Value.ToString().Equals(name))
                    {
                        if (colCompare2 == 0)
                        {
                            dgv.ClearSelection();
                            dgv.CurrentCell = row.Cells[2];
                            dgv.BeginEdit(true);
                            //row.Cells[2].Selected = true;
                            //row.Selected = true;
                            return true;
                        }
                        else
                        {
                            if (row.Cells[colCompare2].Value.Equals(value))
                            {
                                dgv.ClearSelection();
                                dgv.CurrentCell = row.Cells[2];
                                dgv.BeginEdit(true);
                                //row.Cells[2].Selected = true;
                                //row.Selected = true;
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return false;
        }

        private void dgvAnnotations_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                ((TextAnnotation)chartControl1.AnnotationRepository.GetElementByIndex((int)dgvAnnotations.Rows[e.RowIndex].Cells[3].Value)).Text
                    = dgvAnnotations.Rows[e.RowIndex].Cells[2].Value.ToString();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void dgvAnnotations_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            {
                DataGridViewRow dgvr = dgvAnnotations.CurrentRow;
                ChartElementNamed cen = chartReadings .AnnotationRepository.GetElementByName(dgvr.Cells[3].Value.ToString());
                chartReadings.AnnotationRepository.Remove(cen);
                dgvAnnotations.Rows.Remove(dgvAnnotations.CurrentRow);
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            {
                DataGridViewRow dgvr = dataGridView1.CurrentRow;
                ChartElementNamed cen = chartControl1.AnnotationRepository.GetElementByName(dgvr.Cells[3].Value.ToString());
                chartControl1.AnnotationRepository.Remove(cen);
                dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
            }
        }
    }
}