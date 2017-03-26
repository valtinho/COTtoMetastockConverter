using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Reflection;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.ComponentModel;

namespace COTtoMetastockConverter
{
    /// <summary>
    /// Converts the Commitment of Traders's Futures Only Report from the CFTC to Metastock price text files
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _sourceFilePath = "";
        private string _destFolder = "";

        public MainWindow()
        {
            InitializeComponent();
            faCog.Visibility = Visibility.Hidden;
        }

        private void lblCFTCLink_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //open the CFTCLink with the default browser
            Process.Start(((TextBlock)lblCFTCLink.Content).Text);
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var fileBrowser = new OpenFileDialog();
            //default file extension 
            fileBrowser.DefaultExt = ".xls";
            // file extension filter
            fileBrowser.Filter = "Excel Files (*.xls)|*.xls";
            Nullable<bool> selectedFile = fileBrowser.ShowDialog();
            if (selectedFile == true)
            {
                //Excel file path
                _sourceFilePath = fileBrowser.FileName;
                txtFilePath.Text = _sourceFilePath;
                //destination folder is the same folder as the source Excel file
                _destFolder = Path.GetDirectoryName(_sourceFilePath).ToString();
                //show the destination paths
                lblOutputPath.Content = String.Format("{0}\n{1}\n{2}\n{3}",
                     Helpers.getPath(_destFolder, String.Concat(ConfigHelpers.getConfigVal("symbolcotcommercials"), "_W.txt")),
                     Helpers.getPath(_destFolder, String.Concat(ConfigHelpers.getConfigVal("symbolcotlargetraders"), "_W.txt")),
                     Helpers.getPath(_destFolder, String.Concat(ConfigHelpers.getConfigVal("symbolcotsmallspeculators"), "_W.txt")),
                     Helpers.getPath(_destFolder, String.Concat(ConfigHelpers.getConfigVal("symbolcotspread"), "_W.txt")));

                //enable the conversion button once an Excel file is chosen

                btnConvertImg.Source = getImage("button.png");
                btnConvert.IsEnabled = true;
            }
        }

        private ImageSource getImage(string fileName)
        {
            BitmapImage logo = new BitmapImage();
            logo.BeginInit();
            logo.UriSource = new Uri(String.Concat("pack://application:,,,/Images/", fileName));
            logo.EndInit();
            return logo;
        }

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            faCog.Visibility = Visibility.Visible;

            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = false;
            worker.WorkerSupportsCancellation = false;
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(10000);
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        { 
            DataTable sheetData = ExcelHelpers.convertXLSToDataTable(_sourceFilePath);
            if (sheetData != null)
            {
                e.Result = processDataTable(sheetData);
            }
            else
            {
                throw new Exception("ERROR. File does not contain valid data.");
            } 
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Back on the UI thread....UI calls are cool again.
            string msg = "";
            bool hasErrors = false;
            if(e.Error != null)
            {
                hasErrors = true;
                msg = e.Error.Message;
            }
            else
            {
                if(e.Result != null)
                {
                    msg = e.Result.ToString();
                }
            } 
            //show the results form
            if (msg != String.Empty)
            {
                ResultsWindow results = new ResultsWindow(msg);
                results.Show();
            }
            //open the folder only if there were exceptions
            if (!hasErrors)
            {
                Process.Start(_destFolder);
            }
            //reset the controls
            btnConvertImg.Source = getImage("button_disabled.png");
            txtFilePath.Text = "";
            lblOutputPath.Content = "";
            btnConvert.IsEnabled = false;
            faCog.Visibility = Visibility.Hidden;
        }
        //returns the results that need to be displayed on the second form
        private string processDataTable(DataTable sheetData)
        {
            int index_reportDateMMDDYYYY, index_CFTCContractMarketCode, index_openInterestAll, index_nonCommPositionsLongAll, index_nonCommPositionsShortAll, index_commPositionsLongAll, index_commPositionsShortAll, index_nonReptPositionsLongAll, index_nonReptPositionsShortAll;
            index_reportDateMMDDYYYY = index_CFTCContractMarketCode = index_openInterestAll = index_nonCommPositionsLongAll = index_nonCommPositionsShortAll = index_commPositionsLongAll = index_commPositionsShortAll = index_nonReptPositionsLongAll = index_nonReptPositionsShortAll = -1;

            var metastockCommercialLines = new ArrayList();
            var metastockLargeLines = new ArrayList();
            var metastockSmallSpecLines = new ArrayList();
            var metastockSpread = new ArrayList();

            //Determine the index of the important columns
            //assume the first row [0] contains the column names
            for (int i = 0; i < sheetData.Columns.Count; i++)
            {
                if(sheetData.Rows[0][i] != null)
                {
                    string colName = sheetData.Rows[0][i].ToString().Trim().ToLower();
                    if (colName.Equals(ConfigHelpers.getConfigVal("reportdate"))) index_reportDateMMDDYYYY = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("markecode"))) index_CFTCContractMarketCode = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("openinterest"))) index_openInterestAll = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("largetraderlongpositions"))) index_nonCommPositionsLongAll = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("largetradershortpositions"))) index_nonCommPositionsShortAll = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("commerciallongpositions"))) index_commPositionsLongAll = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("commercialshortpositions"))) index_commPositionsShortAll = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("smallspeclongpositions"))) index_nonReptPositionsLongAll = i;
                    else if (colName.Equals(ConfigHelpers.getConfigVal("smallspecshortpositions"))) index_nonReptPositionsShortAll = i;
                }
            }
            string columnNotFoundError = "     - Could not find {0} data in the supplied spreadsheet.";

            if (index_reportDateMMDDYYYY == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Report Date"));
            if (index_CFTCContractMarketCode == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Market Code"));
            if (index_openInterestAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Open Interest"));
            if (index_nonCommPositionsLongAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Large Traders Long Positions"));
            if (index_nonCommPositionsShortAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Large Traders Short Positions"));
            if (index_commPositionsLongAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Commercials Long Positions"));
            if (index_commPositionsShortAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Commercials Short Positions"));
            if (index_nonReptPositionsLongAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Small Speculators Short Positions"));
            if (index_nonReptPositionsShortAll == -1) ErrorHelpers.deferredEx(String.Format(columnNotFoundError, "Small Speculators Short Positions"));

            List<string> errLst = ErrorHelpers.errorMessages();
            if(errLst.Count > 0)
            {
                string errorsStr = String.Join(Environment.NewLine, errLst.ToArray());
                errorsStr = String.Concat("ERROR:", Environment.NewLine, errorsStr, Environment.NewLine, "Please supply a valid Futures-Only COT Traders in Financial Futures file from the CFTC website.");
                //Cancel processing.
                throw new Exception(errorsStr);
            }

            var sbResults = new StringBuilder("", 10000);

            sbResults.AppendFormat("#\t\tDATE\t\t\tCOM\t\tLG\t\tSM\t\tSPR\n");
            int count = 0;
            int totalRows = sheetData.Rows.Count;
            //i=1 to skip the first line
            for (int i = 1; i < totalRows; i++)
            {
                //must have at least the following columns: 
                //date,open,high,low,close,volume 
                DataRow cells = sheetData.Rows[i];
                string strDate = Helpers.getStrValue(cells[index_reportDateMMDDYYYY]);
                string market_Code = Helpers.getStrValue(cells[index_CFTCContractMarketCode]);
                if (market_Code.Equals(ConfigHelpers.getConfigVal("spmarketcode")))
                {
                    string metastockDate = strDate.formatStrDateToMetastockDate(); 
                    if (metastockDate != "")
                    {
                        count++;
                        double openInterest = Helpers.getDblValue(cells[index_openInterestAll]);
                        double largeLong = Helpers.getDblValue(cells[index_nonCommPositionsLongAll]);
                        double largeShort = Helpers.getDblValue(cells[index_nonCommPositionsShortAll]);
                        double commercialsLong = Helpers.getDblValue(cells[index_commPositionsLongAll]);
                        double commercialsShort = Helpers.getDblValue(cells[index_commPositionsShortAll]);
                        double smallSpecsLong = Helpers.getDblValue(cells[index_nonReptPositionsLongAll]);
                        double smallSpecsShort = Helpers.getDblValue(cells[index_nonReptPositionsShortAll]);
                        double netCommercials = 0;
                        double netLargeTraders = 0;
                        double netSmallSpec = 0;
                        double netSpread = 0;
                        //ensure division by zero will never happen by accident
                        if (openInterest > 0)
                        {
                            netCommercials = Math.Round(((commercialsLong - commercialsShort) / openInterest) * 100, 2);
                            netLargeTraders = Math.Round(((largeLong - largeShort) / openInterest) * 100, 2);
                            netSmallSpec = Math.Round(((smallSpecsLong - smallSpecsShort) / openInterest) * 100, 2);
                            netSpread = Math.Round((netCommercials - netSmallSpec), 2);
                        }
                        else
                        {
                            //Cancel processing.
                            throw new Exception(String.Format("ERROR. Cannot convert line {0}. Open Interest is zero.", i.ToString()));
                        }
                        sbResults.AppendFormat("{0}\t\t{1}\t\t{2}\t\t{3}\t\t{4}\t\t{5}\n", count.ToString(), metastockDate, netCommercials, netLargeTraders, netSmallSpec, netSpread);

                        metastockCommercialLines.Add(String.Format("SP500 Fut COT COM,{0},D,{1},000000,{2},{2},{2},{2},0", ConfigHelpers.getConfigVal("symbolcotcommercials"), metastockDate, netCommercials));
                        metastockLargeLines.Add(String.Format("SP500 Fut COT LG,{0},D,{1},000000,{2},{2},{2},{2},0", ConfigHelpers.getConfigVal("symbolcotlargetraders"), metastockDate, netLargeTraders));
                        metastockSmallSpecLines.Add(String.Format("SP500 Fut COT SM,{0},D,{1},000000,{2},{2},{2},{2},0", ConfigHelpers.getConfigVal("symbolcotsmallspeculators"), metastockDate, netSmallSpec));
                        metastockSpread.Add(String.Format("SP500 Fut COT Spread,{0},D,{1},000000,{2},{2},{2},{2},0", ConfigHelpers.getConfigVal("symbolcotspread"), metastockDate, netSpread));
                    }
                    else
                    {
                        //Cancel processing.
                        throw new Exception(String.Format("ERROR. Cannot convert line {0}. Missing date.\n", i.ToString()));
                    } 
                } 
            }
            metastockCommercialLines.Sort();
            metastockLargeLines.Sort();
            metastockSmallSpecLines.Sort();
            metastockSpread.Sort();

            string comFilePath = Path.Combine(_destFolder, String.Format("{0}_W.txt",ConfigHelpers.getConfigVal("symbolcotcommercials")));
            string lgFilePath = Path.Combine(_destFolder, String.Format("{0}_W.txt", ConfigHelpers.getConfigVal("symbolcotlargetraders")));
            string smFilePath = Path.Combine(_destFolder, String.Format("{0}_W.txt", ConfigHelpers.getConfigVal("symbolcotsmallspeculators")));
            string spreadFilePath = Path.Combine(_destFolder, String.Format("{0}_W.txt", ConfigHelpers.getConfigVal("symbolcotspread")));
            
            //save ArrayLists to files
            Helpers.outputArrayListToFile(comFilePath, metastockCommercialLines);
            Helpers.outputArrayListToFile(lgFilePath, metastockLargeLines);
            Helpers.outputArrayListToFile(smFilePath, metastockSmallSpecLines);
            Helpers.outputArrayListToFile(spreadFilePath, metastockSpread);

            sbResults.AppendFormat("\nConversion was successful. {0} lines processed. {1} lines extracted.", totalRows, metastockCommercialLines.Count.ToString());
            
            //return the process' summary as a string  
            return sbResults.ToString();
        }        
    }
}
