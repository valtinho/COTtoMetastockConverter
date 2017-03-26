using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace COTtoMetastockConverter
{
    public static class ConfigHelpers
    {
        private static Hashtable _htConfig;

        public static string getConfigVal(string keyName)
        {
            if (_htConfig == null) loadConfigFile();
            if (_htConfig.ContainsKey(keyName))
            {
                //only stock symbol values are returned in uppercase
                if (keyName.StartsWith("symbol"))
                {
                    return _htConfig[keyName].ToString().ToUpper();
                }
                else
                { 
                    return _htConfig[keyName].ToString().ToLower();
                }
            }
            else
            {
                ErrorHelpers.immediateEx((String.Format("ERROR. Could not find the configuration value for {0}.", keyName)));
            }
            return "";
        }
        private static void loadConfigFile()
        {
            //use reflection to get the parent folder of the exe program
            string configPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "configuration.txt");
            initHashTable();
            if (!File.Exists(configPath))
            {
                //create default configuration.txt if it has been deleted                
                createDefaultConfigurationTxt(configPath);
            }
            else
            {
                loadConfigValues(configPath);
            }
        }
        private static void initHashTable()
        {
            //initializes the Hashtable with default values
            _htConfig = new Hashtable();
            _htConfig.Add("reportdate", "report_date_as_mm_dd_yyyy");
            _htConfig.Add("markecode", "cftc_contract_market_code");
            _htConfig.Add("openinterest", "open_interest_all");
            _htConfig.Add("largetraderlongpositions", "noncomm_positions_long_all");
            _htConfig.Add("largetradershortpositions", "noncomm_positions_short_all");
            _htConfig.Add("commerciallongpositions", "comm_positions_long_all");
            _htConfig.Add("commercialshortpositions", "comm_positions_short_all");
            _htConfig.Add("smallspeclongpositions", "nonrept_positions_long_all");
            _htConfig.Add("smallspecshortpositions", "nonrept_positions_short_all");
            _htConfig.Add("spmarketcode", "138741");
            _htConfig.Add("symbolcotcommercials", "^cotspxcom");
            _htConfig.Add("symbolcotlargetraders", "^cotspxlg");
            _htConfig.Add("symbolcotsmallspeculators", "^cotspxsm");
            _htConfig.Add("symbolcotspread", "^cotspxspread");
            _htConfig.Add("metastockheaders", "<NAME>,<TICKER>,<PER>,<DTYYYYMMDD>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOL>");
        }
        private static void createDefaultConfigurationTxt(string configPath)
        {
            //create the configuration.txt file with default settings
            var sb = new StringBuilder("", 1830);
            sb.Append("#***********************************************************************************\n");
            sb.Append("#***CONFIGURATION FILE**************************************************************\n");
            sb.Append("#***********************************************************************************\n");
            sb.Append("#Lines that start with a pound sign (#) are comments\n");
            sb.Append("#colons (:) separate variable names and values\n");
            sb.Append("#semicolon (;) indicates the end of a line\n");
            sb.Append("#***********************************************************************************\n");
            sb.Append("#Column Header Names are defined with [variable name]:[string name from spreadsheet]\n");
            sb.Append("#Neither the variable names, nor the values are case-sensitive\n");
            sb.Append("#White-spaces are ignored\n");
            sb.Append("#-----------------------------------------------------------------------------------\n");
            sb.Append("reportDate: report_date_as_mm_dd_yyyy;\n");
            sb.Append("markeCode: cftc_contract_market_code;\n");
            sb.Append("openInterest: open_interest_all;\n");
            sb.Append("largeTraderLongPositions: noncomm_positions_long_all;\n");
            sb.Append("largeTraderShortPositions: noncomm_positions_short_all;\n");
            sb.Append("commercialLongPositions: comm_positions_long_all;\n");
            sb.Append("commercialShortPositions: comm_positions_short_all;\n");
            sb.Append("smallSpecLongPositions: nonrept_positions_long_all;\n");
            sb.Append("smallSpecShortPositions: nonrept_positions_short_all;\n");
            sb.Append("#***********************************************************************************\n");
            sb.Append("#STOCK INDEX FUTURE, S&P 500 - INTERNATIONAL MONETARY MARKET\n");
            sb.Append("#or S&P 500 STOCK INDEX - INTERNATIONAL MONETARY MARKET\n");
            sb.Append("#CFTC_Contract_Market_Code\n");
            sb.Append("#-----------------------------------------------------------------------------------\n");
            sb.Append("spMarketCode: 138741;\n");
            sb.Append("#***********************************************************************************\n");
            sb.Append("# Output ticker symbols\n");
            sb.Append("#-----------------------------------------------------------------------------------\n");
            sb.Append("symbolCotCommercials:            ^ COTSPXCOM;\n");
            sb.Append("symbolcotLargeTraders:           ^ COTSPXLG;\n");
            sb.Append("symbolcotSmallSpeculators:       ^ COTSPXSM;\n");
            sb.Append("symbolcotSpread:                 ^ COTSPXSpread;\n");
            sb.Append("#***********************************************************************************\n");
            sb.Append("#Output files Metastock headers\n");
            sb.Append("#-----------------------------------------------------------------------------------\n");
            sb.Append("metastockHeaders:          <NAME>,<TICKER>,<PER>,<DTYYYYMMDD>,<TIME>,<OPEN>,<HIGH>,<LOW>,<CLOSE>,<VOL>\n");
            try
            {
                File.WriteAllText(configPath, sb.ToString());
            }
            catch (UnauthorizedAccessException)
            {
                ErrorHelpers.immediateEx("ERROR. Cannot create configuration file due to insufficient permissions.");
            }
            catch (IOException)
            {
                ErrorHelpers.immediateEx("ERROR. Cannot create configuration file due to an unexpected error.");
            }
            catch (Exception)
            {
                ErrorHelpers.immediateEx("ERROR. Cannot create configuration file due to an unknown error.");
            }

        }
        private static void loadConfigValues(string configPath)
        {
            //loads the values from the configuration.txt file into the existing Hashtable
            string[] allLines = File.ReadAllLines(configPath);
            //select only lines that don't start with #  AND contain a : and a ; (after the :)
            //remove all spaces from the strings before selecting
            IEnumerable<string> configLines =
                from aLine in allLines
                where (!aLine.StartsWith("#") && (aLine.IndexOf(":") > 3 && (aLine.IndexOf(";") > aLine.IndexOf(":"))))
                select Regex.Replace(aLine, @"\s+", "").ToLower();

            //process each valid line into keys and values and update the Hashtable
            foreach (string aLine in configLines)
            {
                int endOfLine = aLine.IndexOf(';');
                if (endOfLine == -1) endOfLine = aLine.Length;
                string[] keyVal = aLine.Substring(0, endOfLine).Split(':');
                if (keyVal.Length == 2)
                {
                    string key = keyVal[0].Trim();
                    string val = keyVal[1].Trim();
                    if (_htConfig.ContainsKey(key) && val != String.Empty) _htConfig[key] = val;
                }
            }
        }

    }
}
