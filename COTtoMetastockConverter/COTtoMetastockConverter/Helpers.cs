using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Windows;

namespace COTtoMetastockConverter
{
    public static class Helpers
    {
        public static bool isNumber(this string value)
        {
            double myNum = 0;
            if (Double.TryParse(value, out myNum))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //converts string dates, such as, MM/dd/yyyy to yyyyMMdd
        public static string formatStrDateToMetastockDate(this string strDate)
        {
            string metastockDate = "";
            if (strDate.Contains("/") || strDate.Contains("-"))
            {
                DateTime dateObj;
                var possibleFormats = new[] { "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/d/yy", "yyyy/MM/dd", "yyyy-MM-dd", "yyyy/M/d" };
                if (DateTime.TryParseExact(strDate, possibleFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateObj))
                {
                    metastockDate = dateObj.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                }
            }
            return metastockDate;
        }
        public static string getStrValue(object val)
        {
            if (val == null) return "";
            return val.ToString().Trim();
        }
        public static string getPath(string folder, string filename)
        {
            return Path.Combine(folder, filename);
        }
        public static double getDblValue(object val)
        {
            if (val == null) return 0;
            double myNum = 0;
            if (Double.TryParse(val.ToString(), out myNum))
            {
                return myNum;
            }
            else
            {
                return 0;
            }
        }
        public static void outputArrayListToFile(string path, ArrayList al)
        {
            //open streamwriter with overwrite option
            using (var writer = new StreamWriter(path, false))
            {
                //add Metastock headers from configuration file
                writer.WriteLine(ConfigHelpers.getConfigVal("metastockheaders").ToUpper());
                foreach (Object obj in al)
                {
                    if (obj != null)
                    {
                        string line = obj.ToString();
                        if (line != String.Empty) writer.WriteLine(line);
                    }
                }
            }
        } 
    }
}
