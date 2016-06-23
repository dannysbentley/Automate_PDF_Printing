using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SOM.RevitTools.Automate_PDF_Printing
{
    class DefaultValues
    {
        // Revit version
        public static string s_strRevitVersion = "2016";

        // Path to main folder
        public static string s_strMainPrintPath = @"C:\PROGRAMMING SOM\AutoPDFPrint" + s_strRevitVersion + @"\";

        // Path to temp folder
        public static string s_strTempPrintPath = s_strMainPrintPath + @"Temp\";

        // Path for log files
        public static string s_strLogFilePath = s_strMainPrintPath + @"Logs\" + Environment.GetEnvironmentVariable("USERNAME") + @"\";

        // Print driver name
        public static string s_strPrintDriverName = "Adobe PDF";

        // Revit Server abbrev
        public static string s_strRSAbbrevName = "RSN:";

        // Log file divide line
        public static string s_strLogFileDivide = "---------------------------------------------------------------------------------------------------------------";

        // Log file double divide line
        public static string s_strLogFileDblDivide = "===============================================================================================================";

        /*------------------------------------------------------------------------------------**/

        public string strRevitVersion
        {
            get
            {
                return s_strRevitVersion;
            }
        }

        public string strMainPrintPath
        {
            get
            {
                return s_strMainPrintPath;
            }
        }

        public string strTempPrintPath
        {
            get
            {
                return s_strTempPrintPath;
            }
        }

        public string strLogFilePath
        {
            get
            {
                return s_strLogFilePath;
            }
        }

        public string strPrintDriverName
        {
            get
            {
                return s_strPrintDriverName;
            }
        }

        public string strRSAbbrevName
        {
            get
            {
                return s_strRSAbbrevName;
            }
        }

        public string strLogFileDivide
        {
            get
            {
                return s_strLogFileDivide;
            }
        }

        public string strLogFileDblDivide
        {
            get
            {
                return s_strLogFileDblDivide;
            }
        }
    }
}
