using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SOM.RevitTools.Automate_PDF_Printing
{
    class Methods
    {
        /*------------------------------------------------------------------------------------**/
        /// Print to PDF process
        /// 
        /// </summary>
        /// <returns> true or false </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public bool PrintToPDF(string strRVTPath, string strSheetSet, string strPrintSet,
            string strPDFOutPath, string strProcessDate, string strLogfileFullPath, int intCntr,
            Document doc)
        {
            string strPDFExistOutFullPath = string.Empty;
            DefaultValues oDefaultValues = new DefaultValues();

            try
            {
                // Get the sheetset specified by user
                ViewSheetSet oVSS = ViewSheetSetGet(strSheetSet, doc);

                // Verify its found
                if (oVSS == null)
                {
                    // Write to log file
                    TextFileWrite(strLogfileFullPath, "Could not find the View/Sheet Set: '" +
                        strSheetSet + "' on line: " + intCntr.ToString() + " : " + DateCurrentGet() +
                        "-" + TimeCurrentGet(":") + ", Processing next line", true);

                    return false;
                }

                // Get print manager settings
                PrintManager oPM = doc.PrintManager;

                // Set the print driver to use
                oPM.SelectNewPrintDriver(oDefaultValues.strPrintDriverName);

                // Get the printset specified by user
                PrintSetting oPrintSetting = PrintSettingGet(strPrintSet, doc);

                // Verify its found
                if (oPrintSetting == null)
                {
                    // Write to log file
                    TextFileWrite(strLogfileFullPath, "Could not find the Print Set: '" +
                        strPrintSet + "' on line: " + intCntr.ToString() + " : " + DateCurrentGet() +
                        "-" + TimeCurrentGet(":") + ", Processing next line", true);

                    return false;
                }

                // Start the transaction
                Transaction transaction = new Transaction(doc);
                transaction.Start("Modify printer settings then print to PDF");

                // Use the printsettings for the printing
                oPM.PrintSetup.CurrentPrintSetting = oPrintSetting;

                // Set the sheetset for the printing
                oPM.PrintRange = Autodesk.Revit.DB.PrintRange.Select;
                ViewSheetSetting viewSheetSetting = oPM.ViewSheetSetting;
                viewSheetSetting.CurrentViewSheetSet = oVSS;

                // Set these flags
                oPM.PrintToFile = true;
                oPM.CombinedFile = false;

                // Set the PDF full path
                strPDFExistOutFullPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\Documents\AutoPDFPrint_Output.pdf";
                oPM.PrintToFileName = strPDFExistOutFullPath;

                // Delete the PDF if it already exists
                if (File.Exists(strPDFExistOutFullPath) == true)
                    File.Delete(strPDFExistOutFullPath);

                // Print to PDF
                oPM.SubmitPrint();

                // End the transaction
                transaction.Commit();

                // Move and rename the PDF
                PDFMoveRename(strProcessDate, Path.GetFileNameWithoutExtension(strRVTPath),
                    strPDFOutPath, strPDFExistOutFullPath, strLogfileFullPath, intCntr);
            }
            catch (Exception ex)
            {
                // Write to log file
                TextFileWrite(strLogfileFullPath, "Error: '" + ex.Message + " : " +
                    DateCurrentGet() + "-" + TimeCurrentGet(":") +
                    ", Processing stopped", true);

                return false;
            }
            return true;
        }

        /*------------------------------------------------------------------------------------**/
        /// Move PDFs to wanted output folder
        /// 
        /// </summary>
        /// <returns> void </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public void PDFMoveRename(string strProcessDate, string strRVTName, string strPDFOutPath,
            string strPDFExistOutFullPath, string strLogfileFullPath, int intCntr)
        {
            // Make sure a backslash is found
            if (!strPDFOutPath.EndsWith("\\"))
                strPDFOutPath += "\\";

            // Create the output full path            
            string strPDFNewOutFullPathNew = strPDFOutPath + strProcessDate + "\\" + strRVTName + ".pdf";

            // Sleep for 3 minutes (time needed will vary depending on the number of sheets created. 
            // This is needed for Revit to finish creating the PDF)
            Thread.Sleep(180000);

            try
            {
                // Verify the PDF exists
                if (File.Exists(strPDFExistOutFullPath) == false)
                {
                    // Write to log file
                    TextFileWrite(strLogfileFullPath, "Could not find the PDF: '" +
                        strPDFExistOutFullPath + "' on line: " + intCntr.ToString() + " : " + DateCurrentGet() +
                        "-" + TimeCurrentGet(":") + ", Processing next line", true);

                    return;
                }

                // Verify the output folder exists, if not create it
                if (Directory.Exists(Path.GetDirectoryName(strPDFNewOutFullPathNew)) == false)
                    Directory.CreateDirectory(Path.GetDirectoryName(strPDFNewOutFullPathNew));

                // Delete the PDF if it already exists
                if (File.Exists(strPDFNewOutFullPathNew) == true)
                    File.Delete(strPDFNewOutFullPathNew);

                // Copy the PDF to the output folder
                File.Copy(strPDFExistOutFullPath, strPDFNewOutFullPathNew, true);

                // Delete the PDF from the original location
                if (File.Exists(strPDFExistOutFullPath) == true)
                    File.Delete(strPDFExistOutFullPath);

                // Write memory of process to log file
                ProcessMemoryGet(strLogfileFullPath, intCntr);

                // Write to log file
                TextFileWrite(strLogfileFullPath, "'" + strPDFNewOutFullPathNew + "' created successfully" + ": " +
                    DateCurrentGet() + "-" + TimeCurrentGet(":"), true);
            }
            catch (Exception ex)
            {
                // Write to log file
                TextFileWrite(strLogfileFullPath, "Error: '" + ex.Message + " : " +
                    DateCurrentGet() + "-" + TimeCurrentGet(":") +
                    ", Processing stopped", true);
            }
        }

        /*------------------------------------------------------------------------------------**/
        /// Open the selected RVT file detached, open all user created worksets
        /// 
        /// </summary>
        /// <returns> Document </returns>
        /// <author>Dan.Tartaglia </author>                              <date>05/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public Document OpenRVTDetached(string strRVTFullPath, UIApplication uiApp)
        {
            DefaultValues oDefaultValues = new DefaultValues();

            try
            {
                Document doc = null;
                WorksetConfiguration oWSConfig = new WorksetConfiguration();
                IList<WorksetPreview> oWSPreviews = new List<WorksetPreview>();
                IList<WorksetId> oWSIDs = new List<WorksetId>();

                // Determine if the path is from Revit Server
                if (strRVTFullPath.Substring(0, oDefaultValues.strRSAbbrevName.Length).ToUpper() !=
                    oDefaultValues.strRSAbbrevName.ToUpper())
                {

                    BasicFileInfo oBasicFileInfo = BasicFileInfo.Extract(strRVTFullPath);

                    // Determine if the RVT is Workshared
                    if (oBasicFileInfo.IsWorkshared == false)
                    {
                        // Open the RVT
                        doc = uiApp.Application.OpenDocumentFile(strRVTFullPath);
                        oBasicFileInfo.Dispose();
                        return doc;
                    }

                    oBasicFileInfo.Dispose();
                }

                ModelPath oModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(strRVTFullPath);

                // Set the OpenOptions to open the RVT Detached from Central
                OpenOptions oOpenOpts = new OpenOptions();
                oOpenOpts.Audit = false;
                //oOpenOpts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;


                // Get the workset data                
                oWSPreviews = WorksharingUtils.GetUserWorksetInfo(oModelPath);

                // Iterate through each workset
                foreach (WorksetPreview oWSPreview in oWSPreviews)
                {
                    oWSIDs.Add(oWSPreview.Id);
                }

                // Open all user created worksets
                oWSConfig.Open(oWSIDs);

                // Set the OpenOptions object
                oOpenOpts.SetOpenWorksetsConfiguration(oWSConfig);

                // Open the RVT
                doc = uiApp.Application.OpenDocumentFile(oModelPath, oOpenOpts);

                //try to sync with central. 
                //TransactWithCentralOptions transact = new TransactWithCentralOptions();
                //SynchronizeWithCentralOptions syncWithCentral = new SynchronizeWithCentralOptions();
                //syncWithCentral.Comment = "Autosaved by the API at " + DateTime.Now;
                //RelinquishOptions relinqushOption = new RelinquishOptions(true);
                //relinqushOption.CheckedOutElements = true;
                //syncWithCentral.SetRelinquishOptions(relinqushOption);
                //uiApp.Application.WriteJournalComment("AutoSave to Central", true);
                //doc.SynchronizeWithCentral(transact, syncWithCentral);

                return doc;
            }
            catch
            {
                return null;
            }
        }

        /*------------------------------------------------------------------------------------**/
        /// End the Revit process
        /// 
        /// </summary>
        /// <returns> void </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public void RevitKill()
        {
            IList<int> intRevitProcIDs = new List<int>();

            // Get all revit.exe process ids
            intRevitProcIDs = ProcessGet("REVIT");

            // Make sure at least one valid id is found
            if (intRevitProcIDs.Count == 0)
                return;

            // Attempt to end all revit.exe precesses
            foreach (int intRevitProcID in intRevitProcIDs)
            {
                // End the process
                Process prs = System.Diagnostics.Process.GetProcessById(intRevitProcID);
                prs.Kill();
            }
        }

        /*------------------------------------------------------------------------------------**/
        /// Get all revit.exe ids
        /// 
        /// </summary>
        /// <returns> IList<int> </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public IList<int> ProcessGet(string strProcess)
        {
            IList<int> intRevitProcIDs = new List<int>();

            // Get the processes
            Process[] processlist = Process.GetProcesses();

            // Iterate through each process
            foreach (Process theprocess in processlist)
            {
                // Look for the wanted process name
                if (theprocess.ProcessName.ToUpper() == strProcess.ToUpper())
                {
                    // Get only processes by the logged in user
                    if (ProcessOwnerGet(theprocess.Id).ToUpper() == Environment.GetEnvironmentVariable("USERNAME").ToUpper())
                    {
                        intRevitProcIDs.Add(theprocess.Id);
                    }
                }
            }
            return intRevitProcIDs;
        }

        /*------------------------------------------------------------------------------------**/
        /// <summary>
        /// Get the user name for a process
        /// </summary>
        /// <returns> string </returns>
        /// <author>Dan.Tartaglia </author>                              <date>06/2010</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public static string ProcessOwnerGet(int processId)
        {
            // Get the processes
            string query = "Select * From Win32_Process Where ProcessID = " + processId;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection processList = searcher.Get();

            // Iterate through each process
            foreach (ManagementObject obj in processList)
            {
                // Get the user nam or owner
                string[] argList = new string[] { string.Empty };
                int returnVal = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList));
                if (returnVal == 0)
                    return argList[0];
            }
            return "NO OWNER";
        }

        /*------------------------------------------------------------------------------------**/
        /// Extract all lines of text from the CFG
        /// 
        /// </summary>
        /// <returns> IList<string> </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public List<string> CFGValuesGet(string strFileFullPath)
        {
            string line;
            List<string> strList = new List<string>();
            DefaultValues oDefaultValues = new DefaultValues();

            // Read the file and display it line by line.
            System.IO.StreamReader file = new System.IO.StreamReader(strFileFullPath);

            // Iterate the text file
            while ((line = file.ReadLine()) != null)
            {
                // Verify the string has enough chars to process
                if (line.Length > 1)
                {
                    // Skip comments
                    if (line.Substring(0, 1) != "#")
                        strList.Add(line);
                }
            }
            file.Close();
            return strList;
        }

        /*------------------------------------------------------------------------------------**/
        /// Get the Revit.exe memory and add to the log file
        /// 
        /// </summary>
        /// <returns> void </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public void ProcessMemoryGet(string strLogfileFullPath, int intCntr)
        {
            // Get the processes
            Process[] processlist = Process.GetProcesses();

            try
            {
                // Iterate through each process
                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName.ToUpper() == "REVIT")
                    {
                        int i = 0;
                        int intMemsize = 0; // intMemsize in Megabyte
                        string strMemsize = "";
                        string strUserName = "";

                        // Get the process memory size
                        PerformanceCounter PC = new PerformanceCounter();
                        PC.CategoryName = "Process";
                        PC.CounterName = "Working Set - Private";
                        PC.InstanceName = theprocess.ProcessName;

                        // Determine if a valid value is found
                        if (ValidValueFound(PC.NextValue()) == false)
                            continue;

                        intMemsize = Convert.ToInt32(PC.NextValue()) / (int)(1024);
                        PC.Close();
                        PC.Dispose();

                        // Iterate through memory value
                        foreach (char c in intMemsize.ToString().Reverse())
                        {
                            i++;

                            // Add char to the string
                            strMemsize = (c.ToString() + strMemsize);

                            // Add a comma after each 3 chars
                            if (i == 3 || i == 6 || i == 9 || i == 12 || i == 15)
                                strMemsize = ("," + strMemsize);
                        }

                        if (strMemsize.Length > 2)
                        {
                            // Remove the first char if it is a comma
                            if (strMemsize.Substring(0, 1) == ",")
                                strMemsize = strMemsize.Substring(1, strMemsize.Length - 1);

                            // Get the username from the process
                            strUserName = ProcessOwnerGet(theprocess.Id);

                            if (strUserName != "" && strUserName != null)

                                // Write to log file
                                TextFileWrite(strLogfileFullPath, "Current Revit.exe Memory Usage: '" +
                                    strMemsize + "K', Username: '" + strUserName + "'", true);
                            else
                                // Write to log file
                                TextFileWrite(strLogfileFullPath, "Current Revit.exe Memory Usage: '" +
                                    strMemsize + "K'", true);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        /*------------------------------------------------------------------------------------**/
        /// Determine if a valid value is found
        /// 
        /// </summary>
        /// <returns> true or false </returns>
        /// <author>Dan.Tartaglia </author>                              <date>05/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public bool ValidValueFound(float fltValue)
        {
            try
            {
                int test = Convert.ToInt32(fltValue);
            }
            catch
            {
                return false;
            }
            return true;
        }

        /*------------------------------------------------------------------------------------**/
        /// Get the CFG file full path
        /// 
        /// </summary>
        /// <returns> ArrayList </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public List<string> CFGFileGet(string strTempPrintPath)
        {
            string[] strFiles;
            List<string> strList = new List<string>();

            strFiles = Directory.GetFiles(strTempPrintPath);

            foreach (string strFile in strFiles)
            {
                if (Path.GetExtension(strFile).ToUpper() == ".CFG")
                    strList.Add(strFile);
            }
            return strList;
        }

        /*------------------------------------------------------------------------------------**/
        /// Delete CFG files in the temp folder
        /// 
        /// </summary>
        /// <returns> void </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public void CFGFileDelete(string strNBBJTempPrintPath)
        {
            DefaultValues oDefaultValues = new DefaultValues();

            try
            {
                string[] strFiles;
                strFiles = Directory.GetFiles(strNBBJTempPrintPath);

                foreach (string strFile in strFiles)
                {
                    if (Path.GetExtension(strFile).ToUpper() == ".CFG")
                        File.Delete(strFile);
                }
            }
            catch
            {
            }
        }

        /*------------------------------------------------------------------------------------**/
        /// Get the specified ViewSheetSet object
        /// 
        /// </summary>
        /// <returns> ViewSheetSet </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public ViewSheetSet ViewSheetSetGet(string strSheetSetName, Document doc)
        {
            FilteredElementCollector oViews = new FilteredElementCollector(doc);
            IList<Element> oViewSheets = oViews.OfClass(typeof(ViewSheetSet)).ToElements();

            foreach (Element elem in oViewSheets)
            {
                ViewSheetSet oViewSheetSet = elem as ViewSheetSet;

                if (strSheetSetName.ToUpper() == oViewSheetSet.Name.ToUpper())
                    return oViewSheetSet;
            }

            return null;
        }

        /*------------------------------------------------------------------------------------**/
        /// Get the specified PrintSetting object
        /// 
        /// </summary>
        /// <returns> PrintSetting </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public PrintSetting PrintSettingGet(string strPrintSettingName, Document doc)
        {
            ICollection<ElementId> oColl = doc.GetPrintSettingIds();

            // Iterate the PrintSettings
            foreach (ElementId elemID in oColl)
            {
                PrintSetting oPS = doc.GetElement(elemID) as PrintSetting;

                if (strPrintSettingName.ToUpper() == oPS.Name.ToUpper())
                    return oPS;
            }

            return null;
        }

        /*------------------------------------------------------------------------------------**/
        /// <summary>
        /// Get the current time HH:mm:ss
        /// </summary>
        /// <returns> string </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2010</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public string TimeCurrentGet(string strDivider)
        {
            return System.DateTime.Now.ToString("HH" + strDivider + "mm" + strDivider + "ss");
        }

        /*------------------------------------------------------------------------------------**/
        /// <summary>
        /// Get the current date yymmdd
        /// </summary>
        /// <returns> string </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2010</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public string DateCurrentGet()
        {
            string strCountryCode = "";
            try
            {
                // Get the name of the current country
                string strCountry = System.Globalization.RegionInfo.CurrentRegion.EnglishName;

                // Set the country code as needed
                switch (strCountry)
                {
                    case "United States":
                        strCountryCode = "en-US";
                        break;

                    case "United Kingdom":
                        strCountryCode = "en-GB";
                        break;

                    case "People's Republic of China":
                        strCountryCode = "zh-CN";
                        break;

                    default:
                        strCountryCode = "en-US";
                        break;
                }

                // Get the current date
                string strDate = (DateTime.Now.ToString("yyyyMMdd",
                    System.Globalization.CultureInfo.GetCultureInfo(strCountryCode)));

                // Verify the string has enough chars to process
                if (strDate.Length == 8)
                    return strDate.Substring(2, 6);
                else
                    return "";
            }
            catch
            {
                return "";
            }
        }

        /*------------------------------------------------------------------------------------**/
        /// <summary>
        /// Write to the output text file
        /// </summary>
        /// <returns> void </returns>
        /// <author>Dan.Tartaglia </author>                              <date>01/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public void TextFileWrite(string strTextFileName, string strValue, bool blnAppend)
        {
            try
            {
                // create a writer and open the file
                TextWriter tw = new StreamWriter(strTextFileName, blnAppend);

                // Add a line of text to the output file
                tw.WriteLine(strValue);

                // close the stream
                tw.Close();
            }
            catch
            {
            }
        }
    }
}
