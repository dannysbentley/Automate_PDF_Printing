#region Namespaces
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using WinForms = System.Windows.Forms;
using System.Management;
#endregion

namespace SOM.RevitTools.Automate_PDF_Printing
{
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class App : IExternalApplication
    {
        // Public variables

        // Only used for the saving of a RVT
        public static Autodesk.Revit.DB.Document m_RevitDocument;
        public static UIControlledApplication m_uiControlApp;
        public List<string> strListFromCFG = new List<string>();
        public static string strProcessDate = string.Empty;
        public static string strLogfileFullPath = string.Empty;
        public string strPDFPrintCFG = string.Empty;

        #region IExternalApplication Members

        /// <summary>
        /// Implement OnStartup method of IExternalApplication interface.
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            DefaultValues oDefaultValues = new DefaultValues();
            Methods oMethods = new Methods();

            // Launch Revit normally if temp folder is not found
            if (Directory.Exists(oDefaultValues.strTempPrintPath) == false)
                return Autodesk.Revit.UI.Result.Succeeded;

            // Get all CFG files from the temp folder
            List<string> strList = new List<string>();
            strList = oMethods.CFGFileGet(oDefaultValues.strTempPrintPath);
            
            // Launch Revit normally if one CFG file is not found
            if (strList.Count != 1)
                return Autodesk.Revit.UI.Result.Succeeded;
            else
                strPDFPrintCFG = strList[0].ToString();

            // Create the logs folder is needed
            if (Directory.Exists(oDefaultValues.strLogFilePath) == false)
                Directory.CreateDirectory(oDefaultValues.strLogFilePath);

            // Get the start date/time
            strProcessDate = oMethods.DateCurrentGet();

            // Log file full path
            strLogfileFullPath = oDefaultValues.strLogFilePath + strProcessDate + ".log";

            // Write to log file
            oMethods.TextFileWrite(strLogfileFullPath, oDefaultValues.strLogFileDblDivide, true);

            // Write to log file
            oMethods.TextFileWrite(strLogfileFullPath, "Process Start: " + oMethods.DateCurrentGet() +
                "-" + oMethods.TimeCurrentGet(":"), true);

            // Extract the lines of text from CFG file
            strListFromCFG = oMethods.CFGValuesGet(strPDFPrintCFG);

            // Delete CFG files in the temp folder
            oMethods.CFGFileDelete(oDefaultValues.strTempPrintPath);

            // Stop if no lines of text are found
            if (strListFromCFG.Count == 0)
            {
                // Write to log file
                oMethods.TextFileWrite(strLogfileFullPath, "No values found in: '" +
                    strPDFPrintCFG + "': " + oMethods.DateCurrentGet() +
                    "-" + oMethods.TimeCurrentGet(":") + ", Closing Revit", true);

                //Close Revit
                oMethods.RevitKill();
            }

            // Intiate the events
            m_uiControlApp = application;
            m_uiControlApp.Idling += new EventHandler<IdlingEventArgs>(idleUpdate);
            m_uiControlApp.DialogBoxShowing += new EventHandler<DialogBoxShowingEventArgs>(HandleDialogBoxShowing);

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        /// <summary>
        /// Implement OnShutdown method of IExternalApplication interface. 
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        #endregion

        /*------------------------------------------------------------------------------------**/
        /// <summary>
        /// Idle Event
        /// </summary>
        /// <returns> void </returns>
        /// <author>Dan.Tartaglia </author>                              <date>04/2013</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public void idleUpdate(object sender, IdlingEventArgs e)
        {
            int i = 0;
            string strProcess = string.Empty;
            string strSheetSet = string.Empty;
            string strPrintSet = string.Empty;
            string strRVTFullPath = string.Empty;
            string strPDFOutPath = string.Empty;
            string strWDrivePath = string.Empty;

            UIApplication uiApp = sender as UIApplication;
            DefaultValues oDefaultValues = new DefaultValues();
            Methods oMethods = new Methods();

            try
            {
                // Itereate through each line of text
                foreach (string strTextLine in strListFromCFG)
                {
                    i++;

                    // Write to log file
                    oMethods.TextFileWrite(strLogfileFullPath, oDefaultValues.strLogFileDivide, true);

                    // Write to log file
                    oMethods.TextFileWrite(strLogfileFullPath, "Processing line " + i.ToString() +
                        " in CFG file: " + oMethods.DateCurrentGet() + "-" +
                        oMethods.TimeCurrentGet(":"), true);

                    // Extract each value into an array
                    string[] strArgArray = strTextLine.Split(',');

                    // Process next line if 4 values are not found
                    if (strArgArray.Count() != 4)
                    {
                        // Write to log file
                        oMethods.TextFileWrite(strLogfileFullPath, "5 parameters not found in line: " +
                            i.ToString() + " in the CFG file processing : " + oMethods.DateCurrentGet() +
                            "-" + oMethods.TimeCurrentGet(":") + ", Processing next line", true);

                        // Process next line of text
                        continue;
                    }

                    // Get each value and remove beginning and trailing spaces
                    strSheetSet = strArgArray[0].ToString().Trim();
                    strPrintSet = strArgArray[1].ToString().Trim();
                    strRVTFullPath = strArgArray[2].ToString().Trim();
                    strPDFOutPath = strArgArray[3].ToString().Trim();

                    if (strRVTFullPath.Length > 3)
                    {
                        // Determine if the path is from Revit Server
                        if (strRVTFullPath.Substring(0, oDefaultValues.strRSAbbrevName.Length).ToUpper() ==
                            oDefaultValues.strRSAbbrevName.ToUpper())
                        {
                            // Verify the RVT file exists
                            if (ModelPathUtils.IsValidUserVisibleFullServerPath(strRVTFullPath) == false)
                            {
                                // Write to log file
                                oMethods.TextFileWrite(strLogfileFullPath, "RVT file not found: '" + strRVTFullPath + "' on line: " +
                                    i.ToString() + " : " + oMethods.DateCurrentGet() +
                                    "-" + oMethods.TimeCurrentGet(":") + ", Processing next line", true);

                                continue;
                            }
                        }
                        else
                        {
                            // Verify the RVT file exists
                            if (File.Exists(strRVTFullPath) == false)
                            {
                                // Write to log file
                                oMethods.TextFileWrite(strLogfileFullPath, "RVT file not found: '" + strRVTFullPath + "' on line: " +
                                    i.ToString() + " : " + oMethods.DateCurrentGet() +
                                    "-" + oMethods.TimeCurrentGet(":") + ", Processing next line", true);

                                continue;
                            }
                        }
                    }

                    // Write to log file
                    oMethods.TextFileWrite(strLogfileFullPath, "File open start: " + oMethods.TimeCurrentGet(":"), true);

                    // Open the RVT detached from Central
                    Document doc = oMethods.OpenRVTDetached(strRVTFullPath, uiApp);

                    // Process next line if Document is null
                    if (doc == null)
                    {
                        // Write to log file
                        oMethods.TextFileWrite(strLogfileFullPath, "Could not open file: '" + strRVTFullPath + "' on line: " +
                            i.ToString() + " : " + oMethods.DateCurrentGet() +
                            "-" + oMethods.TimeCurrentGet(":") + ", Processing next line", true);

                        // Process next line of text
                        continue;
                    }

                    // Print the PDFs
                    oMethods.PrintToPDF(strRVTFullPath, strSheetSet, strPrintSet,
                        strPDFOutPath, strProcessDate, strLogfileFullPath, i, doc);
                }

                // Write to log file
                oMethods.TextFileWrite(strLogfileFullPath, oDefaultValues.strLogFileDivide, true);

                // Write to log file
                oMethods.TextFileWrite(strLogfileFullPath, "Process End: " + oMethods.DateCurrentGet() +
                    "-" + oMethods.TimeCurrentGet(":"), true);

                // Write to log file
                oMethods.TextFileWrite(strLogfileFullPath, oDefaultValues.strLogFileDblDivide, true);

                //Close Revit
                oMethods.RevitKill();
            }
            catch (Exception ex)
            {
                // Write to log file
                oMethods.TextFileWrite(strLogfileFullPath, "Error: '" + ex.Message + " : " +
                    oMethods.DateCurrentGet() + "-" + oMethods.TimeCurrentGet(":") +
                    ", Processing stopped", true);
            }

            // Quit the events
            m_uiControlApp.DialogBoxShowing -= new EventHandler<DialogBoxShowingEventArgs>(HandleDialogBoxShowing);
            m_uiControlApp.Idling -= new EventHandler<IdlingEventArgs>(idleUpdate);
        }

        // Captures event when Revit message boxes display
        /*------------------------------------------------------------------------------------**/
        /// <author>Dan.Tartaglia </author>                              <date>01/2010</date>
        /*--------------+---------------+---------------+---------------+---------------+------*/
        public static void HandleDialogBoxShowing(object sender, Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs e)
        {
            TaskDialogShowingEventArgs e2 = e as TaskDialogShowingEventArgs;
            Document currDoc = null;

            try
            {
                UIApplication uiapp = sender as UIApplication;
                currDoc = uiapp.ActiveUIDocument.Document;
            }
            catch
            {
            }
            if (e2 == null)
            {
                //  Click OK
                e.OverrideResult((int)WinForms.DialogResult.OK);
            }
            if (e2 != null)
            {
                //  Click OK
                e2.OverrideResult((int)WinForms.DialogResult.OK);
            }
        }
    }
}
