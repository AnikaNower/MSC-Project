using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Collections;

namespace QCAutomationFramework.Utils
{
    public class ExcelUtils
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet;
        private object misValue;
        private Hashtable myHashtable;

        public ExcelUtils()
        {
            xlApp = new Excel.Application();
            misValue = System.Reflection.Missing.Value;
        }

        public Excel.Workbook GetWorkBook(string xlFileName)
        {
            xlWorkBook = xlApp.Workbooks.Open(xlFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            return xlWorkBook;
        }

        public Excel.Worksheet GetWorksheet(Excel.Workbook xlWorkbookParam, string xlSheetName)
        {
            xlWorkSheet = (Excel.Worksheet)xlWorkbookParam.Worksheets.get_Item(xlSheetName);

            return xlWorkSheet;
        }

        public IList<string> GetSheetNames(Excel.Workbook xlWorkbookParam)
        {
            IList<string> sheetNames = new List<string>();

            for (int sheetCount = 0; sheetCount < xlWorkbookParam.Sheets.Count; sheetCount++)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetCount + 1);
                if (xlWorkSheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
                {
                    sheetNames.Add(xlWorkSheet.Name);
                }
            }

            return sheetNames;
        }

        public void SaveXL(Excel.Workbook xlWorkbookParam, string XLFileName, string configType, string fileSeparator, bool checkStatus)
        {
            string configFolderPath = string.Empty;
            string configFolderName = string.Empty;
            string configReportName = string.Empty;

            if (configType == "NEW") configFolderName = "New_Config";
            else if (configType == "OLD") configFolderName = "Old_Config";

            configFolderPath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, configFolderName);

            if (fileSeparator == string.Empty) fileSeparator = ".";

            configReportName = GetReportName(XLFileName, fileSeparator, checkStatus) + "_" + configType + ".xls";

            xlWorkbookParam.SaveAs(configFolderPath + @"\" + configReportName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        }

        public string GetReportName(string fileName, string separator, bool checkStatus)
        {
            int index = 0;
            int startIndex;
            int length = fileName.Length;

            if (checkStatus)
            {
                startIndex = fileName.IndexOf(separator, 1) + 1;
                return fileName.Substring(0, startIndex - 1);
            }
            else
            {
                for (startIndex = 1; ; )
                {
                    index = fileName.IndexOf(separator, startIndex);
                    if (index > length || index == -1) break;
                    else startIndex = index + 1;
                }

                return fileName.Substring(0, startIndex - 1);
            }
        }

        public bool RunMacro(Excel.Workbook WB, string macroName, string arg1, string arg2, string arg3, string arg4)
        {
            WB.Application.Run(macroName, arg1, arg2, arg3, arg4);
            //WB.Application.Run(macroName, arg1, arg2, misValue,
            //                 misValue, misValue, misValue, misValue,
            //                 misValue, misValue, misValue, misValue,
            //                 misValue, misValue, misValue, misValue,
            //                 misValue, misValue, misValue, misValue,
            //                 misValue, misValue, misValue, misValue,
            //                 misValue, misValue, misValue, misValue,
            //                 misValue, misValue, misValue);

            return true;
        }

        public void CheckExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        public void KillExcel()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");

            // check to kill the right process
            foreach (Process ExcelProcess in AllProcesses)
            {

                if (myHashtable.ContainsValue(ExcelProcess.Id) == false || myHashtable.ContainsKey(ExcelProcess.Id) == false)
                    ExcelProcess.Kill();
            }

            AllProcesses = null;
        }

        public void CloseXLWorkbook(Excel.Workbook xlWorkbookParam)
        {
            xlWorkbookParam.Close(true, misValue, misValue);
        }
        
        public void CloseXLApplication()
        {
            xlApp.Quit();
        }
    }
}
