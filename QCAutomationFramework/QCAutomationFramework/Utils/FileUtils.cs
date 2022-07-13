using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace QCAutomationFramework.Utils
{
    public class FileUtils
    {
        public FileUtils()
        {
        }

        public string[] OpenFile()
        {
            OpenFileDialog fDialog = new OpenFileDialog();
            fDialog.Title = "open a presentation file";
            fDialog.Filter = "PPT Files(2003)|*.ppt|PPT Files(2007)|*.pptx|All Files|*.*";
            fDialog.Multiselect = true;
            //fDialog.InitialDirectory = @"D:\";
            string[] pptFileCollection = new string[] {};

            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                return fDialog.FileNames;
            }
            else
                return pptFileCollection;
        }

        //public string SaveFile()
        //{
        //    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
        //    saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
        //    saveFileDialog1.Title = "Save File As...";
        //    saveFileDialog1.InitialDirectory = @"C:\";

        //    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        return saveFileDialog1.FileName;
        //    }

        //    return string.Empty;
        //}

        public void WriteToLog(string message, string presFileName)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(Application.StartupPath + @"\Log_Files\LogTool_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + ".txt", true))
            {
                file.WriteLine("---------------------------------------------------------------------");
                file.WriteLine("Error At: " + DateTime.Now.ToString() + "; For file: " + presFileName);
                file.WriteLine("=====================================================================");
                file.WriteLine(message);
                file.WriteLine("---------------------------------END---------------------------------");
            }
        }

    }
}
