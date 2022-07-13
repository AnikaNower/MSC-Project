using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using QCAutomationFramework.Utils;
using System.IO;

namespace QCAutomationFramework
{
    public partial class PresConfigGenerator : Form
    {
        private int detailsCounter = 0;
        private int detailsCounterTable = 0;
        string errorMessage = string.Empty;
        bool error = false;

        private ExcelUtils xlUtils;
        private FileUtils fileUtils;

        public PresConfigGenerator()
        {
            InitializeComponent();

            xlUtils = new ExcelUtils();
            fileUtils = new FileUtils();
            radioButtonNew.Checked = true;

            CreateFolder("New_Config");
            CreateFolder("Old_Config");
            CreateFolder("QC_Result");
            CreateFolder("Log_Files");           

            xlUtils.CheckExcellProcesses();

            buttonAddNew.Click += new EventHandler(buttonAddNew_Click);   //file load kore         
            buttonRemoveNew.Click += new EventHandler(buttonRemoveNew_Click);//file remove kore

            buttonGenerate.Click += new EventHandler(buttonGenerate_Click);
            buttonReport.Click += new EventHandler(buttonReport_Click);

            buttonOpenReportFolder.Click += new EventHandler(buttonOpenReportFolder_Click);
            this.FormClosing += new FormClosingEventHandler(PresConfigGenerator_FormClosing);
            Application.ApplicationExit += new EventHandler(Application_ApplicationExit);
        }

        void buttonOpenReportFolder_Click(object sender, EventArgs e)
        {
            string myPath = Application.StartupPath;
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = myPath + @"\PPT Checker Help.ppt";
            prc.Start();
            //System.Diagnostics.Process.Start("POWERPNT.exe", Application.StartupPath + @"\IMS Data_KPI.ppt");
        }

        void Application_ApplicationExit(object sender, EventArgs e)
        {
            xlUtils.KillExcel();
        }

        void buttonRemoveNew_Click(object sender, EventArgs e)
        {
            if (listBoxFilesNew.SelectedItems.Count > 0)
                listBoxFilesNew.Items.RemoveAt(listBoxFilesNew.SelectedIndex);
        }

        void PresConfigGenerator_FormClosing(object sender, FormClosingEventArgs e)
        {
            xlUtils.KillExcel();
        }

        void buttonGenerate_Click(object sender, EventArgs e)
        {
            string configType = string.Empty;
            if (radioButtonNew.Checked) configType = "NEW";
            else if (radioButtonOld.Checked) configType = "OLD";

            int total = listBoxFilesNew.SelectedItems.Count;
            toolStripProgressBar1.Visible = true;
            this.Refresh();

            string templateFile = Application.StartupPath + "\\QC_Template.xls";

            if (listBoxFilesNew.SelectedItems.Count > 0)
            {
                toolStripProgressBar1.Value = 0;
                foreach (Object listItem in listBoxFilesNew.SelectedItems)
                {
                    string strItem = listItem as string;
                    PresentationUtils generated = new PresentationUtils(textBoxDirectoryNew.Text + @"\" + strItem);// presentationUtils e jay sekhan theke kaj hoy

                    CreatePresConfigXL(generated, templateFile, textBoxDirectoryNew.Text, strItem, configType);

                    generated.CloseApplication();
                    toolStripProgressBar1.Value += 100 / total;
                }
            }
            else
                MessageBox.Show("Please Select file(s) from the Listbox to generate the configuration file(s)", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            toolStripProgressBar1.Visible = false;
        }

        void buttonAddNew_Click(object sender, EventArgs e)
        {
            LoadFiles(listBoxFilesNew, textBoxDirectoryNew);            
        }

        void buttonReport_Click(object sender, EventArgs e)
        {
            string macroFile = string.Empty;
            macroFile = Application.StartupPath + @"\QC_Pres_Config.xls";

            string oldConfigDirectory = System.IO.Path.Combine(Application.StartupPath, "Old_Config");
            string newConfigDirectory = System.IO.Path.Combine(Application.StartupPath, "New_Config");

            if ((new DirectoryInfo(newConfigDirectory)).GetFiles("*.xls").Length != 0 && (new DirectoryInfo(oldConfigDirectory)).GetFiles("*.xls").Length != 0)
            {
                Excel.Workbook WBMacro = xlUtils.GetWorkBook(macroFile);
                fileUtils.WriteToLog(newConfigDirectory, oldConfigDirectory);
                if (xlUtils.RunMacro(WBMacro, "main", newConfigDirectory + @"\", oldConfigDirectory + @"\", Application.StartupPath + @"\QC_Result\", textBoxSpecialWords.Text)) MessageBox.Show("SUCCESS");
                else MessageBox.Show("SOMETHING is GOING WRONG");
            }
            else
            {
                MessageBox.Show("CONFIGURATION FILE is MISSING IN CONFIG FOLDER(s)", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //WBMacro.Save();
        }

        #region Helper Functions

        private void CreateFolder(string configFolderName)
        {
            string configFolderPath = string.Empty;
            configFolderPath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, configFolderName);

            if (!System.IO.Directory.Exists(configFolderPath))
                System.IO.Directory.CreateDirectory(configFolderPath);
        }

        public void LoadFiles(ListBox listBoxObject, TextBox textBoxObject)
        {
            //listBoxObject.Items.Clear();
            string directoryName = string.Empty;

            string[] pptFileCollection = fileUtils.OpenFile();
            
            foreach (string file in pptFileCollection)
            {
                System.IO.FileInfo fi = null;
                try
                {
                    fi = new System.IO.FileInfo(file);
                    listBoxObject.Items.Add(fi.Name);
                }
                catch (System.IO.FileNotFoundException ex)
                {
                    // To inform the user and continue is
                    // sufficient for this demonstration.
                    // Your application may require different behavior.
                    MessageBox.Show(ex.Message);
                    continue;
                }
                directoryName = fi.DirectoryName;
            }

            textBoxObject.Text = directoryName;
        }       

        private void CreatePresConfigXL(PresentationUtils presFile, string templateFilePath, string directoryName, string reportName, string configType)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            int slideCount;
            int slideCounter = 1; //counter for slides
            int lastRowText = 1;
            int lastRowImage = 1;
            int lastRowChart = 1;
            int lastRowChartDetails = 1;
            int lastRowTable = 1;
            int lastTableDetails = 1;

            slideCount = presFile.oPres.Slides.Count;

            //get a template file, in where the qc report will be saved
            //==========================================================
            xlWorkBook = xlUtils.GetWorkBook(templateFilePath);

            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Basic");
            
            BasicConfig(presFile.oPres, xlWorkSheet);

            foreach (PowerPoint.Slide oPPTSl in presFile.oPres.Slides)
            {
                foreach (PowerPoint.Shape oPPTShape in oPPTSl.Shapes)
                {
                    try
                    {
                        
                        if (checkBoxText.Checked && (oPPTShape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox || oPPTShape.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape || oPPTShape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder))
                        {
                            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Text");

                            TextConfig(xlWorkSheet, oPPTShape, lastRowText, slideCounter);

                            lastRowText++;
                        }
                        else if (checkBoxImage.Checked && oPPTShape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                        {
                            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Image");

                            ImageConfig(xlWorkSheet, oPPTShape, lastRowImage, slideCounter);

                            lastRowImage++;
                        }
                        else if (checkBoxChart.Checked && oPPTShape.Type == Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject && oPPTShape.OLEFormat.ProgID.StartsWith("Excel.Chart"))
                        {
                            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Chart");

                            ChartConfig(xlWorkSheet, oPPTShape, lastRowChart, slideCounter);

                            lastRowChart++;

                            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Chart Details");

                            ChartDetailsConfig(xlWorkSheet, oPPTShape, lastRowChartDetails, slideCounter);

                            lastRowChartDetails = lastRowChartDetails + detailsCounter;
                        }
                        else if (checkBoxTable.Checked && oPPTShape.Type == Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject && oPPTShape.OLEFormat.ProgID.StartsWith("Excel.Sheet"))
                        {
                            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Table");

                            TableConfig(xlWorkSheet, oPPTShape, lastRowTable, slideCounter);

                            lastRowTable++;

                            xlWorkSheet = xlUtils.GetWorksheet(xlWorkBook, "Table Details");

                            TableDetailsConfig(xlWorkSheet, oPPTShape, lastTableDetails, slideCounter);

                            lastTableDetails = lastTableDetails + detailsCounterTable;
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessage += "\nMessage From Master -->" + ex.Message + " \n" + "In Slide: " + slideCounter + ", Shape ID: " + oPPTShape.Name;
                        fileUtils.WriteToLog(errorMessage, reportName);
                        error = true;
                    }

                }
                slideCounter++;
            }

            xlUtils.SaveXL(xlWorkBook, reportName, configType, textBoxSeparator.Text, checkBoxOccr.Checked);

            xlUtils.CloseXLWorkbook(xlWorkBook);

            MessageBox.Show("TASK COMPLETE: Configuration Generation");
        }

        private void BasicConfig(PowerPoint.Presentation oPres, Excel.Worksheet xlBasicWorksheet)
        {
            int slideCount;
            int count = 1;

            slideCount = oPres.Slides.Count;

            foreach (PowerPoint.Slide sl in oPres.Slides)
            {
                xlBasicWorksheet.Cells[count + 1, 1] = count;
                xlBasicWorksheet.Cells[count + 1, 2] = sl.Shapes.Count;
                count++;
            }

            xlBasicWorksheet.Cells[count + 1, 1] = slideCount;
            (xlBasicWorksheet.Cells[count + 1, 2] as Excel.Range).Formula = "=SUM(B2:B" + count + ")";

            xlBasicWorksheet.get_Range("A" + count + 1, "B" + count + 1).Font.Bold = true;
        }

        private void TextConfig(Excel.Worksheet xlTextWorksheet, PowerPoint.Shape oPPTShape, int lastRowText, int slideCounter)
        {
            int red, green, blue;

            xlTextWorksheet.Cells[lastRowText + 1, 1] = slideCounter;
            xlTextWorksheet.Cells[lastRowText + 1, 2] = oPPTShape.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 3] = slideCounter + ":" + oPPTShape.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 4] = oPPTShape.TextFrame.TextRange.Text;
            xlTextWorksheet.Cells[lastRowText + 1, 5] = oPPTShape.TextFrame.TextRange.Font.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 6] = oPPTShape.TextFrame.TextRange.Font.Size;

            red = Color.FromArgb(oPPTShape.TextFrame.TextRange.Font.Color.RGB).B;
            green = Color.FromArgb(oPPTShape.TextFrame.TextRange.Font.Color.RGB).G;
            blue = Color.FromArgb(oPPTShape.TextFrame.TextRange.Font.Color.RGB).R;

            xlTextWorksheet.Cells[lastRowText + 1, 7] = "" + red + ", " + green + ", " + blue;
            //(xlWorkSheet.Cells[lastRow + 1, 6] as Excel.Range).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(red, green, blue));
            xlTextWorksheet.Cells[lastRowText + 1, 8] = Convert.ToInt16(oPPTShape.Left);
            xlTextWorksheet.Cells[lastRowText + 1, 9] = Convert.ToInt16(oPPTShape.Top);
            xlTextWorksheet.Cells[lastRowText + 1, 10] = Convert.ToInt16(oPPTShape.Height);
            xlTextWorksheet.Cells[lastRowText + 1, 11] = Convert.ToInt16(oPPTShape.Width);
            xlTextWorksheet.Cells[lastRowText + 1, 12] = oPPTShape.TextFrame.TextRange.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTriStateMixed ? "MIXED" : (oPPTShape.TextFrame.TextRange.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTrue ? "YES" : "NO");
            xlTextWorksheet.Cells[lastRowText + 1, 13] = oPPTShape.TextFrame.TextRange.Font.Italic == Microsoft.Office.Core.MsoTriState.msoTriStateMixed ? "MIXED" : (oPPTShape.TextFrame.TextRange.Font.Italic == Microsoft.Office.Core.MsoTriState.msoTrue ? "YES" : "NO");
        }

        private void ImageConfig(Excel.Worksheet xlTextWorksheet, PowerPoint.Shape oPPTShape, int lastRowText, int slideCounter)
        {
            xlTextWorksheet.Cells[lastRowText + 1, 1] = slideCounter;
            xlTextWorksheet.Cells[lastRowText + 1, 2] = oPPTShape.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 3] = slideCounter + ":" + oPPTShape.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 4] = Convert.ToInt16(oPPTShape.Left);
            xlTextWorksheet.Cells[lastRowText + 1, 5] = Convert.ToInt16(oPPTShape.Top);
            xlTextWorksheet.Cells[lastRowText + 1, 6] = Convert.ToInt16(oPPTShape.Height);
            xlTextWorksheet.Cells[lastRowText + 1, 7] = Convert.ToInt16(oPPTShape.Width);
        }

        private void ChartConfig(Excel.Worksheet xlChartWorksheet, PowerPoint.Shape oPPTShape, int lastRowChart, int slideCounter)
        {
            int seriesCount;
            string cellText = "";
            
            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Chart chart;
            chart = (Excel.Chart)temp_workbook.Charts.get_Item(1);

            seriesCount = (chart.SeriesCollection(Type.Missing) as Excel.SeriesCollection).Count;

            xlChartWorksheet.Cells[lastRowChart + 1, 1] = slideCounter;
            xlChartWorksheet.Cells[lastRowChart + 1, 2] = oPPTShape.Name;
            xlChartWorksheet.Cells[lastRowChart + 1, 3] = slideCounter + ":" + oPPTShape.Name;

            xlChartWorksheet.Cells[lastRowChart + 1, 4] = (chart.ChartType.ToString() == "-4111") ? "Line Chart" : chart.ChartType.ToString().Substring(2) + " Chart";

            if (chart.HasTitle)
                xlChartWorksheet.Cells[lastRowChart + 1, 5] = chart.ChartTitle.Text;
            else
                xlChartWorksheet.Cells[lastRowChart + 1, 5] = "[N/A]";

            if (chart.HasLegend)
                xlChartWorksheet.Cells[lastRowChart + 1, 6] = "[Y]";
            else
                xlChartWorksheet.Cells[lastRowChart + 1, 6] = "[N]";

            xlChartWorksheet.Cells[lastRowChart + 1, 7] = seriesCount;

            if (seriesCount == 1)
            {
                Excel.Series singleSeries = (Excel.Series)chart.SeriesCollection(1);
                Excel.Points ppp = (Excel.Points)singleSeries.Points(Type.Missing);                
                xlChartWorksheet.Cells[lastRowChart + 1, 8] = ppp.Count;
            }

            xlChartWorksheet.Cells[lastRowChart + 1, 9] = Convert.ToInt16(oPPTShape.Left);
            xlChartWorksheet.Cells[lastRowChart + 1, 10] = Convert.ToInt16(oPPTShape.Top);
            xlChartWorksheet.Cells[lastRowChart + 1, 11] = Convert.ToInt16(oPPTShape.Height);
            xlChartWorksheet.Cells[lastRowChart + 1, 12] = Convert.ToInt16(oPPTShape.Width);
            
            Excel.Axis valueAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            Excel.Axis categoryAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);


            if (!chart.ChartType.ToString().Contains("Pie"))
            {
                cellText = valueAxis.HasTitle ? valueAxis.AxisTitle.Text : "[N/A]";
                xlChartWorksheet.Cells[lastRowChart + 1, 13] = cellText;

                cellText = categoryAxis.HasTitle ? categoryAxis.AxisTitle.Text : "[N/A]";
                xlChartWorksheet.Cells[lastRowChart + 1, 15] = cellText;
            }

            try
            {
                xlChartWorksheet.Cells[lastRowChart + 1, 14] = "[" + valueAxis.MajorTickMark + " >> " + valueAxis.MinorTickMark + " >> " + valueAxis.TickLabelPosition + " >> " + valueAxis.MinimumScale + " >> " + valueAxis.MaximumScale +  " >> " + valueAxis.MinorUnit + " >> " + valueAxis.MajorUnit + " >> " + valueAxis.CrossesAt + "]";
            }
            catch (Exception ex)
            {
                errorMessage += "\nMessage From ChartConfig() --> Value Axis : " + ex.Message;
                xlChartWorksheet.Cells[lastRowChart + 1, 14] = "[N/A]";
            }
            
            
            try
            {
                xlChartWorksheet.Cells[lastRowChart + 1, 16] = "[" + categoryAxis.MajorTickMark + " >> " + categoryAxis.MinorTickMark + " >> " + categoryAxis.TickLabelPosition + " >> " + categoryAxis.AxisBetweenCategories + " >> " + categoryAxis.CrossesAt + "]";
            }
            catch (Exception ex)
            {
                errorMessage += "\nMessage From ChartConfig() --> Category Axis : " + ex.Message;
                xlChartWorksheet.Cells[lastRowChart + 1, 16] = "[N/A]";
            }
            
        }

        private void ChartDetailsConfig(Excel.Worksheet xlChartWorksheet, PowerPoint.Shape oPPTShape, int lastRowChart, int slideCounter)
        {
            int seriesCount;
            int red, green, blue;

            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Chart chart;
            chart = (Excel.Chart)temp_workbook.Charts.get_Item(1);

            seriesCount = (chart.SeriesCollection(Type.Missing) as Excel.SeriesCollection).Count;

            //xlChartWorksheet.Cells[lastRowChart + 1, 1] = slideCounter;
            //xlChartWorksheet.Cells[lastRowChart + 1, 2] = oPPTShape.Name;

            detailsCounter = 0;

            if (seriesCount == 1)
            {
                Excel.Series singleSeries = (Excel.Series)chart.SeriesCollection(1);
                Excel.Points ppp = (Excel.Points)singleSeries.Points(Type.Missing);

                foreach (Excel.Point p in ppp)
                {
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 1] = slideCounter;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 2] = oPPTShape.Name;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 3] = slideCounter + ":" + oPPTShape.Name + ":" + "Point " + (detailsCounter + 1);
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 4] = detailsCounter + 1;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 5] = p.DataLabel.Text + (detailsCounter + 1);

                    try
                    {
                        red = ColorTranslator.FromOle(Convert.ToInt32(p.Interior.Color)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(p.Interior.Color)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(p.Interior.Color)).B;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = "[" + red + ", " + green + ", " + blue + "]";
                    }
                    catch (Exception ex)
                    {
                        errorMessage += "\nMessage From ChartDetailsConfig() --> Points Color Configuration : " + ex.Message;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = "[N/A]";
                    }
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 7] = "[N/A]";
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 8] = "[N/A]";

                    detailsCounter++;
                }
            }
            else
            {
                Excel.SeriesCollection sC = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);

                foreach (Excel.Series s in sC)
                {
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 1] = slideCounter;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 2] = oPPTShape.Name;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 3] = slideCounter + ":" + oPPTShape.Name + ":" + s.Name;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 4] = detailsCounter + 1;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 5] = s.Name;
                    try
                    {
                        red = ColorTranslator.FromOle(s.Format.Line.ForeColor.RGB).R;
                        green = ColorTranslator.FromOle(s.Format.Line.ForeColor.RGB).G;
                        blue = ColorTranslator.FromOle(s.Format.Line.ForeColor.RGB).B;

                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = "[" + red + ", " + green + ", " + blue + "]";
                    }
                    catch (Exception ex)
                    {
                        errorMessage += "\nMessage From ChartDetailsConfig() --> Series Color Configuration : " + ex.Message;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = "[N/A]";
                    }

                    try
                    {
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 7] = s.MarkerStyle.ToString();

                        red = ColorTranslator.FromOle(s.MarkerBackgroundColor).R;
                        green = ColorTranslator.FromOle(s.MarkerBackgroundColor).G;
                        blue = ColorTranslator.FromOle(s.MarkerBackgroundColor).B;

                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 8] = "[" + red + ", " + green + ", " + blue + "]";
                    }
                    catch (Exception ex)
                    {
                        errorMessage += "\nMessage From ChartDetailsConfig() --> Series Marker Configuration : " + ex.Message;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 7] = "[N/A]";
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 8] = "[N/A]";
                    }

                    detailsCounter++;
                }
            }

            #region testing code for shape back ground data
            //Excel.Worksheet cWs = (Excel.Worksheet)temp_workbook.Worksheets.get_Item(1);
            ////getting the latest date of a chart
            //MessageBox.Show("Rows: " + cWs.UsedRange.Rows.Count + ", Columns: " + cWs.UsedRange.Columns.Count + ", last value: " + (cWs.Cells[3, 1] as Excel.Range).Text);
            #endregion
        }

        private void TableConfig(Excel.Worksheet xlTableWorksheet, PowerPoint.Shape oPPTShape, int lastRowTable, int slideCounter)
        {
            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Worksheet tableSheet;
            tableSheet = (Excel.Worksheet)temp_workbook.Worksheets.get_Item(1);

            xlTableWorksheet.Cells[lastRowTable + 1, 1] = slideCounter;
            xlTableWorksheet.Cells[lastRowTable + 1, 2] = oPPTShape.Name;
            xlTableWorksheet.Cells[lastRowTable + 1, 3] = slideCounter + ":" + oPPTShape.Name;
            xlTableWorksheet.Cells[lastRowTable + 1, 4] = tableSheet.UsedRange.Rows.Count;

            xlTableWorksheet.Cells[lastRowTable + 1, 5] = tableSheet.UsedRange.Columns.Count;

            xlTableWorksheet.Cells[lastRowTable + 1, 6] = Convert.ToInt16(oPPTShape.Left);
            xlTableWorksheet.Cells[lastRowTable + 1, 7] = Convert.ToInt16(oPPTShape.Top);
            xlTableWorksheet.Cells[lastRowTable + 1, 8] = Convert.ToInt16(oPPTShape.Height);
            xlTableWorksheet.Cells[lastRowTable + 1, 9] = Convert.ToInt16(oPPTShape.Width);
        }

        private void TableDetailsConfig(Excel.Worksheet xlTableWorksheet, PowerPoint.Shape oPPTShape, int lastRowTable, int slideCounter)
        {
            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Worksheet tableSheet;
            tableSheet = (Excel.Worksheet)temp_workbook.Worksheets.get_Item(1);            

            //GET First DATA CELL Address
            object objXLCell = new object();
            bool exitLoop = false;
            bool firstCell = false;

            int row = 0; // it will be starting row number of a data cell
            int col = 0; // it will be starting col number of a data cell

            for (row = 1; row < 20 && !exitLoop; row++)
            {
                firstCell = false;
                for (col = 1; col < 17; col++)
                {
                    objXLCell = (tableSheet.Cells[row, col] as Excel.Range).Value2;                    
                    if (objXLCell != null)
                    {
                        Type paramtype = objXLCell.GetType();
                        string name = paramtype.Name;
                        if (name == "Double" || name == "Integer")
                        {
                            if (!firstCell) continue;
                            else
                            {
                                exitLoop = true;
                                break;
                            }
                        }
                        firstCell = true;
                    }
                }
            }

            row = row - 1;
            ///////////////////////////////////////////////////////////////////


            int red, green, blue;

            //xlTableWorksheet.Cells[lastRowTable + 1, 1] = slideCounter;
            //xlTableWorksheet.Cells[lastRowTable + 1, 2] = oPPTShape.Name;

            detailsCounterTable = 0;
            //ROW HEADER
            for (int loopCol = 1; loopCol < col; loopCol++)
            {
                for (int j = row; j <= (tableSheet.Cells[1000, 1] as Excel.Range).get_End(Excel.XlDirection.xlUp).Row; j++)
                {
                    Excel.Range cell = tableSheet.Cells[j, loopCol] as Excel.Range;                    
                    objXLCell = cell.Value2;

                    if (objXLCell != null)
                    {
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 1] = slideCounter;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 2] = oPPTShape.Name;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 3] = slideCounter + ":" + oPPTShape.Name + ":" + cell.Row.ToString() + ":" + cell.Value2;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 4] = "ROW Header";
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 5] = "" + cell.Row.ToString();
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 6] = "" + cell.Value2;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 7] = "" + cell.RowHeight;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 8] = "[N/A]";
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 9] = "" + cell.Font.Name;

                        red = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color)).B;

                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 10] = "[" + red + ", " + green + ", " + blue + "]";
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 11] = "" + cell.Font.Size;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 12] = "" + cell.Font.Bold;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 13] = "" + cell.Font.Italic;

                        red = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color)).B;

                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 14] = "[" + red + ", " + green + ", " + blue + "]";

                        detailsCounterTable++;
                    }
                }
            }
            

            //COLUMN HEADER
            for (int loopRow = 1; loopRow < row; loopRow++)
            {
                for (int k = col; k <= tableSheet.UsedRange.Columns.Count; k++)
                {
                    Excel.Range cellCol = tableSheet.Cells[loopRow, k] as Excel.Range;
                    objXLCell = cellCol.Value2;
                    object numberFor = new object();

                    if (objXLCell != null)
                    {
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 1] = slideCounter;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 2] = oPPTShape.Name;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 3] = slideCounter + ":" + oPPTShape.Name + ":" + cellCol.Column.ToString() + ":" + cellCol.Value2;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 4] = "COLUMN Header";
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 5] = cellCol.Column.ToString();
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 6] = cellCol.Value2;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 7] = cellCol.ColumnWidth;

                        numberFor = (tableSheet.Cells[row, k] as Excel.Range).NumberFormat;
                        if (GetCellNumberFormat(tableSheet, k, row, numberFor))
                        {
                            xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 8] = numberFor;
                        }
                        else
                        {
                            xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 8] = "[MISMATCH]";
                        }
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 9] = "" + cellCol.Font.Name;

                        red = ColorTranslator.FromOle(Convert.ToInt32(cellCol.Font.Color)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(cellCol.Font.Color)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(cellCol.Font.Color)).B;

                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 10] = "[" + red + ", " + green + ", " + blue + "]";
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 11] = "" + cellCol.Font.Size;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 12] = "" + cellCol.Font.Bold;
                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 13] = "" + cellCol.Font.Italic;

                        red = ColorTranslator.FromOle(Convert.ToInt32(cellCol.Interior.Color)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(cellCol.Interior.Color)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(cellCol.Interior.Color)).B;

                        xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 14] = "[" + red + ", " + green + ", " + blue + "]";

                        detailsCounterTable++;
                    }
                }
            }
        }

        private bool GetCellNumberFormat(Excel.Worksheet tblSht, int col, int startRow, object numberFormat)
        {
            bool match = true;
            int endRow = (tblSht.Cells[1000, col] as Excel.Range).get_End(Excel.XlDirection.xlUp).Row;
            object cellNoFormat = new object();

            for (int i = startRow; i <= endRow; i++)
            {
                cellNoFormat = (tblSht.Cells[i, col] as Excel.Range).NumberFormat;
                if (cellNoFormat.ToString() == numberFormat.ToString()) continue;
                else
                    return false;
            }

            return match;
        }

        private void GetTableRowProperty(Excel.Worksheet tblSht)
        {
            object obj1 = new object();
            bool exitLoop = false;

            int row = 0;
            int col = 0;

            for (row = 1; row < 20 && !exitLoop; row++)
            {
                for (col = 1; col < 17; col++)
                {
                    obj1 = (tblSht.Cells[row, col] as Excel.Range).Value2;
                    if (obj1 != null)
                    {
                        Type paramtype = obj1.GetType();
                        string name = paramtype.Name;
                        if (name == "Double")
                        {
                            exitLoop = true;
                            break;
                        }
                    }
                }
            }

            MessageBox.Show(""+ row + " " + col);

        }

        #endregion
    }
}
