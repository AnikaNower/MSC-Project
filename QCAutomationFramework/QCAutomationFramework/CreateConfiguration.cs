using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

namespace QCAutomationFramework
{
    public class CreateConfiguration
    {
        private int detailsCounter = 0;
        private int detailsCounterTable = 0;

        public CreateConfiguration()
        {
        }

        public void BasicConfig(PowerPoint.Presentation oPres, Excel.Worksheet xlBasicWorksheet)
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

        public void TextConfig(Excel.Worksheet xlTextWorksheet, PowerPoint.Shape oPPTShape, int lastRowText, int slideCounter)
        {
            int red, green, blue;

            xlTextWorksheet.Cells[lastRowText + 1, 1] = slideCounter;
            xlTextWorksheet.Cells[lastRowText + 1, 2] = oPPTShape.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 3] = oPPTShape.TextFrame.TextRange.Text;
            xlTextWorksheet.Cells[lastRowText + 1, 4] = oPPTShape.TextFrame.TextRange.Font.Name;
            xlTextWorksheet.Cells[lastRowText + 1, 5] = oPPTShape.TextFrame.TextRange.Font.Size;

            red = Color.FromArgb(oPPTShape.TextFrame.TextRange.Font.Color.RGB).B;
            green = Color.FromArgb(oPPTShape.TextFrame.TextRange.Font.Color.RGB).G;
            blue = Color.FromArgb(oPPTShape.TextFrame.TextRange.Font.Color.RGB).R;

            xlTextWorksheet.Cells[lastRowText + 1, 6] = "" + red + ", " + green + ", " + blue;
            //(xlWorkSheet.Cells[lastRow + 1, 6] as Excel.Range).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(red, green, blue));
            xlTextWorksheet.Cells[lastRowText + 1, 7] = Convert.ToInt16(oPPTShape.Left);
            xlTextWorksheet.Cells[lastRowText + 1, 8] = Convert.ToInt16(oPPTShape.Top);
            xlTextWorksheet.Cells[lastRowText + 1, 9] = Convert.ToInt16(oPPTShape.Height);
            xlTextWorksheet.Cells[lastRowText + 1, 10] = Convert.ToInt16(oPPTShape.Width);
            xlTextWorksheet.Cells[lastRowText + 1, 11] = oPPTShape.TextFrame.TextRange.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTriStateMixed ? "MIXED" : (oPPTShape.TextFrame.TextRange.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTrue ? "YES" : "NO");
            xlTextWorksheet.Cells[lastRowText + 1, 12] = oPPTShape.TextFrame.TextRange.Font.Italic == Microsoft.Office.Core.MsoTriState.msoTriStateMixed ? "MIXED" : (oPPTShape.TextFrame.TextRange.Font.Italic == Microsoft.Office.Core.MsoTriState.msoTrue ? "YES" : "NO");
        }

        public void ChartConfig(Excel.Worksheet xlChartWorksheet, PowerPoint.Shape oPPTShape, int lastRowChart, int slideCounter)
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

            xlChartWorksheet.Cells[lastRowChart + 1, 3] = (chart.ChartType.ToString() == "-4111") ? "Line Chart" : chart.ChartType.ToString().Substring(2) + " Chart";

            if (chart.HasTitle)
                xlChartWorksheet.Cells[lastRowChart + 1, 4] = chart.ChartTitle.Text;
            else
                (xlChartWorksheet.Cells[lastRowChart + 1, 4] as Excel.Range).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(200, 0, 0));

            if (chart.HasLegend)
                xlChartWorksheet.Cells[lastRowChart + 1, 5] = "[Y]";
            else
                xlChartWorksheet.Cells[lastRowChart + 1, 5] = "[N]";

            xlChartWorksheet.Cells[lastRowChart + 1, 6] = seriesCount;

            if (seriesCount == 1)
            {
                Excel.Series singleSeries = (Excel.Series)chart.SeriesCollection(1);
                Excel.Points ppp = (Excel.Points)singleSeries.Points(Type.Missing);
                xlChartWorksheet.Cells[lastRowChart + 1, 7] = ppp.Count;
            }

            xlChartWorksheet.Cells[lastRowChart + 1, 8] = Convert.ToInt16(oPPTShape.Left);
            xlChartWorksheet.Cells[lastRowChart + 1, 9] = Convert.ToInt16(oPPTShape.Top);
            xlChartWorksheet.Cells[lastRowChart + 1, 10] = Convert.ToInt16(oPPTShape.Height);
            xlChartWorksheet.Cells[lastRowChart + 1, 11] = Convert.ToInt16(oPPTShape.Width);

            Excel.Axis valueAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            Excel.Axis categoryAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);


            if (!chart.ChartType.ToString().Contains("Pie"))
            {
                cellText = valueAxis.HasTitle ? valueAxis.AxisTitle.Text : "[N/A]";
                xlChartWorksheet.Cells[lastRowChart + 1, 12] = cellText;

                cellText = categoryAxis.HasTitle ? categoryAxis.AxisTitle.Text : "[N/A]";
                xlChartWorksheet.Cells[lastRowChart + 1, 14] = cellText;
            }

            try
            {
                xlChartWorksheet.Cells[lastRowChart + 1, 13] = "[" + valueAxis.MajorTickMark + " >> " + valueAxis.MinorTickMark + " >> " + valueAxis.TickLabelPosition + " >> " + valueAxis.MinimumScale + " >> " + valueAxis.MaximumScale + " >> " + valueAxis.MinorUnit + " >> " + valueAxis.MajorUnit + " >> " + valueAxis.CrossesAt + "]";
            }
            catch (Exception ex)
            {
                //errorMessage += "\nMessage From ChartConfig() --> Value Axis : " + ex.Message;
                xlChartWorksheet.Cells[lastRowChart + 1, 13] = "[N/A]";
            }


            try
            {
                xlChartWorksheet.Cells[lastRowChart + 1, 15] = "[" + categoryAxis.MajorTickMark + " >> " + categoryAxis.MinorTickMark + " >> " + categoryAxis.TickLabelPosition + " >> " + categoryAxis.AxisBetweenCategories + " >> " + categoryAxis.CrossesAt + "]";
            }
            catch (Exception ex)
            {
                //errorMessage += "\nMessage From ChartConfig() --> Category Axis : " + ex.Message;
                xlChartWorksheet.Cells[lastRowChart + 1, 15] = "[N/A]";
            }

        }

        public void ChartDetailsConfig(Excel.Worksheet xlChartWorksheet, PowerPoint.Shape oPPTShape, int lastRowChart, int slideCounter)
        {
            int seriesCount;
            int red, green, blue;

            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Chart chart;
            chart = (Excel.Chart)temp_workbook.Charts.get_Item(1);

            seriesCount = (chart.SeriesCollection(Type.Missing) as Excel.SeriesCollection).Count;

            xlChartWorksheet.Cells[lastRowChart + 1, 1] = slideCounter;
            xlChartWorksheet.Cells[lastRowChart + 1, 2] = oPPTShape.Name;

            detailsCounter = 0;

            if (seriesCount == 1)
            {
                Excel.Series singleSeries = (Excel.Series)chart.SeriesCollection(1);
                Excel.Points ppp = (Excel.Points)singleSeries.Points(Type.Missing);

                foreach (Excel.Point p in ppp)
                {
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 3] = detailsCounter + 1;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 4] = "Point " + (detailsCounter + 1);
                    try
                    {
                        red = ColorTranslator.FromOle(Convert.ToInt32(p.Interior.Color)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(p.Interior.Color)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(p.Interior.Color)).B;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 5] = "[" + red + ", " + green + ", " + blue + "]";
                    }
                    catch (Exception ex)
                    {
                        //errorMessage += "\nMessage From ChartDetailsConfig() --> Points Color Configuration : " + ex.Message;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 5] = "[N/A]";
                    }
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = "[N/A]";
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 7] = "[N/A]";

                    detailsCounter++;
                }
            }
            else
            {
                Excel.SeriesCollection sC = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);

                foreach (Excel.Series s in sC)
                {
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 3] = detailsCounter + 1;
                    xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 4] = s.Name;
                    try
                    {
                        red = ColorTranslator.FromOle(Convert.ToInt32(s.Interior.PatternColor)).R;
                        green = ColorTranslator.FromOle(Convert.ToInt32(s.Interior.PatternColor)).G;
                        blue = ColorTranslator.FromOle(Convert.ToInt32(s.Interior.PatternColor)).B;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 5] = "[" + red + ", " + green + ", " + blue + "]";
                    }
                    catch (Exception ex)
                    {
                        //errorMessage += "\nMessage From ChartDetailsConfig() --> Series Color Configuration : " + ex.Message;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 5] = "[N/A]";
                    }

                    try
                    {
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = s.MarkerStyle.ToString();

                        red = ColorTranslator.FromOle(s.MarkerBackgroundColor).R;
                        green = ColorTranslator.FromOle(s.MarkerBackgroundColor).G;
                        blue = ColorTranslator.FromOle(s.MarkerBackgroundColor).B;

                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 7] = "[" + red + ", " + green + ", " + blue + "]";
                    }
                    catch (Exception ex)
                    {
                        //errorMessage += "\nMessage From ChartDetailsConfig() --> Series Marker Configuration : " + ex.Message;
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 6] = "[N/A]";
                        xlChartWorksheet.Cells[lastRowChart + 1 + detailsCounter, 7] = "[N/A]";
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

        public void TableConfig(Excel.Worksheet xlTableWorksheet, PowerPoint.Shape oPPTShape, int lastRowTable, int slideCounter)
        {
            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Worksheet tableSheet;
            tableSheet = (Excel.Worksheet)temp_workbook.Worksheets.get_Item(1);

            xlTableWorksheet.Cells[lastRowTable + 1, 1] = slideCounter;
            xlTableWorksheet.Cells[lastRowTable + 1, 2] = oPPTShape.Name;
            xlTableWorksheet.Cells[lastRowTable + 1, 3] = tableSheet.UsedRange.Rows.Count;

            xlTableWorksheet.Cells[lastRowTable + 1, 4] = tableSheet.UsedRange.Columns.Count;

            xlTableWorksheet.Cells[lastRowTable + 1, 5] = Convert.ToInt16(oPPTShape.Left);
            xlTableWorksheet.Cells[lastRowTable + 1, 6] = Convert.ToInt16(oPPTShape.Top);
            xlTableWorksheet.Cells[lastRowTable + 1, 7] = Convert.ToInt16(oPPTShape.Height);
            xlTableWorksheet.Cells[lastRowTable + 1, 8] = Convert.ToInt16(oPPTShape.Width);
        }

        public void TableDetailsConfig(Excel.Worksheet xlTableWorksheet, PowerPoint.Shape oPPTShape, int lastRowTable, int slideCounter)
        {
            Excel._Workbook temp_workbook;
            temp_workbook = (Excel._Workbook)(oPPTShape.OLEFormat.Object);
            Excel.Worksheet tableSheet;
            tableSheet = (Excel.Worksheet)temp_workbook.Worksheets.get_Item(1);

            int red, green, blue;

            xlTableWorksheet.Cells[lastRowTable + 1, 1] = slideCounter;
            xlTableWorksheet.Cells[lastRowTable + 1, 2] = oPPTShape.Name;

            detailsCounterTable = 0;
            //ROW HEIGHT
            for (int j = 0; j < (tableSheet.Cells[1000, 1] as Excel.Range).get_End(Excel.XlDirection.xlUp).Row; j++)
            {
                Excel.Range cell = tableSheet.Cells[j + 1, 1] as Excel.Range;
                int columnCount = 1;

                while ((Convert.ToString(cell.Value2) == string.Empty || cell.Value2 == null) && columnCount < tableSheet.UsedRange.Columns.Count)
                {
                    cell = cell.Next;
                    columnCount++;
                }

                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 3] = "ROW";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 4] = "" + cell.Row.ToString();
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 5] = "" + cell.RowHeight;
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 6] = "" + cell.Font.Name;

                red = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color)).R;
                green = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color)).G;
                blue = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color)).B;

                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 7] = "[" + red + ", " + green + ", " + blue + "]";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 8] = "" + cell.Font.Size;
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 9] = "" + cell.Font.Bold;
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 10] = "" + cell.Font.Italic;

                red = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color)).R;
                green = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color)).G;
                blue = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color)).B;

                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 11] = "[" + red + ", " + green + ", " + blue + "]";

                detailsCounterTable++;
            }

            //COLUMN WIDTH
            for (int k = 0; k < tableSheet.UsedRange.Columns.Count; k++)
            {
                Excel.Range cellCol = tableSheet.Cells[1, k + 1] as Excel.Range;

                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 3] = "COL";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 4] = cellCol.Column.ToString();
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 5] = cellCol.ColumnWidth;
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 6] = "[N/A]";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 7] = "[N/A]";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 8] = "[N/A]";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 9] = "[N/A]";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 10] = "[N/A]";
                xlTableWorksheet.Cells[lastRowTable + 1 + detailsCounterTable, 10] = "[N/A]";

                detailsCounterTable++;
            }
        }

    }
}
