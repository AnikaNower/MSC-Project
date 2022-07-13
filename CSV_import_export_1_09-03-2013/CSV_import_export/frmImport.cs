using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace CSV_import_export
{
    public partial class frmImport : Form
    {
        public frmImport()
        {
            InitializeComponent();
        }

        private string fileCSV;		//full file name
        private string dirCSV;		//directory of file to import
        
        private string fileNevCSV;	//name (with extension) of file to import - field
        
        public string FileNevCSV	//name (with extension) of file to import - property
        {
            get { return fileNevCSV; }
            set { fileNevCSV = value; }
        }

        
        private string strFormat;	//CSV separator character

        private long rowCount = 0;	//row number of source file

        



        // Browses file with OpenFileDialog control

        private void btnFileOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogCSV = new OpenFileDialog();

            openFileDialogCSV.InitialDirectory = Application.ExecutablePath.ToString();
            openFileDialogCSV.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialogCSV.FilterIndex = 1;
            openFileDialogCSV.RestoreDirectory = true;
            

            if (openFileDialogCSV.ShowDialog() == DialogResult.OK)
            {
                this.txtFileToImport.Text = openFileDialogCSV.FileName.ToString();
                Int64 size = new FileInfo(openFileDialogCSV.FileName).Length;
                filesizeLabel.Text = "Imported file size: " + size + " Bytes / "+ ((size/1024) + 1) + " KB";
            }

        }



        // Delimiter character selection
        private void Format()
        {
            try
            {

                if (rdbSemicolon.Checked)
                {
                    strFormat = "Delimited(;)";
                }
                else if (rdbTab.Checked)
                {
                    strFormat = "TabDelimited";
                }
                else if (rdbSeparatorOther.Checked)
                {
                    strFormat = "Delimited(" + txtSeparatorOtherChar.Text.Trim() + ")";
                }
                else
                {
                    strFormat = "Delimited(;)";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Format");
            }
            finally
            {
            }
        }


        
        /*
         * Loads the csv file into a DataSet.
         */

        public DataSet LoadCSV()
        {
            DataSet ds = new DataSet();
            try
            {
                // Creates and opens an ODBC connection
                string strConnString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + this.dirCSV.Trim() + ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
                string sql_select;
                OdbcConnection conn;
                conn = new OdbcConnection(strConnString.Trim());
                conn.Open();
                if (numberTextBox.Text == "" && queryTextBox.Text == "")
                {
                    sql_select = "select * from [" + this.FileNevCSV.Trim() + "]";
                }

                //Creates the select command text
                //if (numberOfRows == -1)
                //{
                //    sql_select = "select * from [" + this.FileNevCSV.Trim() + "]";
                //}
                else
                //{
                //    int numberOfRows = Convert.ToInt32(numberTextBox.Text.ToString());
                //    sql_select = "select top " + numberOfRows + " * from [" + this.FileNevCSV.Trim() + "]";
                //}

                    //if (queryTextBox.Text != "")
                {
                    string query = queryTextBox.Text.ToString();
                    sql_select = query;
                    //+"[" + this.FileNevCSV.Trim() + "]";
                }
                OdbcDataAdapter obj_oledb_da = new OdbcDataAdapter(sql_select, conn);

                //Fills dataset with the records from CSV file
                obj_oledb_da.Fill(ds, "csv1");
                ds.Tables[0].Columns.Add("S#", typeof(Int32)).SetOrdinal(0);

                Int32 serial = 1;

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    dr["S#"] = serial++;
                }


                //Rows count
                this.rowCount = ds.Tables[0].Rows.Count;
                this.lblProgress.Text = "Imported: " + this.rowCount.ToString() + "/" + this.rowCount.ToString() + " row(s)";
                this.lblProgress.Refresh();
                //closes the connection
                conn.Close();
            }
            catch (Exception e) //Error
            {
                MessageBox.Show(e.Message, "Error - LoadCSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ds;
        }


        // Checks if a file was given.
        private bool fileCheck()
        {
            if ((fileCSV == "") || (fileCSV == null) || (dirCSV == "") || (dirCSV == null) || (FileNevCSV == "") || (FileNevCSV == null))
            {
                MessageBox.Show("Select a CSV file to load first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                return true;
            }
        }




        private void btnPreview_Click(object sender, EventArgs e)
        {
            dataGridView_preView.Refresh();
            loadPreview();
        }


        /* Loads the preview of CSV file in the DataGridView control.
         */

        private void loadPreview()
        {
            try
            {

                // select format, encoding, an write the schema file
                Format();
                this.dataGridView_preView.DataSource = LoadCSV();
                this.dataGridView_preView.DataMember = "csv1";

                //rowCountPreview();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error - loadPreview", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //public void rowCountPreview()
        // {

        //    DataSet ds = LoadCSV();
        //    // gets the number of rows
        //    this.rowCount = ds.Tables[0].Rows.Count;
        //    this.lblProgress.Text = "Imported: " + this.rowCount.ToString() + "/" + this.rowCount.ToString() + " row(s)";
        //    this.lblProgress.Refresh();

        //}

    

        private void tbFile_TextChanged(object sender, EventArgs e)
        {
            // full file name
            this.fileCSV = this.txtFileToImport.Text;

            // creates a System.IO.FileInfo object to retrive information from selected file.
            System.IO.FileInfo fi = new System.IO.FileInfo(this.fileCSV);
            // retrives directory
            this.dirCSV = fi.DirectoryName.ToString();
            // retrives file name with extension
            this.FileNevCSV = fi.Name.ToString();

            // database table name
            //this.txtTableName.Text = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length).Replace(" ", "_");
        }

        private void txtSeparatorOtherChar_TextChanged(object sender, EventArgs e)
        {
            this.rdbSeparatorOther.Checked = true;
        }


       
        
        
        
        
        
        
        //New file config
        private string newfileCSV;		//full file name
        private string newdirCSV;		//directory of file to import

        private string newfileNevCSV;	//name (with extension) of file to import - field

        public string NewFileNevCSV	//name (with extension) of file to import - property
        {
            get { return newfileNevCSV; }
            set { newfileNevCSV = value; }
        }


        private string strFormatNew;	//CSV separator character

        private long rowCountNew = 0;	//row number of source file





        // Browses file with OpenFileDialog control

        private void newFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogCSV = new OpenFileDialog();

            openFileDialogCSV.InitialDirectory = Application.ExecutablePath.ToString();
            openFileDialogCSV.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialogCSV.FilterIndex = 1;
            openFileDialogCSV.RestoreDirectory = true;

            if (openFileDialogCSV.ShowDialog() == DialogResult.OK)
            {
                this.newFileTextBox.Text = openFileDialogCSV.FileName.ToString();
            }

        }



        // Delimiter character selection
        private void FormatNew()
        {
            try
            {

                if (newRdbSemicolon.Checked)
                {
                    strFormatNew = "Delimited(;)";
                }
                else if (newRdbTab.Checked)
                {
                    strFormatNew = "TabDelimited";
                }
                else if (newRdbSeparatorOther.Checked)
                {
                    strFormatNew = "Delimited(" + newTxtSeparatorOtherChar.Text.Trim() + ")";
                }
                else
                {
                    strFormatNew = "Delimited(;)";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Format");
            }
            finally
            {
            }
        }



        /*
         * Loads the csv file into a DataSet.
         */

        public DataSet LoadCSVnew()
        {
            DataSet dsNew = new DataSet();
            try
            {
                // Creates and opens an ODBC connection
                string strConnStringNew = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + this.newdirCSV.Trim() + ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
                string sql_selectNew;
                OdbcConnection newConn;
                newConn = new OdbcConnection(strConnStringNew.Trim());
                newConn.Open();
                if (newNumberTextBox.Text == "" && newQueryTextBox.Text == "")
                {
                    sql_selectNew = "select * from [" + this.NewFileNevCSV.Trim() + "]";
                }

                //Creates the select command text
                //if (numberOfRows == -1)
                //{
                //    sql_selectNew = "select * from [" + this.FileNevCSV.Trim() + "]";
                //}
                else
                //{
                //    int numberOfRowsNew = Convert.ToInt32(newNumberTextBox.Text.ToString());
                //    sql_selectNew = "select top " + numberOfRowsNew + " * from [" + this.FileNevCSV.Trim() + "]";
                //}

                //if (newQueryTextBox.Text != "")
                {
                    string queryNew = newQueryTextBox.Text.ToString();
                    sql_selectNew = queryNew;
                    //+"[" + this.FileNevCSV.Trim() + "]";
                }
                OdbcDataAdapter obj_oledb_da = new OdbcDataAdapter(sql_selectNew, newConn);

                //Fills dataset with the records from CSV file
                obj_oledb_da.Fill(dsNew, "csv2");
                dsNew.Tables[0].Columns.Add("S#", typeof(Int32)).SetOrdinal(0);

                Int32 serial = 1;

                foreach (DataRow dr in dsNew.Tables[0].Rows)
                {
                    dr["S#"] = serial++;
                }


                //Rows count
                this.rowCountNew = dsNew.Tables[0].Rows.Count;
                this.lblProgressNew.Text = "Imported: " + this.rowCountNew.ToString() + "/" + this.rowCountNew.ToString() + " row(s)";
                this.lblProgressNew.Refresh();
                //closes the connection
                newConn.Close();
            }
            catch (Exception e) //Error
            {
                MessageBox.Show(e.Message, "Error - LoadCSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return dsNew;
        }


        // Checks if a file was given.
        //private bool fileCheckNew()
        //{
        //    if ((newfileCSV == "") || (newfileCSV == null) || (newdirCSV == "") || (newdirCSV == null) || (NewFileNevCSV == "") || (NewFileNevCSV == null))
        //    {
        //        MessageBox.Show("Select a CSV file to load first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}




        private void newLoadButton_Click(object sender, EventArgs e)
        {
            dataGridView_preViewNew.Refresh();
            loadPreviewNew();
        }


        /* Loads the preview of CSV file in the DataGridView control.
         */

        private void loadPreviewNew()
        {
            try
            {

                // select format, encoding, an write the schema file
                FormatNew();
                this.dataGridView_preViewNew.DataSource = LoadCSVnew();
                this.dataGridView_preViewNew.DataMember = "csv2";

                //rowCountPreview();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error - loadPreview", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //public void rowCountPreview()
        // {

        //    DataSet ds = LoadCSV();
        //    // gets the number of rows
        //    this.rowCount = ds.Tables[0].Rows.Count;
        //    this.lblProgress.Text = "Imported: " + this.rowCount.ToString() + "/" + this.rowCount.ToString() + " row(s)";
        //    this.lblProgress.Refresh();

        //}



        private void newFileTextBox_TextChanged(object sender, EventArgs e)
        {
            // full file name
            this.newfileCSV = this.newFileTextBox.Text;

            // creates a System.IO.FileInfo object to retrive information from selected file.
            System.IO.FileInfo fi = new System.IO.FileInfo(this.newfileCSV);
            // retrives directory
            this.newdirCSV = fi.DirectoryName.ToString();
            // retrives file name with extension
            this.NewFileNevCSV = fi.Name.ToString();

            // database table name
            //this.txtTableName.Text = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length).Replace(" ", "_");
        }

        private void newTxtSeparatorOtherChar_TextChanged(object sender, EventArgs e)
        {
            this.newRdbSeparatorOther.Checked = true;
        }
       

        //End of new File config
      
        private void subtractBtn_Click(object sender, EventArgs e)
        {
            subTractDataGridView.Refresh();
            DataSet ds = new DataSet();
            ds = (DataSet)(dataGridView_preView.DataSource);
            DataTable Table1 = ds.Tables[0];

            DataSet ds2 = new DataSet();
            ds2 = (DataSet)(dataGridView_preViewNew.DataSource);
            DataTable Table2 = ds2.Tables[0];

            DataTable newDt = Subtract(Table1.DefaultView.ToTable("FirstTable"), Table2.DefaultView.ToTable("SecondTable").Copy());
            subTractDataGridView.DataSource = newDt;

            foreach(DataGridViewRow dgr in subTractDataGridView.Rows)
            {
                for (int i = 1; i < dgr.Cells.Count; i++)
                {         
                        string d = dgr.Cells[i].Value.ToString();
                        //string str = d.Substring(1, 8);
                        if (d.Length > 10 && d.Substring(1, 8) == "Previous")
                        {
                            dgr.Cells[i].Style.BackColor = Color.Red;
                            
                        }
                        Int32 x;
                        if (Int32.TryParse(dgr.Cells[i].Value.ToString(), out x))
                        {
                            if (x > 0 || x < 0)
                            {
                                dgr.Cells[i].Style.BackColor = Color.Red;
                            }
                              
                        }
                        
                        subTractDataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                
            
            }
           
        }

        public static DataTable Subtract(DataTable First, DataTable Second)
        {
            DataSet ds = new DataSet();
            ds.Tables.AddRange(new DataTable[] { First.Copy(), Second.Copy() });

            //DataTable newDt = new DataTable();

            //ds.Tables.Add(newDt);
            if (First.Columns.Count == Second.Columns.Count)
            {
                
                for (int i = 0; i < First.Rows.Count; i++)
                {
                    //DataRow dr = newDt.NewRow();
                    //newDt.Rows.Add(dr);
                    for (int j = 1; j < Second.Columns.Count; j++)
                    {
                        

                        //DataColumn dc = new DataColumn();
                        //newDt.Columns.Add(dc);
                        double a;
                        
                        if (Double.TryParse(First.Rows[i][j].ToString(), out a))
                        {
                            First.Rows[i][j] = Convert.ToInt32(First.Rows[i][j]) - Convert.ToInt32(Second.Rows[i][j]);
                        }
                        else
                        {
                            if (First.Rows[i][j].ToString() == Second.Rows[i][j].ToString())
                            {
                                First.Rows[i][j] = Second.Rows[i][j].ToString();
                            }
                            else
                            {
                                First.Rows[i][j] = "[Previous Value: " + First.Rows[i][j].ToString() + "],[Current Value: " + Second.Rows[i][j].ToString() + "]";

                            }

                        }
                    
                    }
                }

             
            }

            return First;
        }

        private void differenceButton_Click(object sender, EventArgs e)
        {
            outputDataGridView.Refresh();
            DataSet ds = new DataSet();
            ds = LoadCSV();
            DataTable Table1 = ds.Tables[0];

            DataSet ds2 = new DataSet();
            ds2 = LoadCSVnew();
            DataTable Table2 = ds2.Tables[0];

            DataTable dataTable3 = Difference(Table1.DefaultView.ToTable("FirstTable"), Table2.DefaultView.ToTable("SecondTable").Copy());
            outputDataGridView.DataSource = dataTable3;

        }

        /*
        * 
        In summary the code works as follows:
        -------------------------------------

        Create new empty table
        Create a DataSet and add tables.
        Get a reference to all columns in both tables
        Create a DataRelation
        Using the DataRelation add rows with no child rows.
        Return table
        * 
        */



        public static DataTable Difference(DataTable First, DataTable Second)
        {

        //Create Empty Table

        DataTable table = new DataTable("Difference");



        //Must use a Dataset to make use of a DataRelation object

        using (DataSet ds = new DataSet())
        {

            //Add tables

            ds.Tables.AddRange(new DataTable[] { First.Copy(), Second.Copy() });

            //Get Columns for DataRelation

            DataColumn[] firstcolumns = new DataColumn[ds.Tables[0].Columns.Count];

            for (int i = 0; i < firstcolumns.Length; i++)
            {

                 firstcolumns[i] = ds.Tables[0].Columns[i];

            }



            DataColumn[] secondcolumns = new DataColumn[ds.Tables[1].Columns.Count];

            for (int i = 0; i < secondcolumns.Length; i++)
            {

                secondcolumns[i] = ds.Tables[1].Columns[i];

            }

            //Create DataRelation

            DataRelation r = new DataRelation(string.Empty, firstcolumns, secondcolumns, false);

            ds.Relations.Add(r);



            //Create columns for return table

            for (int i = 0; i < First.Columns.Count; i++)
            {

                table.Columns.Add(First.Columns[i].ColumnName, First.Columns[i].DataType);

            }



            //If First Row not in Second, Add to return table.

                table.BeginLoadData();

            foreach (DataRow parentrow in ds.Tables[0].Rows)
            {

                DataRow[] childrows = parentrow.GetChildRows(r);

                if (childrows == null || childrows.Length == 0)

                table.LoadDataRow(parentrow.ItemArray, true);

            }

         table.EndLoadData();

        }

        return table;

        }

       
  
      
    }  
}