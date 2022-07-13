using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace CSV_import_export
{
    public partial class FileSizeForm : Form
    {
        public FileSizeForm()
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
        private void selectButton_Click(object sender, EventArgs e)
        {
            Stream myStream;
            OpenFileDialog openFileDialogCSV = new OpenFileDialog();

            openFileDialogCSV.InitialDirectory = Application.ExecutablePath.ToString();
            openFileDialogCSV.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialogCSV.FilterIndex = 2;
            openFileDialogCSV.Multiselect = true;
            openFileDialogCSV.RestoreDirectory = true;


            if (openFileDialogCSV.ShowDialog() == DialogResult.OK)
            {

                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                ds.Tables.Add(dt);
                dt.Columns.Add("S#", typeof(int));
                dt.Columns.Add("File Name", typeof(string));
                dt.Columns.Add("File Size", typeof(int));
                Int32 k = 0;
                Int32 i = 0;

                foreach (String file in openFileDialogCSV.FileNames)
                {
                    try
                    {

                        if ((myStream = openFileDialogCSV.OpenFile()) != null)
                        {

                            this.fileCSV = file;

                            System.IO.FileInfo fi = new System.IO.FileInfo(this.fileCSV);

                            this.dirCSV = fi.DirectoryName.ToString();

                            this.FileNevCSV = fi.Name.ToString();
                            Int64 size = new FileInfo(openFileDialogCSV.FileNames[k]).Length;
                            Int64 sizeKb = (size / 1024) + 1;

                            using (myStream)
                            {
                                i++;
                                dt.Rows.Add(i, FileNevCSV, sizeKb);

                                //for (int i=0; i < ; i++)
                                //{
                                //    DataRow dr = dt.NewRow();
                                //    dt.Rows.Add(dr);
                                //    for (int j=0; j <file.Length; j++)
                                //    {
                                //        DataColumn dc = new DataColumn();
                                //        dt.Columns.Add(dc);
                                //        if (j == 0)
                                //        {
                                //            dt.Rows[i][j] = fileNevCSV;
                                //        }
                                //        else 
                                //        {
                                //            dt.Rows[i][j] = ((size/1024)+1);
                                //        }

                                //    }
                                //}                         

                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    }
                    k++;
                }
                dataGridView1.DataSource = ds.Tables[0];

            }


        }


        /// <summary>
        /// New File size config
        /// </summary>

        private string newfileCSV;		//full file name
        private string newdirCSV;		//directory of file to import

        private string newfileNevCSV;	//name (with extension) of file to import - field

        public string NewFileNevCSV	//name (with extension) of file to import - property
        {
            get { return newfileNevCSV; }
            set { newfileNevCSV = value; }
        }
        private void newFileButton_Click(object sender, EventArgs e)
        {
            Stream newmyStream;
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.InitialDirectory = Application.ExecutablePath.ToString();
            ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.Multiselect = true;
            ofd.RestoreDirectory = true;


            if (ofd.ShowDialog() == DialogResult.OK)
            {

                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                ds.Tables.Add(dt);
                dt.Columns.Add("S#", typeof(int));
                dt.Columns.Add("File Name", typeof(string));
                dt.Columns.Add("File Size", typeof(int));
                Int32 k = 0;
                Int32 i = 0;

                foreach (String file in ofd.FileNames)
                {
                    try
                    {

                        if ((newmyStream = ofd.OpenFile()) != null)
                        {

                            this.newfileCSV = file;

                            System.IO.FileInfo fi = new System.IO.FileInfo(this.fileCSV);

                            this.newdirCSV = fi.DirectoryName.ToString();

                            this.NewFileNevCSV = fi.Name.ToString();
                            Int64 size = new FileInfo(ofd.FileNames[k]).Length;
                            Int64 sizeKb = (size / 1024) + 1;

                            using (newmyStream)
                            {
                                i++;
                                dt.Rows.Add(i, NewFileNevCSV, sizeKb);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    }
                    k++;
                }
                newDataGridView.DataSource = ds.Tables[0];

            }
        }

        private void diffButton_Click(object sender, EventArgs e)
        {
            diffDataGridView.Refresh();
            //DataSet ds = new DataSet();
            DataTable Table1 = (DataTable)(dataGridView1.DataSource);

            DataTable Table2 = (DataTable)(newDataGridView.DataSource);

            DataTable newDt = Subtract(Table1.DefaultView.ToTable("FirstTable"), Table2.DefaultView.ToTable("SecondTable").Copy());
            diffDataGridView.DataSource = newDt;
        }

        //public static int count = 0;
        public static DataTable Subtract(DataTable First, DataTable Second)
        {
            DataSet ds = new DataSet();
            ds.Tables.AddRange(new DataTable[] { First.Copy(), Second.Copy() });

            
            if (First.Columns.Count == Second.Columns.Count)
            {

                for (int i = 0; i < First.Rows.Count; i++)
                {
                    
                    for (int j = 2; j < Second.Columns.Count; j++)
                    {
                        
                        ds.Tables[0].Rows[i][j] = Convert.ToInt32(First.Rows[i][j].ToString()) - Convert.ToInt32(Second.Rows[i][j].ToString());
                        
                        
                    }
                }


            }

            return ds.Tables[0];
        }
    }

}