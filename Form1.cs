using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics.Eventing.Reader;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography;
using System.Xml.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using System.IO;
using ClosedXML.Excel;
using OfficeOpenXml;
using System.ComponentModel.Composition.Primitives;
//using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Vml;
using System.Runtime.ConstrainedExecution;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Globalization;
//using OfficeOpenXml.Core.ExcelPackage;




namespace master_buku
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        public string ConnectionString;
        public SqlConnection conn;

        public string thismonth = DateTime.Today.ToString("yyyyMM") ;
        public int lok = 8;      
        
        public string folderPath = "D:\\Master Data\\";


        private void button1_Click(object sender, EventArgs e)
        {
            string str = Clipboard.GetText();
            string[] pindahbaris = str.Split('\r');
            str = "";

            for (int i = 0; i < pindahbaris.Length; i++)
            {
                  if (pindahbaris[i].Length >2)//  if (pindahbaris[i] != "")
                    listBox1.Items.Add(pindahbaris[i]);
             
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0)
            {
              if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.Rows.Clear();
                }

                if (radioButton1.Checked)
                {
                    PindahData();
                    setHeaderMaster();
                    list_items();
                    count_isbn();
                }
                else if(radioButton2.Checked)
                {
                    PindahData(); 
                    setHeaderStock();
                    CariStock();
                }
                else
                {
                    MessageBox.Show("Silakan Klik Pilihan antara Mencari Stok atau Master");
                }

            }
            else
            {
                MessageBox.Show("Masukkan ISBN untuk cek master atau Kode untuk cek Stock");

            }

        }


        private void setHeaderStock()
        {
            dataGridView1.Columns[0].HeaderCell.Value = "KODE";
            dataGridView1.Columns[1].HeaderCell.Value = "ISBN";
            dataGridView1.Columns[2].HeaderCell.Value = "JUDUL";
            dataGridView1.Columns[3].HeaderCell.Value = "BINLOC";
            dataGridView1.Columns[4].HeaderCell.Value = "SGD";
            dataGridView1.Columns[5].HeaderCell.Value = "BULKY";
            dataGridView1.Columns[6].HeaderCell.Value = "SGB";
            dataGridView1.Columns[7].HeaderCell.Value = "TOKO";
        }



        private void setHeaderMaster()
        {

            dataGridView1.Columns[0].HeaderCell.Value = "ISBN";
            dataGridView1.Columns[1].HeaderCell.Value = "KODE";
            dataGridView1.Columns[2].HeaderCell.Value = "JUDUL";
            dataGridView1.Columns[3].HeaderCell.Value = "KAT";
            dataGridView1.Columns[4].HeaderCell.Value = "STATUS";
            dataGridView1.Columns[5].HeaderCell.Value = "TGL MASUK";
            dataGridView1.Columns[6].HeaderCell.Value = "TGL EDIT";
            dataGridView1.Columns[7].HeaderCell.Value = "DBL";

        }




        private void CariStock()
        {

            if (dataGridView1.Rows.Count > 0)
            {
                conn.Open();

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i = i + 1)
                {

                    //mengko ditambah jml kode sing isbn podho

                    string str = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                    dataGridView1.Rows[i].Cells[0].Value = str;
                    SqlCommand cmd = new SqlCommand("Select left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "[DESC] judul, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select format(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset, '###,###') from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select format(iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset, '###,###') from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "TOKO = FORMAT((select SUM(iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                 "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION NOT LIKE 'JK%'), '###,###')" +

                 "FROM ICITEM " +
                  "  where OPTFLD4 = '" + str + "' and inactive=0", conn);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        //SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("select optfld4 from icitem where itemno ='" + str + "'", conn);
                        //DataSet ds = new DataSet();


                        // if (ds.Tables.Count >0 )
                        if (reader.Read())
                        {
                            // textBox1.Text = reader[0].ToString();
                            dataGridView1.Rows[i].Cells[1].Value = reader["ISBN"].ToString().Trim();
                            dataGridView1.Rows[i].Cells[2].Value = reader["JUDUL"].ToString();
                            dataGridView1.Rows[i].Cells[3].Value = reader["BINLOC"].ToString();
                            dataGridView1.Rows[i].Cells[4].Value = reader["SGD"].ToString().Trim();
                            dataGridView1.Rows[i].Cells[5].Value = reader["BULKY"].ToString();
                            dataGridView1.Rows[i].Cells[6].Value = reader["SGB"].ToString();
                            dataGridView1.Rows[i].Cells[7].Value = reader["TOKO"].ToString();
                            
                        }
                    }

                }
           
                conn.Close();

            }

     


            
        }


        public static void SaveToExcel(DataSet dataset, String excelFilePath)
        {
            using(SpreadsheetDocument document =  SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                Sheets sheets = workbookpart.Workbook.AppendChild(new Sheets());

                foreach(DataTable table in dataset.Tables)
                {
                    UInt32Value sheetCount = 0;
                    sheetCount++;

                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    Sheet sheet = new Sheet() { Id = workbookpart.GetIdOfPart(worksheetPart), SheetId = sheetCount, Name = table.TableName };
                    sheets.AppendChild(sheet);

                    Row headerRow = new Row();

                    List<String> columns = new List<string>();
                    foreach ( System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();

                        foreach (String col in columns)
                            {
                                Cell cell = new Cell();
                                cell.DataType = CellValues.String;
                                cell.CellValue = new CellValue(dsrow[col].ToString());
                                newRow.AppendChild(cell);
                            }
                    
                        sheetData.AppendChild(newRow);
                    }
                     
                }
                
                workbookpart.Workbook.Save();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
                dataGridView2.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                progressBar1.Visible = true;
                progressBar1.Value = 99;

                if (lok == 0)
                {
                    AllActive();
                }
                else if (lok == 1)
                {
                    AllActiveExp();
                }
                else if (lok == 2)
                {
                    AllInactiveExp();
                }
                else if (lok == 3)
                {
                    ThisMonthNewExp();
                }
                else if (lok == 4)
                {
                    ThisMonthUpdateExp();
                }
                else if (lok == 5)
                {
                    ThisMonthInactiveExp();
                }
                else if (lok == 6)
                {
                    UpdateDataPeriodExp();
                }
                else {
            
                 UpdateInactivePeriodExp();

                }

            
        }

        public void AllDataExp()
        {
            lok = 0;
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                  "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
               "FROM ICITEM " +
                 "where left(optfld3,2)<>'09' " +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            var usave = folderPath + "Master List per " + DateTime.Today.ToString("dd-MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

          
            conn.Close();

            dataGridView2.Enabled = true;

            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export Berhasil, silakan cek " + usave);

        }

        public void AllActiveExp()
        {
            lok = 1;
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                  "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where left(optfld3,2)<>'09' and inactive=0 " +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            var usave = folderPath + "Master List Active per " + DateTime.Today.ToString("dd-MMM-yyyy") + ".xlsx" ;

            SaveToExcel(ds, usave);


            conn.Close();
            dataGridView2.Enabled = true;
        
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);

        }


        public void AllInactiveExp()
        {
            lok = 2;
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                   "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                "FROM ICITEM " +
                 "where left(optfld3,2)<>'09'  and inactive=1" +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            var usave = folderPath + "Master Inactive per " + DateTime.Today.ToString("dd-MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

            conn.Close();

            dataGridView2.Enabled = true;
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);


            
        }


        public void ThisMonthNewExp()
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where left(optfld3,2)<>'09'  and inactive=1" +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            var usave = folderPath + "Master Baru per " + DateTime.Today.ToString("MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

            conn.Close();

            conn.Close();
            dataGridView2.Enabled = true;
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);

        }

        public void ThisMonthUpdateExp()
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where left(optfld3,2)<>'09'  and inactive=0 " +
                 "and optdate<> datelastmn " +
                 "and datelastmn like '" + thismonth + "%'" +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);


            var usave = folderPath + "Master Update Active  per " + DateTime.Today.ToString("MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

            conn.Close();

            conn.Close();
            dataGridView2.Enabled = true;
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);

     
        }

        public void ThisMonthInactiveExp()
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where left(optfld3,2)<>'09'  and inactive=1 " +
                 "and optdate<> datelastmn " +
                 "and datelastmn like '" + thismonth + "%'" +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);


            var usave = folderPath + "Master Update Inactive  per " + DateTime.Today.ToString("MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

            conn.Close();

            conn.Close();
            dataGridView2.Enabled = true;
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);

        }


        public void UpdateDataPeriodExp()
        {
            String datea = dateTimePicker1.Value.ToString("yyyyMMdd");
            String dateb = dateTimePicker2.Value.ToString("yyyyMMdd");

            int datefirst = Convert.ToInt32(datea);
            int datelast = Convert.ToInt32(dateb);
            lok = 6;
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                "where icitem.INACTIVE = 0 and " +
                "left(optfld3, 2) <> '09' " +
                "and audtdate<> optdate  " +
                "and audtdate >= '" + datefirst + "' " +
                "and audtdate <= '" + datelast + " '" +
                "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);


            var usave = folderPath + "Master Update per " + dateTimePicker1.Value.ToString("dd-MMM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

            conn.Close();

            conn.Close();
            dataGridView2.Enabled = true;
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);


          
        }


        public void UpdateInactivePeriodExp()
        {

            String datea = dateTimePicker1.Value.ToString("yyyyMMdd");
            String dateb = dateTimePicker2.Value.ToString("yyyyMMdd");

            int datefirst = Int32.Parse(datea);
            int datelast = Int32.Parse(dateb);
          
            lok = 6;
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                "where icitem.INACTIVE = 1 and " +
                "left(optfld3, 2) <> '09' " +
                "and audtdate<> optdate  " +
                "and audtdate >= '" + datefirst + "' " +
                "and audtdate <= '" + datelast + " '" +
                "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);


            var usave = folderPath + "Master Inactive per " + dateTimePicker1.Value.ToString("dd-MMM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MMM-yyyy") + ".xlsx";

            SaveToExcel(ds, usave);

            conn.Close();

            conn.Close();
            dataGridView2.Enabled = true;
            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;

            MessageBox.Show("Export berhasil, silakan cek " + usave);



           

        }

        private void list_items()
        {


            if (dataGridView1.Rows.Count > 0)
            {
                conn.Open();

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i = i + 1)
                {

                    //mengko ditambah jml kode sing isbn podho

                    string str = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                    SqlCommand cmd = new SqlCommand("Select optfld4, [desc] judul, optfld3, inactive, OPTDATE, datelastmn from icitem where itemno = '" + str + "' and left(optfld3,2)<>'09'", conn);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        //SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("select optfld4 from icitem where itemno ='" + str + "'", conn);
                        //DataSet ds = new DataSet();


                        // if (ds.Tables.Count >0 )
                        if (reader.Read())
                        {
                            // textBox1.Text = reader[0].ToString();
                            dataGridView1.Rows[i].Cells[1].Value = reader["optfld4"].ToString().Trim();
                            dataGridView1.Rows[i].Cells[2].Value = reader["judul"].ToString();
                            dataGridView1.Rows[i].Cells[3].Value = reader["optfld3"].ToString();
                            dataGridView1.Rows[i].Cells[4].Value = reader["inactive"].ToString().Trim();
                            dataGridView1.Rows[i].Cells[5].Value = reader["optdate"].ToString();
                            dataGridView1.Rows[i].Cells[6].Value = reader["datelastmn"].ToString();
                        }
                        else
                        {
                            dataGridView1.Rows[i].Cells[2].Value = "ISBN belum terdaftar";
                        }
                    }

                }
                conn.Close();

            }
        }

        private void count_isbn()
        {
            conn.Open();

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i = i + 1)
            {
                if (dataGridView1.Rows[i].Cells[2].ToString() != "")
                {
                    string str = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                    SqlCommand cmd = new SqlCommand("Select count(*) from icitem where itemno = '" + str + "'", conn);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                      

                        if (Convert.ToInt32(reader.Read()) > 1 )
                        {
                            dataGridView1.Rows[i].Cells[7].Value = "Yes";
                        }
                        else
                        {

                            dataGridView1.Rows[i].Cells[7].Value = "No";
                        }

                    }
                }
            }

            conn.Close() ;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            ConnectionString = @"Data Source=10.10.10.215;Initial Catalog=brdjkt;Persist Security Info=True;User ID=sa;Password=kentudsapi;Integrated Security=False;"; //10.10.10.215

            conn = new SqlConnection(ConnectionString);
            conn.Close();

            //this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma",11);
            label1.Visible = false;
            label2.Visible = false;
            dateTimePicker2.Visible = false;
            dateTimePicker1.Visible = false;
            dateTimePicker1.Value = DateTime.Now.AddDays(-30);
            dateTimePicker2.Value = DateTime.Now;
            progressBar1.Visible = false;

            
            if (!Directory.Exists(folderPath))
            {
              Directory.CreateDirectory(folderPath);
            }
        }


        private void hitung_isbn()
        {
            conn.Open();

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i = i + 1)
            {
                //mengko ditambah jml kode sing isbn podho

                string std = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                SqlCommand cmd = new SqlCommand("Select sum(*) from icitem where itemno = '" + std + "' and left(optfld3,2)<>'09'", conn);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        // textBox1.Text = reader[0].ToString();
                        dataGridView1.Rows[i].Cells[7].Value = reader.ToString().Trim();

                    }
                }

            }
            conn.Close();

        }

        private void PindahData()
        {
            for ( int i = 0;i < listBox1.Items.Count ;i++) {
                dataGridView1.Rows.Add(listBox1.Items[i]);                
            }
            listBox1.Items.Clear();
            dataGridView1.Columns[0].Width = 130;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.BringToFront();
            progressBar1.Value = 95;
            dataGridView2.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;


            if (comboBox1.SelectedIndex == 0)
            {
                AllData();
            } 
            else if (comboBox1.SelectedIndex == 1)
            {
                AllActive();
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                AllInactive();
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                ThisMonthNew();
            }
            else if (comboBox1.SelectedIndex == 4)
            {
                ThisMonthUpdate();
            }
            else if (comboBox1.SelectedIndex == 5)
            {
               ThisMonthInactive();
            }
            else if (comboBox1.SelectedIndex == 6)
            {
                UpdateDataPeriod();
            }
            else if (comboBox1.SelectedIndex == 7)
            {
                UpdateInactivePeriod();
            }
            else 
            {
                MessageBox.Show("Silakan Pilih Jenis List");
                return;
            }


            hitungstock();
            SetDG2();
            setPCRF();

            dataGridView2.Enabled = true;

            progressBar1.Visible = false;
            button6.Enabled = true;
            button5.Enabled = true;


        }

        private void AllData()
        {
            lok = 0;

            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where left(optfld3,2)<>'09' " +
                 "order by audtdate asc", conn);

            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];

            
            conn.Close();
            
        }

        private void AllActive()
        {
            lok = 1;

            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                   "where icitem.INACTIVE=0 and " +
                "left(optfld3,2)<>'09' " +
                "order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

          
            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();


        }

        private void AllInactive()
        {
            lok = 2;


            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 " FROM ICITEM " +
                   "where icitem.INACTIVE = 1 and " +
                "left(optfld3,2)<>'09' " +
                "order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();

            

        }


        private void ThisMonthInactive()
        {
            lok = 5;



            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where icitem.INACTIVE=1 and " +
                "left(optfld3,2)<>'09' " +
                "and datelastmn like '" + thismonth + "%'" +
                "order by audtdate asc", conn);

            //            SqlCommand cmd = new SqlCommand("Select itemno ISBN, OPTFLD4 KODE, [desc] JUDUL, optfld3 KAT, inactive STATUS, optdate TGL_AWAL, datelastmn TGL_UPDATE from icitem where inactive = 1 and left(optfld3,2)<>'09' and datelastmn like '" + thismonth + "%' order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();

          

        }

        private void ThisMonthNew()
        {
            lok = 3;



            conn.Open();

            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                "where icitem.INACTIVE=0 and " +
                "left(optfld3,2)<>'09' " +
                "and optdate like '" + thismonth + "%'" +
                "order by audtdate asc", conn);

            //            SqlCommand cmd = new SqlCommand("Select itemno ISBN, OPTFLD4 KODE, [desc] JUDUL, optfld3 KAT, inactive STATUS, optdate TGL_AWAL, datelastmn TGL_UPDATE from icitem where left(optfld3,2)<>'09' and optdate like'" + thismonth + "%' order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();

        }


        private void ThisMonthUpdate()
        {
            lok = 4;


            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                "FROM ICITEM " +
                "where icitem.INACTIVE=0 and " +
                "left(optfld3,2)<>'09' " +
                "and optdate<> datelastmn " +
                "and datelastmn like '" + thismonth + "%'" +
                "order by audtdate asc", conn);

            //            SqlCommand cmd = new SqlCommand("Select itemno ISBN, OPTFLD4 KODE, [desc] JUDUL, optfld3 KAT, inactive STATUS, optdate TGL_AWAL, datelastmn TGL_UPDATE from icitem where left(optfld3,2)<>'09' and optdate <> datelastmn and datelastmn like'" + thismonth + "%' order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();

       

        }

        private void UpdateDataPeriod()
        {
            String datea = dateTimePicker1.Value.ToString("yyyyMMdd");
            String dateb = dateTimePicker2.Value.ToString("yyyyMMdd");

            int datefirst = Convert.ToInt32(datea);
            int datelast = Convert.ToInt32(dateb);
            lok = 6;


            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
                "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                 "FROM ICITEM " +
                 "where icitem.INACTIVE = 0 and " +
                "left(optfld3, 2) <> '09' " +
                "and audtdate<> optdate  " +
                "and audtdate >= '" + datefirst + "' " +
                "and audtdate <= '" + datelast + " '" +
                "order by audtdate asc", conn);

            //            SqlCommand cmd = new SqlCommand("Select itemno ISBN, OPTFLD4 KODE, [desc] JUDUL, optfld3 KAT, inactive STATUS, optdate TGL_AWAL, datelastmn TGL_UPDATE from icitem where audtdate <> optdate and audtdate>='" + datefirst + "' and audtdate <= '" + datelast + "' and left(optfld3,2)<>'09' order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();



        }

        private void UpdateInactivePeriod()
        {
            String datea = dateTimePicker1.Value.ToString("yyyyMMdd");
            String dateb = dateTimePicker2.Value.ToString("yyyyMMdd");

            int datefirst = Int32.Parse(datea);
            int datelast = Int32.Parse(dateb);
            lok = 7;



            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT " +
            "left(OPTFLD4, 7) CODE, " +
                "left(icitem.ITEMNO, 16) ISBN, " +
                "left(FMTITEMNO,17) BARCODE, " +
                "right(left(itemno, 16), 13)[ISBN - 9], " +
                "OPTFLD4 KODE, " +
                "SALEPRICE = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'std'), " +
                "MPR = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mpr'), " +
                "MP2 = (select SALEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist = 'mp2'), " +
                "optfld2 PRODID,  " +
                "optfld2 PCFZ,  " +
                "PUBLISHER = (select [data] from csopt where icitem.optfld5 = csopt.code ), " +
                "comment1 AUTHOR, " +
                "[DESC], " +
                "BASEPRICE = (select BASEPRICE from icpric where icpric.itemno = ICITEM.itemno and pricelist= 'std'), " +
                "LEFT(OPTFLD6, 3) CUR, " +
                "CVR_PRICE = case when ( charindex('/',OPTFLD6,1)-5)>0 then substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))-5) else substring(ltrim(optfld6),4, ( charindex('/',OPTFLD6,1))) end ," +
                "NET_PRICE = (select substring(ltrim(optfld6), ( charindex('/',OPTFLD6,1))+5,len(optfld6)+1))," +
                "CATEGORY = (select LEFT(CODE,5) + [DATA] from CSOPT where csopt.code = icitem.optfld3), " +
                "optfld1 STATUS, " +
                "ICITEM.PICKINGSEQ BINLOC, " +
                "SGD = (select(iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300'), " +
                "icitem.comment4 BULKY, " +
                "SGB = (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                    "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'), " +
                "Total = (select (iciloc.qtyrenocst - ICILOC.qtyshnocst + ICILOC.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   "where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk300') + " +
                   " (select (iciloc.qtyrenocst - iciloc.qtyshnocst + iciloc.qtyadnocst + iciloc.qtyoffset) from iciloc " +
                   " where iciloc.itemno = ICITEM.itemno and iciloc.LOCATION = 'jk301'),  " +
                "OPTFLD5 PUBLISHER," +
                "OPTFLD6 [CVR-NETPRICE]," +
                 "optdate TGL_AWAL, " +
                 "datelastmn TGL_UPDATE " +
                     "FROM ICITEM " +
                           "where icitem.INACTIVE=1 and " +
                  "left(optfld3,2)<>'09'  " +
                  " and audtdate<> optdate " +
                  "and audtdate >= '" + datefirst + "' " +
                  "and audtdate <= '" + datelast + " '" +
                  "order by audtdate asc", conn);

            //            SqlCommand cmd = new SqlCommand("Select itemno ISBN, OPTFLD4 KODE, [desc] JUDUL, optfld3 KAT, inactive STATUS, optdate TGL_AWAL, datelastmn TGL_UPDATE from icitem where inactive= 1 and left(optfld3,2)<>'09'  and audtdate <> optdate and audtdate>='" + datefirst + "' and audtdate<='" + datelast + "' order by audtdate asc", conn);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();

        }


        private void SetDG2() {
            this.dataGridView2.Columns[0].Width = 100; //kode
      
            this.dataGridView2.Columns[1].Width = 130; //isbn
            
            this.dataGridView2.Columns[2].Width = 130; //barcode
            
            this.dataGridView2.Columns[3].Width = 100; //isbn9
            
            this.dataGridView2.Columns[4].Width = 90;//code
            
            this.dataGridView2.Columns[5].Width = 100;  //saleprice
            this.dataGridView2.Columns[5].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[6].Width = 100;  //mpr
            this.dataGridView2.Columns[6].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[7].Width = 100; //mp2
            this.dataGridView2.Columns[7].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[8].Width = 60; //prodid

            this.dataGridView2.Columns[9].Width = 120; //pczf

            this.dataGridView2.Columns[10].Width = 120; //publisher

            this.dataGridView2.Columns[11].Width = 120; // author

            this.dataGridView2.Columns[12].Width = 160; //desc

            this.dataGridView2.Columns[13].Width = 100; //baseprice
            this.dataGridView2.Columns[13].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[14].Width = 90; //curr

            this.dataGridView2.Columns[15].Width = 100; //cvrprice
            this.dataGridView2.Columns[15].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[16].Width = 100; //netprice
            this.dataGridView2.Columns[16].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[17].Width = 100; //cat

            this.dataGridView2.Columns[18].Width = 100; //status


            this.dataGridView2.Columns[19].Width = 140; //binloc

            this.dataGridView2.Columns[20].Width = 100; //sgd
            this.dataGridView2.Columns[20].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[21].Width = 90; //bulk

            this.dataGridView2.Columns[22].Width = 100; //sgb
            this.dataGridView2.Columns[22].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[22].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[23].Width = 100; //total
            this.dataGridView2.Columns[23].DefaultCellStyle.Format = "#,##0";
            this.dataGridView2.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            this.dataGridView2.Columns[24].Width = 100; //publisher


            this.dataGridView2.Columns[25].Width = 100; //cover-net

            this.dataGridView2.Columns[26].Width = 100; //tgl awal

//            this.dataGridView2.Columns[27].Width = 100; //tgl update

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 6 ||  comboBox1.SelectedIndex == 7){
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                label1.Visible = true;
                label2.Visible = true;
            }
            else
            { 
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
            }
        }


        private void setPCRF()
        {
            for (int i = 0; i < dataGridView2.Rows.Count-1; i++)
            {

                if (dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "A" || dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "AS" || dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "B" || dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "BS")
                {
                    dataGridView2.Rows[i].Cells[9].Value = "P";

                }


                else if (dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "C" || dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "CS" || dataGridView2.Rows[i].Cells[8].Value.ToString().Trim() == "H")
                {
                    dataGridView2.Rows[i].Cells[9].Value = "F";

                }
                else
                {
                    dataGridView2.Rows[i].Cells[9].Value = "C";

                }
            }
        }

        private void hitungstock()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
               if (dataGridView2.Rows[i].Cells[20].Value == DBNull.Value) 
                {
                   dataGridView2.Rows[i].Cells[20].Value = "0";

                }

                if  (dataGridView2.Rows[i].Cells[20].ToString() == "")
                {
                    dataGridView2.Rows[i].Cells[20].Value = "0";

                }


               if (dataGridView2.Rows[i].Cells[22].Value == DBNull.Value)
                {
                    dataGridView2.Rows[i].Cells[22].Value = "0";
                }

                if (dataGridView2.Rows[i].Cells[22].ToString() == "")
                {
                    dataGridView2.Rows[i].Cells[22].Value = "0";

                }

                int a = Convert.ToInt32(dataGridView2.Rows[i].Cells[20].Value);
                int b = Convert.ToInt32(dataGridView2.Rows[i].Cells[22].Value);


                dataGridView2.Rows[i].Cells[23].Value = a +b;

            }
        }

        public static DataTable DataGridView_to_dt(DataGridView dv)
        {
            DataTable ExportDataTable = new DataTable();

            foreach (DataGridViewColumn col in dv.Columns)
            {
                ExportDataTable.Columns.Add(col.Name);
            }
            foreach (DataGridViewRow row in dv.Rows)
            {
                DataRow drow = ExportDataTable.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    drow[cell.ColumnIndex] = cell.Value;
                }
                ExportDataTable.Rows.Add(drow);
            }
            return ExportDataTable;
        }



        private void button4_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count > 0)
            { 
                DataTable dt = DataGridView_to_dt(dataGridView1);
                var name1 = dt.Rows[0][1];

                DataSet ds = new DataSet();
                //ds.Tables.Add(dt);

                //ExportDataSet(ds, "D:\\Cek_Master.xlsx");

                var wbook = new XLWorkbook();

                if (radioButton2.Checked)
                {
                    var ws = wbook.AddWorksheet("Stock");

                    ws.Cell("A1").Value = "Cek Stock Item";
                    ws.Cell("A3").Value = "KODE";
                    ws.Cell("B3").Value = "ISBN";
                    ws.Cell("C3").Value = "JUDUL";
                    ws.Cell("D3").Value = "PICK";
                    ws.Cell("E3").Value = "SGD";
                    ws.Cell("F3").Value = "BULK";
                    ws.Cell("G3").Value = "SGB";
                    ws.Cell("H3").Value = "TOKO";



                    ws.Cell("A4").InsertData(dt);


                    var c1 = ws.Column("A");
                    c1.Width = 15;

                    var c2 = ws.Column("B");
                    c2.Width = 10;

                    var c3 = ws.Column("C");
                    c3.Width = 50;

                    var c4 = ws.Column("D");
                    c4.Width = 8;

                    var c5 = ws.Column("E");
                    c5.Width = 8;

                    var c6 = ws.Column("F");
                    c6.Width = 15;

                    var c7 = ws.Column("G");
                    c7.Width = 18;
                    c7.ColumnNumber();
                    var c8 = ws.Column("H");
                    c8.Width = 12;


                    wbook.SaveAs(folderPath + "Cek Stock.xlsx");

                    MessageBox.Show("Export berhasil, silakan cek D:Cek_Stock.xlsx");
                }
                else {
                    var ws = wbook.AddWorksheet("master");

                    ws.Cell("A1").Value = "Cek master ISBN";
                    ws.Cell("A3").Value = "ISBN";
                    ws.Cell("B3").Value = "KODE";
                    ws.Cell("C3").Value = "JUDUL";
                    ws.Cell("D3").Value = "KAT";
                    ws.Cell("E3").Value = "STATUS";
                    ws.Cell("F3").Value = "TGL MASUK";
                    ws.Cell("G3").Value = "UPDATE TERAKHIR";
                    ws.Cell("H3").Value = "DOBEL";



                    ws.Cell("A4").InsertData(dt);


                    var c1 = ws.Column("A");
                    c1.Width = 15;

                    var c2 = ws.Column("B");
                    c2.Width = 10;

                    var c3 = ws.Column("C");
                    c3.Width = 50;

                    var c4 = ws.Column("D");
                    c4.Width = 8;

                    var c5 = ws.Column("E");
                    c5.Width = 8;

                    var c6 = ws.Column("F");
                    c6.Width = 15;

                    var c7 = ws.Column("G");
                    c7.Width = 18;
                    c7.ColumnNumber();
                    var c8 = ws.Column("H");
                    c8.Width = 12;


                    wbook.SaveAs(folderPath + "Cek Master.xlsx");

                    MessageBox.Show("Export berhasil, silakan cek D:Cek_Master.xlsx");

                }


            }
        }


        

        private static void ExportDataSet(DataSet ds, string destination)
            {
                using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {

                    var workbookPart = workbook.AddWorkbookPart();

                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                    workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();


                    foreach (System.Data.DataTable table in ds.Tables)
                    {

                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);



                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        uint sheetId = 1;
                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                        {
                            sheetId =
                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }


                       
                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                        sheets.Append(sheet);

                        DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

         

                    List<String> columns = new List<string>();
                        foreach (System.Data.DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                            headerRow.AppendChild(cell);
                        }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            foreach (String col in columns)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                                newRow.AppendChild(cell);
                            }


                        }

                    
                    }
                }

            }


    }

}
