using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Threading;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;

namespace TXTLOG_TO_EXCEL_Project
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region 1.  Checked DKDMS.exe Execution
            //取得此process的名稱
            String name = Process.GetCurrentProcess().ProcessName;
            //取得所有與目前process名稱相同的process
            Process[] ps = Process.GetProcessesByName(name);
            //ps.Length > 1 表示此proces已重複執行
            if (ps.Length > 1)
            {
                System.Environment.Exit(System.Environment.ExitCode);
            }
            #endregion

            this.Text = "TXT LOG_TO_EXCEL Ver :" + FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString();
        }

        private void lbl_MDB_Click(object sender, EventArgs e)
        {
            dlg_MDB.Title = "開啟MDB文件";
            dlg_MDB.Filter = "MDB files (*.MDB)|*.MDB";
            dlg_MDB.FilterIndex = 1;
            dlg_MDB.RestoreDirectory = true;
            dlg_MDB.Multiselect = false;
            if (dlg_MDB.ShowDialog() == DialogResult.OK)
            {
                DataTable dt = new DataTable();
                txt_MDBPATH.Text = dlg_MDB.FileName;
                string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + dlg_MDB.FileName;
                try
                {
                    using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                    {
                        conn.Open();
                        OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from SPCLogData order by VarID", sConnectionString);
                        adapter.Fill(dt);
                        conn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("讀取MDB檔失敗:" + ex.Message);
                }
            }
        }

        private void btn_Report_Click(object sender, EventArgs e)
        {
            if(txt_TXTPATH.Text == "")
            {
                MessageBox.Show("請先選擇TXT LOG檔!");
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            string strReadline = ""; //取得內容
            StreamReader reader = new StreamReader(txt_TXTPATH.Text, System.Text.Encoding.Default); //作業系統目前 ANSI 字碼頁的編碼方式               
            if ((strReadline = reader.ReadToEnd()) != null)
            {
                string[] ReadlineArray = Regex.Split(strReadline, "================================================================================");
                if (TableToExcel(ReadlineArray, dlg_TXT.FileName.Replace("txt", "xlsx")))
                {
                    MessageBox.Show("匯出成功!");
                }
            }
            reader.Close();
            this.Cursor = Cursors.Default;
        }

        private void lbl_TXT_Click(object sender, EventArgs e)
        {
            dlg_TXT.Title = "開啟LOG(txt)文件";
            dlg_TXT.Filter = "txt files (*.txt)|*.txt";
            dlg_TXT.FilterIndex = 1;
            dlg_TXT.RestoreDirectory = true;
            dlg_TXT.Multiselect = false;
            if (dlg_TXT.ShowDialog() == DialogResult.OK)
            {               
                txt_TXTPATH.Text = dlg_TXT.FileName;               
            }
        }

        #region Datable匯出成Excel
        /// <summary>
        /// Datable匯出成Excel
        /// </summary>
        /// <param name="dt">內容列</param>
        /// <param name="file">檔名</param>
        private bool TableToExcel(string[] arraystr, string file)
        {
            int iCurrentRow = 0;  //記憶已被使用的列
            IWorkbook workbook;

            try
            {
                workbook = new XSSFWorkbook();
            }
            catch
            {
                workbook = new HSSFWorkbook();
                file = file.Replace("xlsx", "xls");
            }

            ISheet sheet1 = workbook.CreateSheet("ALL");
            
            //編輯字體
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.FontName = "Tahoma";    //字型
            font.FontHeightInPoints = 10;  //字體大小

            //表頭
            try 
            {
                string[] ReadlineCOL = Regex.Split(arraystr[0].ToString(), "\r\n");
                for (int i = 0; i < ReadlineCOL.Length; i++)
                {
                    XSSFCellStyle col_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    //建立新的列與欄位
                    IRow col_row = sheet1.CreateRow(i);
                    ICell col_cell = col_row.CreateCell(0);
                    if (i == 8)
                    {
                        if (ReadlineCOL[i].ToString().Contains("PASS"))
                        {
                            col_style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;                            
                        }
                        else
                        {
                            col_style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;                           
                        }
                        col_style.FillPattern = FillPattern.SolidForeground;  //FillPattern=填滿樣式，但填滿樣式只有前景，即SolidForeground，也就是說要將自己以為的”背景色”，改為指定前景色(ForegroundColor)
                        sheet1.AddMergedRegion(new CellRangeAddress(8, 8, 0, 7)); //MergedRegion=合併區域                       
                    }
                    col_style.SetFont(font);  //設置字體樣式                   
                    col_cell.SetCellValue(ReadlineCOL[i].ToString());
                    col_cell.CellStyle = col_style;

                    //記憶已被使用的列
                    if (ReadlineCOL.Length - 1 == i)
                    {
                        iCurrentRow = i;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("表頭寫入失敗:" + ex.Message);
                return false;
            }

            //內容
            try
            {
                //PASS
                XSSFCellStyle stylePASS = (XSSFCellStyle)workbook.CreateCellStyle();
                stylePASS.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
                stylePASS.FillPattern = FillPattern.SolidForeground;  //FillPattern = 填滿樣式
                stylePASS.SetFont(font);

                //FAIL
                XSSFCellStyle styleFAIL = (XSSFCellStyle)workbook.CreateCellStyle();
                styleFAIL.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                styleFAIL.FillPattern = FillPattern.SolidForeground;  //FillPattern = 填滿樣式
                styleFAIL.SetFont(font);

                //Default
                XSSFCellStyle styleDefault = (XSSFCellStyle)workbook.CreateCellStyle();
                styleDefault.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
                styleDefault.FillPattern = FillPattern.SolidForeground;  //FillPattern = 填滿樣式
                styleDefault.SetFont(font);

                for (int i = 1; i < arraystr.Length; i++)
                {
                    string sDetail = arraystr[i].Trim();
                    string[] arrayDetail = Regex.Split(sDetail, "\r\n");

                    for (int j = 0; j < arrayDetail.Length; j++)
                    {                       
                        iCurrentRow = iCurrentRow + 1;                       
                        IRow row1 = sheet1.CreateRow(iCurrentRow); 
                        ICell cell = row1.CreateCell(0);

                        if (arrayDetail[j].ToString().Contains("PASS"))
                        {
                            cell.CellStyle = stylePASS;
                            sheet1.AddMergedRegion(new CellRangeAddress(iCurrentRow, iCurrentRow, 0, 7)); //MergedRegion=合併區
                        }
                        else if (arrayDetail[j].ToString().Contains("FAIL"))
                        {
                            cell.CellStyle = styleFAIL;
                            sheet1.AddMergedRegion(new CellRangeAddress(iCurrentRow, iCurrentRow, 0, 7)); //MergedRegion=合併區
                        }
                        else
                        {
                            cell.CellStyle = styleDefault;
                        }
                       
                        cell.SetCellValue(arrayDetail[j].ToString());
                    }

                    //每個STEP結尾在空兩列
                    for(int h = 0; h < 2; h++)
                    {
                        iCurrentRow = iCurrentRow + 1;
                        IRow space_row1 = sheet1.CreateRow(iCurrentRow);
                        ICell space_cell = space_row1.CreateCell(0);
                        space_cell.SetCellValue("");
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("內容寫入失敗:" + ex.Message);
                return false;
            }

            try
            {
                MemoryStream stream = new MemoryStream();
                workbook.Write(stream);
                byte[] buf = stream.ToArray();
                stream.Flush();

                //儲存為Excel檔案  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show("寫入EXCEL檔失敗:" + ex.Message);
                return false;
            }

            return true;
        }
        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
        }

    }
}
