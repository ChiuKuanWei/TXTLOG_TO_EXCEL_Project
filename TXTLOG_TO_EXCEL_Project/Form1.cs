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
using System.Collections;
using TXTLOG_TO_EXCEL_Project.Module;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;

namespace TXTLOG_TO_EXCEL_Project
{
    public partial class Form1 : Form
    {
        ArrayList AL_CellIndex = new ArrayList();  //記憶Sheet(ALL)結果為FAIL的項目儲存格索引 

        DataTable _dt;  //取得每個測試項的描述

        string sNormalFile = "";  //取得正常路徑

        /// <summary>
        /// 取得Turn On Timing Test_Multi or Single各項數據
        /// </summary>
        ArrayList arrayList_TurnON = new ArrayList(); 

        /// <summary>
        /// 取得Hold Up Timing Test_Multi or Single各項數據
        /// </summary>
        ArrayList arrayList_HoldUp = new ArrayList();

        /// <summary>
        /// 取得PS OFF Time 80611各項數據
        /// </summary>
        ArrayList arrayList_PSOFF = new ArrayList();

        /// <summary>
        /// 取得PS ON Delay time 80611各項數據
        /// </summary>
        ArrayList arrayList_PSON = new ArrayList();

        /// <summary>
        /// 取得Input Output Eff Noise Multi or Single各項數據
        /// </summary>
        ArrayList arrayList_EffNoise = new ArrayList();

        /// <summary>
        /// 取得AC_DC Line Sag Surge 80611N-2_Multi各項數據
        /// </summary>
        ArrayList arrayList_DropOut = new ArrayList();

        /// <summary>
        /// 取得AC Line Sag Surge 80611N-2_Multi各項數據
        /// </summary>
        ArrayList arrayList_SagSurge = new ArrayList();

        /// <summary>
        /// 取得Input Output Accuracy Multi or Single各項數據
        /// </summary>
        ArrayList arrayList_Accuracy = new ArrayList();

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
                txt_MDBPATH.Text = dlg_MDB.FileName;
                sNormalFile = txt_MDBPATH.Text.Substring(0, txt_MDBPATH.Text.LastIndexOf("\\") + 1) + "TESTINFO.MDB";  //取得正常檔案
                Select_MDB_Data(sNormalFile);
            }
        }

        /// <summary>
        /// 抓取MDB檔內容
        /// </summary>
        /// <param name="sNormalFile">檔案路徑</param>
        private void Select_MDB_Data(string sNormalFile)
        {
            _dt = new DataTable();
            string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + sNormalFile;
            try
            {
                using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                {
                    conn.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from TPInfo where Key like 'SeqExt%'", sConnectionString);
                    adapter.Fill(_dt);
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取MDB檔失敗:" + ex.Message);
            }
        }

        private void btn_Report_Click(object sender, EventArgs e)
        {
            if (txt_MDBPATH.Text == "")
            {
                MessageBox.Show("請先選擇MDB檔!");
                return;
            }
            if (txt_TXTPATH.Text == "")
            {
                MessageBox.Show("請先選擇TXT LOG檔!");
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            string strReadline = ""; //取得內容
            StreamReader reader = new StreamReader(txt_TXTPATH.Text, System.Text.Encoding.Default); //作業系統目前 ANSI 字碼頁的編碼方式               
            if ((strReadline = reader.ReadToEnd()) != null)
            {
                if(_dt.Rows.Count == 0)
                {
                    Select_MDB_Data(sNormalFile);
                }
                string[] ReadlineArray = Regex.Split(strReadline, "================================================================================");
                if (DataToExcel(ReadlineArray, dlg_TXT.FileName.Replace("txt", "xlsx")))
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

        #region 匯出成Excel_Sheet(Item ALL)
        /// <summary>
        /// Datable匯出成Excel
        /// </summary>
        /// <param name="dt">內容列</param>
        /// <param name="file">檔名</param>
        private bool DataToExcel(string[] arraystr, string file)
        {
            #region 清空陣列
            if (AL_CellIndex.Count > 0)
            {
                AL_CellIndex.Clear();
            }

            if(arrayList_TurnON.Count > 0)
            {
                arrayList_TurnON.Clear();               
            }
            arrayList_TurnON.Add
                (
                     "SEQ" + "|" + "Name" + "|" + "A(12V)" + "|" + "A(12Vsb)" + "|" + "A(PWOK)" + "|" + "A(Vin_Good)" + "|" + "A(SMBAlert)"
                     + "|" + "B(12V)" + "|" + "B(12Vsb)" + "|" + "B(PWOK)" + "|" + "B(Vin_Good)" + "|" + "B(SMBAlert)" + "|"
                     + "Trig_S(12V)" + "|" + "Trig_S(12Vsb)" + "|" + "Trig_S(PWOK)" + "|" + "Trig_S(Vin_Good)" + "|" + "Trig_S(SMBAlert)" + "|"
                     + "Trig_E(12V)" + "|" + "Trig_E(12Vsb)" + "|" + "Trig_E(PWOK)" + "|" + "Trig_E(Vin_Good)" + "|" + "Trig_E(SMBAlert)" + "|"
                     + "T_Max(12V)" + "|" + "T_Max(12Vsb)" + "|" + "T_Max(PWOK)" + "|" + "T_Max(Vin_Good)" + "|" + "T_Max(SMBAlert)" + "|"
                     + "T_Min(12V)" + "|" + "T_Min(12Vsb)" + "|" + "T_Min(PWOK)" + "|" + "T_Min(Vin_Good)" + "|" + "T_Min(SMBAlert)" + "|"
                     + "Td_Max(12V)" + "|" + "Td_Max(12Vsb)" + "|" + "Td_Max(PWOK)" + "|" + "Td_Max(Vin_Good)" + "|" + "Td_Max(SMBAlert)" + "|"
                     + "Td_Min(12V)" + "|" + "Td_Min(12Vsb)" + "|" + "Td_Min(PWOK)" + "|" + "Td_Min(Vin_Good)" + "|" + "Td_Min(SMBAlert)" + "|"
                     + "Line" + "|" + "Load(12V)" + "|" + "Load(12Vsb)" + "|" + "Load(PWOK)" + "|" + "Load(Vin_Good)" + "|" + "Load(SMBAlert)"
                );

            if (arrayList_HoldUp.Count > 0)
            {
                arrayList_HoldUp.Clear();
            }
            arrayList_HoldUp.Add
                (
                     "SEQ" + "|" + "Name" + "|" + "A(12V)" + "|" + "A(12Vsb)" + "|" + "A(PWOK)" + "|" + "A(Vin_Good)" + "|" + "A(SMBAlert)"
                     + "|" + "B(12V)" + "|" + "B(12Vsb)" + "|" + "B(PWOK)" + "|" + "B(Vin_Good)" + "|" + "B(SMBAlert)" + "|"
                     + "Trig_S(12V)" + "|" + "Trig_S(12Vsb)" + "|" + "Trig_S(PWOK)" + "|" + "Trig_S(Vin_Good)" + "|" + "Trig_S(SMBAlert)" + "|"
                     + "Trig_E(12V)" + "|" + "Trig_E(12Vsb)" + "|" + "Trig_E(PWOK)" + "|" + "Trig_E(Vin_Good)" + "|" + "Trig_E(SMBAlert)" + "|"
                     + "T_Max(12V)" + "|" + "T_Max(12Vsb)" + "|" + "T_Max(PWOK)" + "|" + "T_Max(Vin_Good)" + "|" + "T_Max(SMBAlert)" + "|"
                     + "T_Min(12V)" + "|" + "T_Min(12Vsb)" + "|" + "T_Min(PWOK)" + "|" + "T_Min(Vin_Good)" + "|" + "T_Min(SMBAlert)" + "|"
                     + "Td_Max(12V)" + "|" + "Td_Max(12Vsb)" + "|" + "Td_Max(PWOK)" + "|" + "Td_Max(Vin_Good)" + "|" + "Td_Max(SMBAlert)" + "|"
                     + "Td_Min(12V)" + "|" + "Td_Min(12Vsb)" + "|" + "Td_Min(PWOK)" + "|" + "Td_Min(Vin_Good)" + "|" + "Td_Min(SMBAlert)" + "|"
                     + "Line" + "|" + "Load(12V)" + "|" + "Load(12Vsb)" + "|" + "Load(PWOK)" + "|" + "Load(Vin_Good)" + "|" + "Load(SMBAlert)"
                );

            if (arrayList_PSOFF.Count > 0)
            {
                arrayList_PSOFF.Clear();
            }
            arrayList_PSOFF.Add
                (
                     "SEQ" + "|" + "Name" + "|" + "A(12V)" + "|" + "A(12Vsb)" + "|" + "A(PWOK)" + "|" + "A(Vin_Good)" + "|" + "A(SMBAlert)"
                     + "|" + "B(12V)" + "|" + "B(12Vsb)" + "|" + "B(PWOK)" + "|" + "B(Vin_Good)" + "|" + "B(SMBAlert)" + "|"
                     + "Trig_S(12V)" + "|" + "Trig_S(12Vsb)" + "|" + "Trig_S(PWOK)" + "|" + "Trig_S(Vin_Good)" + "|" + "Trig_S(SMBAlert)" + "|"
                     + "Trig_E(12V)" + "|" + "Trig_E(12Vsb)" + "|" + "Trig_E(PWOK)" + "|" + "Trig_E(Vin_Good)" + "|" + "Trig_E(SMBAlert)" + "|"
                     + "T_Max(12V)" + "|" + "T_Max(12Vsb)" + "|" + "T_Max(PWOK)" + "|" + "T_Max(Vin_Good)" + "|" + "T_Max(SMBAlert)" + "|"
                     + "T_Min(12V)" + "|" + "T_Min(12Vsb)" + "|" + "T_Min(PWOK)" + "|" + "T_Min(Vin_Good)" + "|" + "T_Min(SMBAlert)" + "|"
                     + "Td_Max(12V)" + "|" + "Td_Max(12Vsb)" + "|" + "Td_Max(PWOK)" + "|" + "Td_Max(Vin_Good)" + "|" + "Td_Max(SMBAlert)" + "|"
                     + "Td_Min(12V)" + "|" + "Td_Min(12Vsb)" + "|" + "Td_Min(PWOK)" + "|" + "Td_Min(Vin_Good)" + "|" + "Td_Min(SMBAlert)"                     
                );

            if (arrayList_PSON.Count > 0)
            {
                arrayList_PSON.Clear();
            }
            arrayList_PSON.Add
                (
                     "SEQ" + "|" + "Name" + "|" + "A(12V)" + "|" + "A(12Vsb)" + "|" + "A(PWOK)" + "|" + "A(Vin_Good)" + "|" + "A(SMBAlert)"
                     + "|" + "B(12V)" + "|" + "B(12Vsb)" + "|" + "B(PWOK)" + "|" + "B(Vin_Good)" + "|" + "B(SMBAlert)" + "|"
                     + "Trig_S(12V)" + "|" + "Trig_S(12Vsb)" + "|" + "Trig_S(PWOK)" + "|" + "Trig_S(Vin_Good)" + "|" + "Trig_S(SMBAlert)" + "|"
                     + "Trig_E(12V)" + "|" + "Trig_E(12Vsb)" + "|" + "Trig_E(PWOK)" + "|" + "Trig_E(Vin_Good)" + "|" + "Trig_E(SMBAlert)" + "|"
                     + "T_Max(12V)" + "|" + "T_Max(12Vsb)" + "|" + "T_Max(PWOK)" + "|" + "T_Max(Vin_Good)" + "|" + "T_Max(SMBAlert)" + "|"
                     + "T_Min(12V)" + "|" + "T_Min(12Vsb)" + "|" + "T_Min(PWOK)" + "|" + "T_Min(Vin_Good)" + "|" + "T_Min(SMBAlert)" + "|"
                     + "Td_Max(12V)" + "|" + "Td_Max(12Vsb)" + "|" + "Td_Max(PWOK)" + "|" + "Td_Max(Vin_Good)" + "|" + "Td_Max(SMBAlert)" + "|"
                     + "Td_Min(12V)" + "|" + "Td_Min(12Vsb)" + "|" + "Td_Min(PWOK)" + "|" + "Td_Min(Vin_Good)" + "|" + "Td_Min(SMBAlert)"
                );

            if (arrayList_EffNoise.Count > 0)
            {
                arrayList_EffNoise.Clear();
            }
            arrayList_EffNoise.Add
                (
                     "SEQ" + "|" + "Name" + "|" + "Ripple(12V)" + "|" + "Ripple(12Vsb)" + "|" + "Ripple(PWOK)" + "|" + "Ripple(Vin_Good)" + "|" + "Ripple(SMBAlert)"
                     + "|" + "Vin" + "|" + "Frequ" + "|"
                     + "Vout_Max(12V)" + "|" + "Vout_Max(12Vsb)" + "|" + "Vout_Max(PWOK)" + "|" + "Vout_Max(Vin_Good)" + "|" + "Vout_Max(SMBAlert)" + "|"
                     + "Vout_Min(12V)" + "|" + "Vout_Min(12Vsb)" + "|" + "Vout_Min(PWOK)" + "|" + "Vout_Min(Vin_Good)" + "|" + "Vout_Min(SMBAlert)" + "|"                      
                     + "Load(12V)" + "|" + "Load(12Vsb)" + "|" + "Load(PWOK)" + "|" + "Load(Vin_Good)" + "|" + "Load(SMBAlert)" 
                );

            #endregion

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
            ISheet sheet_Item = workbook.CreateSheet("Item");
            ISheet sheet_ALL = workbook.CreateSheet("ALL");

            //超連結字體
            XSSFFont hyperlink_font = (XSSFFont)workbook.CreateFont();
            hyperlink_font.FontName = "Calibri";    //字型
            hyperlink_font.FontHeightInPoints = 12;  //字體大小
            hyperlink_font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            hyperlink_font.Underline = NPOI.SS.UserModel.FontUnderlineType.Single;  //底線

            //正常字體
            XSSFFont normal_font = (XSSFFont)workbook.CreateFont();
            normal_font.FontName = "Calibri";    //字型
            normal_font.FontHeightInPoints = 12;  //字體大小

            //正常藍色字體
            XSSFFont normal_Blue_font = (XSSFFont)workbook.CreateFont();
            normal_Blue_font.FontName = "Calibri";    //字型
            normal_Blue_font.FontHeightInPoints = 12;  //字體大小    
            normal_Blue_font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;

            //Sheet(ALL)表頭
            try
            {                
                string[] ReadlineCOL = Regex.Split(arraystr[0].ToString(), "\r\n|\n");
                for (int i = 0; i < ReadlineCOL.Length; i++)
                {
                    XSSFCellStyle col_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    //建立新的列與欄位
                    IRow col_row = sheet_ALL.CreateRow(i);
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
                        sheet_ALL.AddMergedRegion(new CellRangeAddress(8, 8, 0, 7)); //MergedRegion=合併區域                       
                    }
                    col_style.SetFont(normal_font);  //設置字體樣式
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
                MessageBox.Show("Sheet(ALL)表頭寫入失敗:" + ex.Message);
                return false;
            }

            //Sheet(ALL)內容
            try
            {
                //PASS
                XSSFCellStyle stylePASS = (XSSFCellStyle)workbook.CreateCellStyle();
                stylePASS.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
                stylePASS.FillPattern = FillPattern.SolidForeground;  //FillPattern = 填滿樣式
                stylePASS.SetFont(normal_font);

                //FAIL
                XSSFCellStyle styleFAIL = (XSSFCellStyle)workbook.CreateCellStyle();
                styleFAIL.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                styleFAIL.FillPattern = FillPattern.SolidForeground;  //FillPattern = 填滿樣式
                styleFAIL.SetFont(normal_font);

                //Default
                XSSFCellStyle styleDefault = (XSSFCellStyle)workbook.CreateCellStyle();
                styleDefault.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
                styleDefault.FillPattern = FillPattern.SolidForeground;  //FillPattern = 填滿樣式
                styleDefault.SetFont(normal_font);

                for (int i = 1; i < arraystr.Length; i++)
                {
                    string sDetail = arraystr[i].Trim();
                    string[] arrayDetail = Regex.Split(sDetail, "\r\n|\n");


                    #region 取得每項Turn ON測項數據區
                    if (arrayDetail[0].ToString().Contains("Turn On Timing Test_Multi or Single"))
                    {
                        //"STEP.9(UUT Test seq.9) : Turn On Timing Test_Multi or Single(264V_63Hz_H) - PASS"

                        //SEQ
                        string input = arrayDetail[0].ToString();
                        string pattern = @"seq\.(\d+)";

                        Match match = Regex.Match(input, pattern);
                        if (match.Success)
                        {
                            TurnOn_Column.SEQ = match.Groups[1].Value;  //取得match.Groups[0].Value，則返回"seq.9"
                        }

                        //Name
                        string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + sNormalFile;
                        DataTable _dtName = new DataTable();
                        try
                        {
                            using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                            {
                                conn.Open();
                                OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from TPInfo where Key = 'SeqExt" + TurnOn_Column.SEQ + "'", sConnectionString);
                                adapter.Fill(_dtName);
                                conn.Close();
                            }
                            TurnOn_Column.Name = _dtName.Rows[0]["Value"].ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("抓取turn on測項描述內容失敗:" + ex.Message);
                            return false;
                        }

                        //A(12V) B(12V) Trig_S(12V) Trig_E(12V)
                        input = arrayDetail[21].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        string replacement = " ";
                        string result = Regex.Replace(input, pattern, replacement);
                        string[] arrayresult = result.Split(' ');
                        TurnOn_Column.A12V = arrayresult[4].ToString();
                        TurnOn_Column.B12V = arrayresult[5].ToString();
                        TurnOn_Column.Trig_S12V = arrayresult[2].ToString();
                        TurnOn_Column.Trig_E12V = arrayresult[3].ToString();

                        //A(12Vsb) B(12Vsb) Trig_S(12Vsb) Trig_E(12Vsb)
                        input = arrayDetail[22].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.A12Vsb = arrayresult[4].ToString();
                        TurnOn_Column.B12Vsb = arrayresult[5].ToString();
                        TurnOn_Column.Trig_S12Vsb = arrayresult[2].ToString();
                        TurnOn_Column.Trig_E12Vsb = arrayresult[3].ToString();

                        //A(PWOK) B(PWOK) Trig_S(PWOK) Trig_E(PWOK)
                        input = arrayDetail[23].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.APWOK = arrayresult[4].ToString();
                        TurnOn_Column.BPWOK = arrayresult[5].ToString();
                        TurnOn_Column.Trig_SPWOK = arrayresult[2].ToString();
                        TurnOn_Column.Trig_EPWOK = arrayresult[3].ToString();

                        //A(Vin_Good) B(Vin_Good) Trig_S(Vin_Good) Trig_E(Vin_Good)
                        input = arrayDetail[24].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.AVin_Good = arrayresult[4].ToString();
                        TurnOn_Column.BVin_Good = arrayresult[5].ToString();
                        TurnOn_Column.Trig_SVin_Good = arrayresult[2].ToString();
                        TurnOn_Column.Trig_EVin_Good = arrayresult[3].ToString();

                        //A(SMBAlert) B(SMBAlert) Trig_S(SMBAlert) Trig_E(SMBAlert)
                        input = arrayDetail[25].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.ASMBAlert = arrayresult[4].ToString();
                        TurnOn_Column.BSMBAlert = arrayresult[5].ToString();
                        TurnOn_Column.Trig_SSMBAlert = arrayresult[2].ToString();
                        TurnOn_Column.Trig_ESMBAlert = arrayresult[3].ToString();

                        //T_Max(12V) T_Min(12V)
                        input = arrayDetail[30].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.T_Max12V = arrayresult[1].ToString();
                        TurnOn_Column.T_Min12V = arrayresult[2].ToString();

                        //T_Max(12Vsb) T_Min(12Vsb)
                        input = arrayDetail[31].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.T_Max12Vsb = arrayresult[1].ToString();
                        TurnOn_Column.T_Min12Vsb = arrayresult[2].ToString();

                        //T_Max(PWOK) T_Min(PWOK)
                        input = arrayDetail[32].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.T_MaxPWOK = arrayresult[1].ToString();
                        TurnOn_Column.T_MinPWOK = arrayresult[2].ToString();

                        //T_Max(Vin_Good) T_Min(Vin_Good)
                        input = arrayDetail[33].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.T_MaxVin_Good = arrayresult[1].ToString();
                        TurnOn_Column.T_MinVin_Good = arrayresult[2].ToString();

                        //T_Max(SMBAlert) T_Min(SMBAlert)
                        input = arrayDetail[34].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.T_MaxSMBAlert = arrayresult[1].ToString();
                        TurnOn_Column.T_MinSMBAlert = arrayresult[2].ToString();

                        //Td_Max(12V) Td_Min(12V)
                        input = arrayDetail[39].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Td_Max12V = arrayresult[2].ToString();
                        TurnOn_Column.Td_Min12V = arrayresult[3].ToString();

                        //Td_Max(12Vsb) Td_Min(12Vsb)
                        input = arrayDetail[40].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Td_Max12Vsb = arrayresult[2].ToString();
                        TurnOn_Column.Td_Min12Vsb = arrayresult[3].ToString();

                        //Td_Max(PWOK) Td_Min(PWOK)
                        input = arrayDetail[41].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Td_MaxPWOK = arrayresult[2].ToString();
                        TurnOn_Column.Td_MinPWOK = arrayresult[3].ToString();

                        //Td_Max(Vin_Good) Td_Min(Vin_Good)
                        input = arrayDetail[42].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Td_MaxVin_Good = arrayresult[2].ToString();
                        TurnOn_Column.Td_MinVin_Good = arrayresult[3].ToString();

                        //Td_Max(SMBAlert) Td_Min(SMBAlert)
                        input = arrayDetail[43].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Td_MaxSMBAlert = arrayresult[2].ToString();
                        TurnOn_Column.Td_MinSMBAlert = arrayresult[3].ToString();

                        //Line
                        input = arrayDetail[2].ToString().Trim();
                        input = input.Replace(" ", "");
                        arrayresult = input.Split('=');
                        TurnOn_Column.Line = arrayresult[1].ToString();

                        //Load(12V)
                        input = arrayDetail[21].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Load12V = arrayresult[1].ToString();

                        //Load(12Vsb)
                        input = arrayDetail[22].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.Load12Vsb = arrayresult[1].ToString();

                        //Load(PWOK)
                        input = arrayDetail[23].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.LoadPWOK = arrayresult[1].ToString();

                        //Load(Vin_Good)
                        input = arrayDetail[24].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.LoadVin_Good = arrayresult[1].ToString();

                        //Load(Vin_Good)
                        input = arrayDetail[25].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        TurnOn_Column.LoadSMBAlert = arrayresult[1].ToString();
                        
                        arrayList_TurnON.Add
                        (
                              TurnOn_Column.SEQ + "|" + TurnOn_Column.Name + "|" 
                            + TurnOn_Column.A12V + "|" + TurnOn_Column.A12Vsb + "|" + TurnOn_Column.APWOK + "|" + TurnOn_Column.AVin_Good + "|" + TurnOn_Column.ASMBAlert + "|" 
                            + TurnOn_Column.B12V + "|" + TurnOn_Column.B12Vsb + "|" + TurnOn_Column.BPWOK + "|" + TurnOn_Column.BVin_Good + "|" + TurnOn_Column.BSMBAlert + "|"
                            + TurnOn_Column.Trig_S12V + "|" + TurnOn_Column.Trig_S12Vsb + "|" + TurnOn_Column.Trig_SPWOK + "|" + TurnOn_Column.Trig_SVin_Good + "|" + TurnOn_Column.Trig_SSMBAlert + "|" 
                            + TurnOn_Column.Trig_E12V + "|" + TurnOn_Column.Trig_E12Vsb + "|" + TurnOn_Column.Trig_EPWOK + "|" + TurnOn_Column.Trig_EVin_Good + "|" + TurnOn_Column.Trig_ESMBAlert + "|" 
                            + TurnOn_Column.T_Max12V + "|" + TurnOn_Column.T_Max12Vsb + "|" + TurnOn_Column.T_MaxPWOK + "|" + TurnOn_Column.T_MaxVin_Good + "|" + TurnOn_Column.T_MaxSMBAlert + "|"
                            + TurnOn_Column.T_Min12V + "|" + TurnOn_Column.T_Min12Vsb + "|" + TurnOn_Column.T_MinPWOK + "|" + TurnOn_Column.T_MinVin_Good + "|" + TurnOn_Column.T_MinSMBAlert + "|" 
                            + TurnOn_Column.Td_Max12V + "|" + TurnOn_Column.Td_Max12Vsb + "|" + TurnOn_Column.Td_MaxPWOK + "|" + TurnOn_Column.Td_MaxVin_Good + "|" + TurnOn_Column.Td_MaxSMBAlert + "|"
                            + TurnOn_Column.Td_Min12V + "|" + TurnOn_Column.Td_Min12Vsb + "|" + TurnOn_Column.Td_MinPWOK + "|" + TurnOn_Column.Td_MinVin_Good + "|" + TurnOn_Column.Td_MinSMBAlert + "|"
                            + TurnOn_Column.Line + "|" + TurnOn_Column.Load12V + "|" + TurnOn_Column.Load12Vsb + "|" + TurnOn_Column.LoadPWOK + "|" + TurnOn_Column.LoadVin_Good + "|" + TurnOn_Column.LoadSMBAlert
                        );

                    }
                    #endregion

                    #region 取得每項Hold Up測項數據區
                    if (arrayDetail[0].ToString().Contains("Hold Up Timing Test_Multi or Single"))
                    {
                        //"STEP.14(UUT Test seq.14) : Hold Up Timing Test_Multi or Single(264V_63Hz_10 PASS"

                        //SEQ
                        string input = arrayDetail[0].ToString();
                        string pattern = @"seq\.(\d+)";

                        Match match = Regex.Match(input, pattern);
                        if (match.Success)
                        {
                            HoldUp_Column.SEQ = match.Groups[1].Value;  //取得match.Groups[0].Value，則返回"seq.14"
                        }

                        //Name
                        string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + sNormalFile;
                        DataTable _dtName = new DataTable();
                        try
                        {
                            using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                            {
                                conn.Open();
                                OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from TPInfo where Key = 'SeqExt" + HoldUp_Column.SEQ + "'", sConnectionString);
                                adapter.Fill(_dtName);
                                conn.Close();
                            }
                            HoldUp_Column.Name = _dtName.Rows[0]["Value"].ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("抓取hold up測項描述內容失敗:" + ex.Message);
                            return false;
                        }

                        //A(12V) B(12V) Trig_S(12V) Trig_E(12V)
                        input = arrayDetail[22].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        string replacement = " ";
                        string result = Regex.Replace(input, pattern, replacement);
                        string[] arrayresult = result.Split(' ');
                        HoldUp_Column.A12V = arrayresult[4].ToString();
                        HoldUp_Column.B12V = arrayresult[5].ToString();
                        HoldUp_Column.Trig_S12V = arrayresult[2].ToString();
                        HoldUp_Column.Trig_E12V = arrayresult[3].ToString();

                        //A(12Vsb) B(12Vsb) Trig_S(12Vsb) Trig_E(12Vsb)
                        input = arrayDetail[23].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.A12Vsb = arrayresult[4].ToString();
                        HoldUp_Column.B12Vsb = arrayresult[5].ToString();
                        HoldUp_Column.Trig_S12Vsb = arrayresult[2].ToString();
                        HoldUp_Column.Trig_E12Vsb = arrayresult[3].ToString();

                        //A(PWOK) B(PWOK) Trig_S(PWOK) Trig_E(PWOK)
                        input = arrayDetail[24].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.APWOK = arrayresult[4].ToString();
                        HoldUp_Column.BPWOK = arrayresult[5].ToString();
                        HoldUp_Column.Trig_SPWOK = arrayresult[2].ToString();
                        HoldUp_Column.Trig_EPWOK = arrayresult[3].ToString();

                        //A(Vin_Good) B(Vin_Good) Trig_S(Vin_Good) Trig_E(Vin_Good)
                        input = arrayDetail[25].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.AVin_Good = arrayresult[4].ToString();
                        HoldUp_Column.BVin_Good = arrayresult[5].ToString();
                        HoldUp_Column.Trig_SVin_Good = arrayresult[2].ToString();
                        HoldUp_Column.Trig_EVin_Good = arrayresult[3].ToString();

                        //A(SMBAlert) B(SMBAlert) Trig_S(SMBAlert) Trig_E(SMBAlert)
                        input = arrayDetail[26].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.ASMBAlert = arrayresult[4].ToString();
                        HoldUp_Column.BSMBAlert = arrayresult[5].ToString();
                        HoldUp_Column.Trig_SSMBAlert = arrayresult[2].ToString();
                        HoldUp_Column.Trig_ESMBAlert = arrayresult[3].ToString();

                        //T_Max(12V) T_Min(12V)
                        input = arrayDetail[31].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.T_Max12V = arrayresult[1].ToString();
                        HoldUp_Column.T_Min12V = arrayresult[2].ToString();

                        //T_Max(12Vsb) T_Min(12Vsb)
                        input = arrayDetail[32].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.T_Max12Vsb = arrayresult[1].ToString();
                        HoldUp_Column.T_Min12Vsb = arrayresult[2].ToString();

                        //T_Max(PWOK) T_Min(PWOK)
                        input = arrayDetail[33].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.T_MaxPWOK = arrayresult[1].ToString();
                        HoldUp_Column.T_MinPWOK = arrayresult[2].ToString();

                        //T_Max(Vin_Good) T_Min(Vin_Good)
                        input = arrayDetail[34].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.T_MaxVin_Good = arrayresult[1].ToString();
                        HoldUp_Column.T_MinVin_Good = arrayresult[2].ToString();

                        //T_Max(SMBAlert) T_Min(SMBAlert)
                        input = arrayDetail[35].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.T_MaxSMBAlert = arrayresult[1].ToString();
                        HoldUp_Column.T_MinSMBAlert = arrayresult[2].ToString();

                        //Td_Max(12V) Td_Min(12V)
                        input = arrayDetail[42].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Td_Max12V = arrayresult[2].ToString();
                        HoldUp_Column.Td_Min12V = arrayresult[3].ToString();

                        //Td_Max(12Vsb) Td_Min(12Vsb)
                        input = arrayDetail[43].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Td_Max12Vsb = arrayresult[2].ToString();
                        HoldUp_Column.Td_Min12Vsb = arrayresult[3].ToString();

                        //Td_Max(PWOK) Td_Min(PWOK)
                        input = arrayDetail[44].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Td_MaxPWOK = arrayresult[2].ToString();
                        HoldUp_Column.Td_MinPWOK = arrayresult[3].ToString();

                        //Td_Max(Vin_Good) Td_Min(Vin_Good)
                        input = arrayDetail[45].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Td_MaxVin_Good = arrayresult[2].ToString();
                        HoldUp_Column.Td_MinVin_Good = arrayresult[3].ToString();

                        //Td_Max(SMBAlert) Td_Min(SMBAlert)
                        input = arrayDetail[46].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Td_MaxSMBAlert = arrayresult[2].ToString();
                        HoldUp_Column.Td_MinSMBAlert = arrayresult[3].ToString();

                        //Line
                        input = arrayDetail[2].ToString().Trim();
                        input = input.Replace(" ", "");
                        arrayresult = input.Split('=');
                        HoldUp_Column.Line = arrayresult[1].ToString();

                        //Load(12V)
                        input = arrayDetail[22].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Load12V = arrayresult[1].ToString();

                        //Load(12Vsb)
                        input = arrayDetail[23].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.Load12Vsb = arrayresult[1].ToString();

                        //Load(PWOK)
                        input = arrayDetail[24].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.LoadPWOK = arrayresult[1].ToString();

                        //Load(Vin_Good)
                        input = arrayDetail[25].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.LoadVin_Good = arrayresult[1].ToString();

                        //Load(SMBAlert)
                        input = arrayDetail[26].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        HoldUp_Column.LoadSMBAlert = arrayresult[1].ToString();

                        arrayList_HoldUp.Add
                        (
                              HoldUp_Column.SEQ + "|" + HoldUp_Column.Name + "|"
                            + HoldUp_Column.A12V + "|" + HoldUp_Column.A12Vsb + "|" + HoldUp_Column.APWOK + "|" + HoldUp_Column.AVin_Good + "|" + HoldUp_Column.ASMBAlert + "|"
                            + HoldUp_Column.B12V + "|" + HoldUp_Column.B12Vsb + "|" + HoldUp_Column.BPWOK + "|" + HoldUp_Column.BVin_Good + "|" + HoldUp_Column.BSMBAlert + "|"
                            + HoldUp_Column.Trig_S12V + "|" + HoldUp_Column.Trig_S12Vsb + "|" + HoldUp_Column.Trig_SPWOK + "|" + HoldUp_Column.Trig_SVin_Good + "|" + HoldUp_Column.Trig_SSMBAlert + "|"
                            + HoldUp_Column.Trig_E12V + "|" + HoldUp_Column.Trig_E12Vsb + "|" + HoldUp_Column.Trig_EPWOK + "|" + HoldUp_Column.Trig_EVin_Good + "|" + HoldUp_Column.Trig_ESMBAlert + "|"
                            + HoldUp_Column.T_Max12V + "|" + HoldUp_Column.T_Max12Vsb + "|" + HoldUp_Column.T_MaxPWOK + "|" + HoldUp_Column.T_MaxVin_Good + "|" + HoldUp_Column.T_MaxSMBAlert + "|"
                            + HoldUp_Column.T_Min12V + "|" + HoldUp_Column.T_Min12Vsb + "|" + HoldUp_Column.T_MinPWOK + "|" + HoldUp_Column.T_MinVin_Good + "|" + HoldUp_Column.T_MinSMBAlert + "|"
                            + HoldUp_Column.Td_Max12V + "|" + HoldUp_Column.Td_Max12Vsb + "|" + HoldUp_Column.Td_MaxPWOK + "|" + HoldUp_Column.Td_MaxVin_Good + "|" + HoldUp_Column.Td_MaxSMBAlert + "|"
                            + HoldUp_Column.Td_Min12V + "|" + HoldUp_Column.Td_Min12Vsb + "|" + HoldUp_Column.Td_MinPWOK + "|" + HoldUp_Column.Td_MinVin_Good + "|" + HoldUp_Column.Td_MinSMBAlert + "|"
                            + HoldUp_Column.Line + "|" + HoldUp_Column.Load12V + "|" + HoldUp_Column.Load12Vsb + "|" + HoldUp_Column.LoadPWOK + "|" + HoldUp_Column.LoadVin_Good + "|" + HoldUp_Column.LoadSMBAlert
                        );

                    }
                    #endregion

                    #region 取得每項PS OFF測項數據區
                    if (arrayDetail[0].ToString().Contains("PS OFF Time 80611"))
                    {
                        //"STEP.20(UUT Test seq.20) : PS OFF Time 80611_1(90V_47Hz_LL) ---- (1'560) -- FAIL"

                        //SEQ
                        string input = arrayDetail[0].ToString();
                        string pattern = @"seq\.(\d+)";

                        Match match = Regex.Match(input, pattern);
                        if (match.Success)
                        {
                            PSOFF_Column.SEQ = match.Groups[1].Value;  //取得match.Groups[0].Value，則返回"seq.20"
                        }

                        //Name
                        string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + sNormalFile;
                        DataTable _dtName = new DataTable();
                        try
                        {
                            using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                            {
                                conn.Open();
                                OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from TPInfo where Key = 'SeqExt" + PSOFF_Column.SEQ + "'", sConnectionString);
                                adapter.Fill(_dtName);
                                conn.Close();
                            }
                            PSOFF_Column.Name = _dtName.Rows[0]["Value"].ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("抓取ps off測項描述內容失敗:" + ex.Message);
                            return false;
                        }

                        //A(12V) B(12V) Trig_S(12V) Trig_E(12V)
                        input = arrayDetail[22].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        string replacement = " ";
                        string result = Regex.Replace(input, pattern, replacement);
                        string[] arrayresult = result.Split(' ');
                        PSOFF_Column.A12V = arrayresult[3].ToString();
                        PSOFF_Column.B12V = arrayresult[4].ToString();
                        PSOFF_Column.Trig_S12V = arrayresult[1].ToString();
                        PSOFF_Column.Trig_E12V = arrayresult[2].ToString();

                        //A(12Vsb) B(12Vsb) Trig_S(12Vsb) Trig_E(12Vsb)
                        input = arrayDetail[23].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.A12Vsb = arrayresult[3].ToString();
                        PSOFF_Column.B12Vsb = arrayresult[4].ToString();
                        PSOFF_Column.Trig_S12Vsb = arrayresult[1].ToString();
                        PSOFF_Column.Trig_E12Vsb = arrayresult[2].ToString();

                        //A(PWOK) B(PWOK) Trig_S(PWOK) Trig_E(PWOK)
                        input = arrayDetail[24].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.APWOK = arrayresult[3].ToString();
                        PSOFF_Column.BPWOK = arrayresult[4].ToString();
                        PSOFF_Column.Trig_SPWOK = arrayresult[1].ToString();
                        PSOFF_Column.Trig_EPWOK = arrayresult[2].ToString();

                        //A(Vin_Good) B(Vin_Good) Trig_S(Vin_Good) Trig_E(Vin_Good)
                        input = arrayDetail[25].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.AVin_Good = arrayresult[3].ToString();
                        PSOFF_Column.BVin_Good = arrayresult[4].ToString();
                        PSOFF_Column.Trig_SVin_Good = arrayresult[1].ToString();
                        PSOFF_Column.Trig_EVin_Good = arrayresult[2].ToString();

                        //A(SMBAlert) B(SMBAlert) Trig_S(SMBAlert) Trig_E(SMBAlert)
                        input = arrayDetail[26].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.ASMBAlert = arrayresult[3].ToString();
                        PSOFF_Column.BSMBAlert = arrayresult[4].ToString();
                        PSOFF_Column.Trig_SSMBAlert = arrayresult[1].ToString();
                        PSOFF_Column.Trig_ESMBAlert = arrayresult[2].ToString();

                        //T_Max(12V) T_Min(12V)
                        input = arrayDetail[31].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.T_Max12V = arrayresult[1].ToString();
                        PSOFF_Column.T_Min12V = arrayresult[2].ToString();

                        //T_Max(12Vsb) T_Min(12Vsb)
                        input = arrayDetail[32].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.T_Max12Vsb = arrayresult[1].ToString();
                        PSOFF_Column.T_Min12Vsb = arrayresult[2].ToString();

                        //T_Max(PWOK) T_Min(PWOK)
                        input = arrayDetail[33].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.T_MaxPWOK = arrayresult[1].ToString();
                        PSOFF_Column.T_MinPWOK = arrayresult[2].ToString();

                        //T_Max(Vin_Good) T_Min(Vin_Good)
                        input = arrayDetail[34].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.T_MaxVin_Good = arrayresult[1].ToString();
                        PSOFF_Column.T_MinVin_Good = arrayresult[2].ToString();

                        //T_Max(SMBAlert) T_Min(SMBAlert)
                        input = arrayDetail[35].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.T_MaxSMBAlert = arrayresult[1].ToString();
                        PSOFF_Column.T_MinSMBAlert = arrayresult[2].ToString();

                        //Td_Max(12V) Td_Min(12V)
                        input = arrayDetail[40].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.Td_Max12V = arrayresult[2].ToString();
                        PSOFF_Column.Td_Min12V = arrayresult[3].ToString();

                        //Td_Max(12Vsb) Td_Min(12Vsb)
                        input = arrayDetail[41].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.Td_Max12Vsb = arrayresult[2].ToString();
                        PSOFF_Column.Td_Min12Vsb = arrayresult[3].ToString();

                        //Td_Max(PWOK) Td_Min(PWOK)
                        input = arrayDetail[42].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.Td_MaxPWOK = arrayresult[2].ToString();
                        PSOFF_Column.Td_MinPWOK = arrayresult[3].ToString();

                        //Td_Max(Vin_Good) Td_Min(Vin_Good)
                        input = arrayDetail[43].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.Td_MaxVin_Good = arrayresult[2].ToString();
                        PSOFF_Column.Td_MinVin_Good = arrayresult[3].ToString();

                        //Td_Max(SMBAlert) Td_Min(SMBAlert)
                        input = arrayDetail[44].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSOFF_Column.Td_MaxSMBAlert = arrayresult[2].ToString();
                        PSOFF_Column.Td_MinSMBAlert = arrayresult[3].ToString();                      

                        arrayList_PSOFF.Add
                        (
                              PSOFF_Column.SEQ + "|" + PSOFF_Column.Name + "|"
                            + PSOFF_Column.A12V + "|" + PSOFF_Column.A12Vsb + "|" + PSOFF_Column.APWOK + "|" + PSOFF_Column.AVin_Good + "|" + PSOFF_Column.ASMBAlert + "|"
                            + PSOFF_Column.B12V + "|" + PSOFF_Column.B12Vsb + "|" + PSOFF_Column.BPWOK + "|" + PSOFF_Column.BVin_Good + "|" + PSOFF_Column.BSMBAlert + "|"
                            + PSOFF_Column.Trig_S12V + "|" + PSOFF_Column.Trig_S12Vsb + "|" + PSOFF_Column.Trig_SPWOK + "|" + PSOFF_Column.Trig_SVin_Good + "|" + PSOFF_Column.Trig_SSMBAlert + "|"
                            + PSOFF_Column.Trig_E12V + "|" + PSOFF_Column.Trig_E12Vsb + "|" + PSOFF_Column.Trig_EPWOK + "|" + PSOFF_Column.Trig_EVin_Good + "|" + PSOFF_Column.Trig_ESMBAlert + "|"
                            + PSOFF_Column.T_Max12V + "|" + PSOFF_Column.T_Max12Vsb + "|" + PSOFF_Column.T_MaxPWOK + "|" + PSOFF_Column.T_MaxVin_Good + "|" + PSOFF_Column.T_MaxSMBAlert + "|"
                            + PSOFF_Column.T_Min12V + "|" + PSOFF_Column.T_Min12Vsb + "|" + PSOFF_Column.T_MinPWOK + "|" + PSOFF_Column.T_MinVin_Good + "|" + PSOFF_Column.T_MinSMBAlert + "|"
                            + PSOFF_Column.Td_Max12V + "|" + PSOFF_Column.Td_Max12Vsb + "|" + PSOFF_Column.Td_MaxPWOK + "|" + PSOFF_Column.Td_MaxVin_Good + "|" + PSOFF_Column.Td_MaxSMBAlert + "|"
                            + PSOFF_Column.Td_Min12V + "|" + PSOFF_Column.Td_Min12Vsb + "|" + PSOFF_Column.Td_MinPWOK + "|" + PSOFF_Column.Td_MinVin_Good + "|" + PSOFF_Column.Td_MinSMBAlert                           
                        );

                    }
                    #endregion

                    #region 取得每項PS ON測項數據區
                    if (arrayDetail[0].ToString().Contains("PS ON Delay time 80611"))
                    {
                        //"STEP.19(UUT Test seq.19) : PS ON Delay time 80611(90V_47Hz_L) ---- (0'858)  PASS"

                        //SEQ
                        string input = arrayDetail[0].ToString();
                        string pattern = @"seq\.(\d+)";

                        Match match = Regex.Match(input, pattern);
                        if (match.Success)
                        {
                            PSON_Column.SEQ = match.Groups[1].Value;  //取得match.Groups[0].Value，則返回"seq.19"
                        }

                        //Name
                        string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + sNormalFile;
                        DataTable _dtName = new DataTable();
                        try
                        {
                            using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                            {
                                conn.Open();
                                OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from TPInfo where Key = 'SeqExt" + PSON_Column.SEQ + "'", sConnectionString);
                                adapter.Fill(_dtName);
                                conn.Close();
                            }
                            PSON_Column.Name = _dtName.Rows[0]["Value"].ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("抓取ps on測項描述內容失敗:" + ex.Message);
                            return false;
                        }

                        //A(12V) B(12V) Trig_S(12V) Trig_E(12V)
                        input = arrayDetail[22].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        string replacement = " ";
                        string result = Regex.Replace(input, pattern, replacement);
                        string[] arrayresult = result.Split(' ');
                        PSON_Column.A12V = arrayresult[3].ToString();
                        PSON_Column.B12V = arrayresult[4].ToString();
                        PSON_Column.Trig_S12V = arrayresult[1].ToString();
                        PSON_Column.Trig_E12V = arrayresult[2].ToString();

                        //A(12Vsb) B(12Vsb) Trig_S(12Vsb) Trig_E(12Vsb)
                        input = arrayDetail[23].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.A12Vsb = arrayresult[3].ToString();
                        PSON_Column.B12Vsb = arrayresult[4].ToString();
                        PSON_Column.Trig_S12Vsb = arrayresult[1].ToString();
                        PSON_Column.Trig_E12Vsb = arrayresult[2].ToString();

                        //A(PWOK) B(PWOK) Trig_S(PWOK) Trig_E(PWOK)
                        input = arrayDetail[24].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.APWOK = arrayresult[3].ToString();
                        PSON_Column.BPWOK = arrayresult[4].ToString();
                        PSON_Column.Trig_SPWOK = arrayresult[1].ToString();
                        PSON_Column.Trig_EPWOK = arrayresult[2].ToString();

                        //A(Vin_Good) B(Vin_Good) Trig_S(Vin_Good) Trig_E(Vin_Good)
                        input = arrayDetail[25].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.AVin_Good = arrayresult[3].ToString();
                        PSON_Column.BVin_Good = arrayresult[4].ToString();
                        PSON_Column.Trig_SVin_Good = arrayresult[1].ToString();
                        PSON_Column.Trig_EVin_Good = arrayresult[2].ToString();

                        //A(SMBAlert) B(SMBAlert) Trig_S(SMBAlert) Trig_E(SMBAlert)
                        input = arrayDetail[26].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.ASMBAlert = arrayresult[3].ToString();
                        PSON_Column.BSMBAlert = arrayresult[4].ToString();
                        PSON_Column.Trig_SSMBAlert = arrayresult[1].ToString();
                        PSON_Column.Trig_ESMBAlert = arrayresult[2].ToString();

                        //T_Max(12V) T_Min(12V)
                        input = arrayDetail[31].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.T_Max12V = arrayresult[1].ToString();
                        PSON_Column.T_Min12V = arrayresult[2].ToString();

                        //T_Max(12Vsb) T_Min(12Vsb)
                        input = arrayDetail[32].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.T_Max12Vsb = arrayresult[1].ToString();
                        PSON_Column.T_Min12Vsb = arrayresult[2].ToString();

                        //T_Max(PWOK) T_Min(PWOK)
                        input = arrayDetail[33].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.T_MaxPWOK = arrayresult[1].ToString();
                        PSON_Column.T_MinPWOK = arrayresult[2].ToString();

                        //T_Max(Vin_Good) T_Min(Vin_Good)
                        input = arrayDetail[34].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.T_MaxVin_Good = arrayresult[1].ToString();
                        PSON_Column.T_MinVin_Good = arrayresult[2].ToString();

                        //T_Max(SMBAlert) T_Min(SMBAlert)
                        input = arrayDetail[35].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.T_MaxSMBAlert = arrayresult[1].ToString();
                        PSON_Column.T_MinSMBAlert = arrayresult[2].ToString();

                        //Td_Max(12V) Td_Min(12V)
                        input = arrayDetail[41].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.Td_Max12V = arrayresult[2].ToString();
                        PSON_Column.Td_Min12V = arrayresult[3].ToString();

                        //Td_Max(12Vsb) Td_Min(12Vsb)
                        input = arrayDetail[42].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.Td_Max12Vsb = arrayresult[2].ToString();
                        PSON_Column.Td_Min12Vsb = arrayresult[3].ToString();

                        //Td_Max(PWOK) Td_Min(PWOK)
                        input = arrayDetail[43].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.Td_MaxPWOK = arrayresult[2].ToString();
                        PSON_Column.Td_MinPWOK = arrayresult[3].ToString();

                        //Td_Max(Vin_Good) Td_Min(Vin_Good)
                        input = arrayDetail[44].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.Td_MaxVin_Good = arrayresult[2].ToString();
                        PSON_Column.Td_MinVin_Good = arrayresult[3].ToString();

                        //Td_Max(SMBAlert) Td_Min(SMBAlert)
                        input = arrayDetail[45].ToString().Trim();
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        PSON_Column.Td_MaxSMBAlert = arrayresult[2].ToString();
                        PSON_Column.Td_MinSMBAlert = arrayresult[3].ToString();

                        arrayList_PSON.Add
                        (
                              PSON_Column.SEQ + "|" + PSON_Column.Name + "|"
                            + PSON_Column.A12V + "|" + PSON_Column.A12Vsb + "|" + PSON_Column.APWOK + "|" + PSON_Column.AVin_Good + "|" + PSON_Column.ASMBAlert + "|"
                            + PSON_Column.B12V + "|" + PSON_Column.B12Vsb + "|" + PSON_Column.BPWOK + "|" + PSON_Column.BVin_Good + "|" + PSON_Column.BSMBAlert + "|"
                            + PSON_Column.Trig_S12V + "|" + PSON_Column.Trig_S12Vsb + "|" + PSON_Column.Trig_SPWOK + "|" + PSON_Column.Trig_SVin_Good + "|" + PSON_Column.Trig_SSMBAlert + "|"
                            + PSON_Column.Trig_E12V + "|" + PSON_Column.Trig_E12Vsb + "|" + PSON_Column.Trig_EPWOK + "|" + PSON_Column.Trig_EVin_Good + "|" + PSON_Column.Trig_ESMBAlert + "|"
                            + PSON_Column.T_Max12V + "|" + PSON_Column.T_Max12Vsb + "|" + PSON_Column.T_MaxPWOK + "|" + PSON_Column.T_MaxVin_Good + "|" + PSON_Column.T_MaxSMBAlert + "|"
                            + PSON_Column.T_Min12V + "|" + PSON_Column.T_Min12Vsb + "|" + PSON_Column.T_MinPWOK + "|" + PSON_Column.T_MinVin_Good + "|" + PSON_Column.T_MinSMBAlert + "|"
                            + PSON_Column.Td_Max12V + "|" + PSON_Column.Td_Max12Vsb + "|" + PSON_Column.Td_MaxPWOK + "|" + PSON_Column.Td_MaxVin_Good + "|" + PSON_Column.Td_MaxSMBAlert + "|"
                            + PSON_Column.Td_Min12V + "|" + PSON_Column.Td_Min12Vsb + "|" + PSON_Column.Td_MinPWOK + "|" + PSON_Column.Td_MinVin_Good + "|" + PSON_Column.Td_MinSMBAlert
                        );

                    }
                    #endregion

                    #region 取得每項Eff_Noise測項數據區
                    if (arrayDetail[0].ToString().Contains("Input Output Eff Noise Multi or Single"))
                    {
                        //"STEP.78(UUT Test seq.78) : Input Output Eff Noise Multi or Single(90V_47Hz_ FAIL"

                        //SEQ
                        string input = arrayDetail[0].ToString();
                        string pattern = @"seq\.(\d+)";

                        Match match = Regex.Match(input, pattern);
                        if (match.Success)
                        {
                            Eff_Noise_Column.SEQ = match.Groups[1].Value;  //取得match.Groups[0].Value，則返回"seq.78"
                        }

                        //Name
                        string sConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + sNormalFile;
                        DataTable _dtName = new DataTable();
                        try
                        {
                            using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                            {
                                conn.Open();
                                OleDbDataAdapter adapter = new OleDbDataAdapter(@"select * from TPInfo where Key = 'SeqExt" + Eff_Noise_Column.SEQ + "'", sConnectionString);
                                adapter.Fill(_dtName);
                                conn.Close();
                            }
                            Eff_Noise_Column.Name = _dtName.Rows[0]["Value"].ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("抓取Eff_Noise測項描述內容失敗:" + ex.Message);
                            return false;
                        }

                        //Ripple(12V) 
                        input = arrayDetail[63].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        string replacement = " ";
                        string result = Regex.Replace(input, pattern, replacement);
                        string[] arrayresult = result.Split(' ');
                        Eff_Noise_Column.Ripple12V = arrayresult[1].ToString();

                        //Ripple(12Vsb) 
                        input = arrayDetail[64].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Ripple12Vsb = arrayresult[1].ToString();

                        //Ripple(PWOK) 
                        input = arrayDetail[65].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.RipplePWOK = arrayresult[1].ToString();

                        //Ripple(Vin_Good) 
                        input = arrayDetail[66].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.RippleVin_Good = arrayresult[1].ToString();

                        //Ripple(SMBAlert) 
                        input = arrayDetail[67].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.RippleSMBAlert = arrayresult[1].ToString();

                        //Vin
                        input = arrayDetail[2].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = "";
                        result = Regex.Replace(input, pattern, replacement);
                        Eff_Noise_Column.Vin = result.Split('=')[1].ToString();

                        //Frequ
                        input = arrayDetail[3].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = "";
                        result = Regex.Replace(input, pattern, replacement);
                        Eff_Noise_Column.Frequ = result.Split('=')[1].ToString();

                        //Vout_Min(12V) Vout_Max(12V)
                        input = arrayDetail[31].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Vout_Max12V = arrayresult[1].ToString();
                        Eff_Noise_Column.Vout_Min12V = arrayresult[2].ToString();

                        //Vout_Min(12Vsb) Vout_Max(12Vsb)
                        input = arrayDetail[32].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Vout_Max12Vsb = arrayresult[1].ToString();
                        Eff_Noise_Column.Vout_Min12Vsb = arrayresult[2].ToString();

                        //Vout_Min(PWOK) Vout_Max(PWOK)
                        input = arrayDetail[33].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Vout_MaxPWOK = arrayresult[1].ToString();
                        Eff_Noise_Column.Vout_MinPWOK = arrayresult[2].ToString();

                        //Vout_Min(Vin_Good) Vout_Max(Vin_Good)
                        input = arrayDetail[34].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Vout_MaxVin_Good = arrayresult[1].ToString();
                        Eff_Noise_Column.Vout_MinVin_Good = arrayresult[2].ToString();

                        //Vout_Min(SMBAlert) Vout_Max(SMBAlert)
                        input = arrayDetail[35].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Vout_MaxSMBAlert = arrayresult[1].ToString();
                        Eff_Noise_Column.Vout_MinSMBAlert = arrayresult[2].ToString();

                        //Load(12V)
                        input = arrayDetail[9].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Load12V = arrayresult[1].ToString();

                        //Load(12Vsb)
                        input = arrayDetail[10].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.Load12Vsb = arrayresult[1].ToString();

                        //Load(PWOK)
                        input = arrayDetail[11].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.LoadPWOK = arrayresult[1].ToString();

                        //Load(Vin_Good)
                        input = arrayDetail[12].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.LoadVin_Good = arrayresult[1].ToString();

                        //Load(SMBAlert)
                        input = arrayDetail[13].ToString().Trim();
                        //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                        pattern = @"\s+";
                        replacement = " ";
                        result = Regex.Replace(input, pattern, replacement);
                        arrayresult = result.Split(' ');
                        Eff_Noise_Column.LoadSMBAlert = arrayresult[1].ToString();

                        arrayList_EffNoise.Add
                        (
                              Eff_Noise_Column.SEQ + "|" + Eff_Noise_Column.Name + "|"
                            + Eff_Noise_Column.Ripple12V + "|" + Eff_Noise_Column.Ripple12Vsb + "|" + Eff_Noise_Column.RipplePWOK + "|" + Eff_Noise_Column.RippleVin_Good + "|" + Eff_Noise_Column.RippleSMBAlert + "|"
                            + Eff_Noise_Column.Vin + "|" + Eff_Noise_Column.Frequ + "|"
                            + Eff_Noise_Column.Vout_Max12V + "|" + Eff_Noise_Column.Vout_Max12Vsb + "|" + Eff_Noise_Column.Vout_MaxPWOK + "|" + Eff_Noise_Column.Vout_MaxVin_Good + "|" + Eff_Noise_Column.Vout_MaxSMBAlert + "|"
                            + Eff_Noise_Column.Vout_Min12V + "|" + Eff_Noise_Column.Vout_Min12Vsb + "|" + Eff_Noise_Column.Vout_MinPWOK + "|" + Eff_Noise_Column.Vout_MinVin_Good + "|" + Eff_Noise_Column.Vout_MinSMBAlert + "|"               
                            + Eff_Noise_Column.Load12V + "|" + Eff_Noise_Column.Load12Vsb + "|" + Eff_Noise_Column.LoadPWOK + "|" + Eff_Noise_Column.LoadVin_Good + "|" + Eff_Noise_Column.LoadSMBAlert
                            
                        );

                    }
                    #endregion

                    for (int j = 0; j < arrayDetail.Length; j++)
                    {                       
                        iCurrentRow = iCurrentRow + 1;                       
                        IRow row1 = sheet_ALL.CreateRow(iCurrentRow); 
                        ICell cell = row1.CreateCell(0);

                        if (arrayDetail[j].ToString().Contains("PASS"))
                        {
                            AL_CellIndex.Add(iCurrentRow + 1);
                            cell.CellStyle = stylePASS;
                            sheet_ALL.AddMergedRegion(new CellRangeAddress(iCurrentRow, iCurrentRow, 0, 7)); //MergedRegion=合併區
                        }
                        else if (arrayDetail[j].ToString().Contains("FAIL"))
                        {
                            AL_CellIndex.Add(iCurrentRow + 1);
                            cell.CellStyle = styleFAIL;
                            sheet_ALL.AddMergedRegion(new CellRangeAddress(iCurrentRow, iCurrentRow, 0, 7)); //MergedRegion=合併區
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
                        IRow space_row1 = sheet_ALL.CreateRow(iCurrentRow);
                        ICell space_cell = space_row1.CreateCell(0);
                        space_cell.SetCellValue("");
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(ALL)內容寫入失敗:" + ex.Message);
                return false;
            }

            //Sheet(Item)內容
            try
            {
                iCurrentRow = 0;
                for (int i = 1; i < arraystr.Length; i++)
                {
                    XSSFCellStyle Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    if (!arraystr[i].ToString().Contains("STEP"))
                    {
                        continue;
                    }
                    else
                    {
                        string sItem = arraystr[i].ToString().Trim();
                        sItem = sItem.Replace("\r\n", "  ");
                        IRow Item_row1 = sheet_Item.CreateRow(iCurrentRow);
                        ICell Item_cell = Item_row1.CreateCell(0);

                        int iPASSIndex = sItem.LastIndexOf("PASS");
                        int iFAILIndex = sItem.LastIndexOf("FAIL");

                        if(iPASSIndex != -1)
                        {
                            sItem = sItem.Substring(0, iPASSIndex + 4);
                        }
                        if(iFAILIndex != -1)
                        {
                            sItem = sItem.Substring(0, iFAILIndex + 4);
                        }
                        
                        Item_style.SetFont(hyperlink_font);

                        XSSFHyperlink link = new XSSFHyperlink(HyperlinkType.Document);
                        link.Address = "#ALL" + "!A" + AL_CellIndex[0].ToString();  //設置超連結跳轉的位址
                        AL_CellIndex.RemoveAt(0);
                        Item_cell.Hyperlink = link;

                        sheet_Item.AddMergedRegion(new CellRangeAddress(iCurrentRow, iCurrentRow, 0, 7));  //MergedRegion=合併區

                        Item_cell.SetCellValue(sItem);  //STEP.1(UUT Test seq.1) : Clear PROG Mode ---- (0'031) --------------------- PASS
                        Item_cell.CellStyle = Item_style;

                        //新增項目描述內容
                        XSSFCellStyle Des_style = (XSSFCellStyle)workbook.CreateCellStyle();
                        Des_style.SetFont(normal_font);
                        ICell Item_celltxt = Item_row1.CreateCell(10);
                        Item_celltxt.SetCellValue(_dt.Rows[0]["Value"].ToString());
                        Item_celltxt.CellStyle = Des_style;
                        _dt.Rows.RemoveAt(0);

                        iCurrentRow += 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(Item)內容寫入失敗:" + ex.Message);
                return false;
            }

            #region Sheet(Turn on)內容
            try
            {
                ISheet sheet_TurnON = workbook.CreateSheet("turn on");
                iCurrentRow = 0;
                for (int i = 0; i < arrayList_TurnON.Count; i++)
                {
                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 黑字體
                    XSSFCellStyle TurnON_Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    TurnON_Item_style.VerticalAlignment = VerticalAlignment.Center;
                    TurnON_Item_style.Alignment = HorizontalAlignment.Center;
                    TurnON_Item_style.BorderTop = BorderStyle.Medium;
                    TurnON_Item_style.BorderBottom = BorderStyle.Medium;
                    TurnON_Item_style.BorderLeft = BorderStyle.Medium;
                    TurnON_Item_style.BorderRight = BorderStyle.Medium;
                    TurnON_Item_style.WrapText = true;
                    TurnON_Item_style.SetFont(normal_font);

                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 藍字體
                    XSSFCellStyle TurnON_Data_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    TurnON_Data_style.VerticalAlignment = VerticalAlignment.Center;
                    TurnON_Data_style.Alignment = HorizontalAlignment.Center;
                    TurnON_Data_style.BorderTop = BorderStyle.Medium;
                    TurnON_Data_style.BorderBottom = BorderStyle.Medium;
                    TurnON_Data_style.BorderLeft = BorderStyle.Medium;
                    TurnON_Data_style.BorderRight = BorderStyle.Medium;
                    TurnON_Data_style.WrapText = true;
                    TurnON_Data_style.SetFont(normal_Blue_font);

                    string sTurnON_Data = arrayList_TurnON[i].ToString();
                    string[] arrayTurnON_Data = sTurnON_Data.Split('|');
                    IRow TurnON_row1 = sheet_TurnON.CreateRow(iCurrentRow);
                    TurnON_row1.HeightInPoints = 30;  //設定每個儲存格列高

                    for (int j = 0; j < arrayTurnON_Data.Length; j++)
                    {
                        ICell TurnON_cell = TurnON_row1.CreateCell(j);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_TurnON.SetColumnWidth(j, 14 * 256);  //設定每個儲存格欄寬 
                        if (j < 2 || i == 0)
                        {
                            TurnON_cell.CellStyle = TurnON_Item_style;
                        }
                        else
                        {
                            TurnON_cell.CellStyle = TurnON_Data_style;
                        }

                        //判斷字串是否為Double數值
                        double dNumber;

                        bool isDouble = double.TryParse(arrayTurnON_Data[j].ToString(), out dNumber);
                        if (isDouble)
                        {
                            /*ToString("0.##");在這個格式化字符串中，對於整數部分，'0' 表示一定要顯示，而對於小數部分，# 表示如果有數字則顯示，
                              否則不顯示，並自動將小數點後的零去掉。*/
                            TurnON_cell.SetCellValue(dNumber.ToString("0.##")); 
                        }
                        else
                        {
                            if (arrayTurnON_Data[j].ToString().Contains("*"))
                            {
                                TurnON_cell.SetCellValue(arrayTurnON_Data[j].ToString().Replace(arrayTurnON_Data[j].ToString(),"*"));
                            }
                            else
                            {
                                TurnON_cell.SetCellValue(arrayTurnON_Data[j].ToString());
                            }                           
                        }

                    }
                                                         
                    iCurrentRow += 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(turn on)內容寫入失敗:" + ex.Message);
                return false;
            }
            #endregion

            #region Sheet(Hold Up)內容
            try
            {
                ISheet sheet_HoldUp = workbook.CreateSheet("hold up");
                iCurrentRow = 0;
                for (int k = 0; k < arrayList_HoldUp.Count; k++)
                {
                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 黑字體
                    XSSFCellStyle HoldUp_Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    HoldUp_Item_style.VerticalAlignment = VerticalAlignment.Center;
                    HoldUp_Item_style.Alignment = HorizontalAlignment.Center;
                    HoldUp_Item_style.BorderTop = BorderStyle.Medium;
                    HoldUp_Item_style.BorderBottom = BorderStyle.Medium;
                    HoldUp_Item_style.BorderLeft = BorderStyle.Medium;
                    HoldUp_Item_style.BorderRight = BorderStyle.Medium;
                    HoldUp_Item_style.WrapText = true;
                    HoldUp_Item_style.SetFont(normal_font);

                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 藍字體
                    XSSFCellStyle HoldUp_Data_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    HoldUp_Data_style.VerticalAlignment = VerticalAlignment.Center;
                    HoldUp_Data_style.Alignment = HorizontalAlignment.Center;
                    HoldUp_Data_style.BorderTop = BorderStyle.Medium;
                    HoldUp_Data_style.BorderBottom = BorderStyle.Medium;
                    HoldUp_Data_style.BorderLeft = BorderStyle.Medium;
                    HoldUp_Data_style.BorderRight = BorderStyle.Medium;
                    HoldUp_Data_style.WrapText = true;
                    HoldUp_Data_style.SetFont(normal_Blue_font);

                    string sHoldUp_Data = arrayList_HoldUp[k].ToString();
                    string[] arrayHoldUp_Data = sHoldUp_Data.Split('|');
                    IRow HoldUp_row1 = sheet_HoldUp.CreateRow(iCurrentRow);
                    HoldUp_row1.HeightInPoints = 30;  //設定每個儲存格列高

                    for (int l = 0; l < arrayHoldUp_Data.Length; l++)
                    {
                        ICell HoldUp_cell = HoldUp_row1.CreateCell(l);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_HoldUp.SetColumnWidth(l, 14 * 256);  //設定每個儲存格欄寬 
                        if (l < 2 || k == 0)
                        {
                            HoldUp_cell.CellStyle = HoldUp_Item_style;
                        }
                        else
                        {
                            HoldUp_cell.CellStyle = HoldUp_Data_style;
                        }

                        //判斷字串是否為Double數值
                        double dNumber;

                        bool isDouble = double.TryParse(arrayHoldUp_Data[l].ToString(), out dNumber);
                        if (isDouble)
                        {
                            /*ToString("0.##");在這個格式化字符串中，對於整數部分，'0' 表示一定要顯示，而對於小數部分，# 表示如果有數字則顯示，
                              否則不顯示，並自動將小數點後的零去掉。*/
                            HoldUp_cell.SetCellValue(dNumber.ToString("0.##"));
                        }
                        else
                        {
                            if (arrayHoldUp_Data[l].ToString().Contains("*"))
                            {
                                HoldUp_cell.SetCellValue(arrayHoldUp_Data[l].ToString().Replace(arrayHoldUp_Data[l].ToString(), "*"));
                            }
                            else
                            {
                                HoldUp_cell.SetCellValue(arrayHoldUp_Data[l].ToString());
                            }
                        }

                    }

                    iCurrentRow += 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(hold up)內容寫入失敗:" + ex.Message);
                return false;
            }
            #endregion

            #region Sheet(PSOFF)內容
            try
            {
                ISheet sheet_PSOFF = workbook.CreateSheet("ps off");
                iCurrentRow = 0;
                for (int m = 0; m < arrayList_PSOFF.Count; m++)
                {
                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 黑字體
                    XSSFCellStyle PSOFF_Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    PSOFF_Item_style.VerticalAlignment = VerticalAlignment.Center;
                    PSOFF_Item_style.Alignment = HorizontalAlignment.Center;
                    PSOFF_Item_style.BorderTop = BorderStyle.Medium;
                    PSOFF_Item_style.BorderBottom = BorderStyle.Medium;
                    PSOFF_Item_style.BorderLeft = BorderStyle.Medium;
                    PSOFF_Item_style.BorderRight = BorderStyle.Medium;
                    PSOFF_Item_style.WrapText = true;
                    PSOFF_Item_style.SetFont(normal_font);

                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 藍字體
                    XSSFCellStyle PSOFF_Data_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    PSOFF_Data_style.VerticalAlignment = VerticalAlignment.Center;
                    PSOFF_Data_style.Alignment = HorizontalAlignment.Center;
                    PSOFF_Data_style.BorderTop = BorderStyle.Medium;
                    PSOFF_Data_style.BorderBottom = BorderStyle.Medium;
                    PSOFF_Data_style.BorderLeft = BorderStyle.Medium;
                    PSOFF_Data_style.BorderRight = BorderStyle.Medium;
                    PSOFF_Data_style.WrapText = true;
                    PSOFF_Data_style.SetFont(normal_Blue_font);

                    string sPSOFF_Data = arrayList_PSOFF[m].ToString();
                    string[] arrayPSOFF_Data = sPSOFF_Data.Split('|');
                    IRow PSOFF_row1 = sheet_PSOFF.CreateRow(iCurrentRow);
                    PSOFF_row1.HeightInPoints = 30;  //設定每個儲存格列高

                    for (int n = 0; n < arrayPSOFF_Data.Length; n++)
                    {
                        ICell PSOFF_cell = PSOFF_row1.CreateCell(n);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_PSOFF.SetColumnWidth(n, 14 * 256);  //設定每個儲存格欄寬 
                        if (n < 2 || m == 0)
                        {
                            PSOFF_cell.CellStyle = PSOFF_Item_style;
                        }
                        else
                        {
                            PSOFF_cell.CellStyle = PSOFF_Data_style;
                        }

                        //判斷字串是否為Double數值
                        double dNumber;

                        bool isDouble = double.TryParse(arrayPSOFF_Data[n].ToString(), out dNumber);
                        if (isDouble)
                        {
                            /*ToString("0.##");在這個格式化字符串中，對於整數部分，'0' 表示一定要顯示，而對於小數部分，# 表示如果有數字則顯示，
                              否則不顯示，並自動將小數點後的零去掉。*/
                            PSOFF_cell.SetCellValue(dNumber.ToString("0.##"));
                        }
                        else
                        {
                            if (arrayPSOFF_Data[n].ToString().Contains("*"))
                            {
                                PSOFF_cell.SetCellValue(arrayPSOFF_Data[n].ToString().Replace(arrayPSOFF_Data[n].ToString(), "*"));
                            }
                            else
                            {
                                PSOFF_cell.SetCellValue(arrayPSOFF_Data[n].ToString());
                            }
                        }

                    }

                    iCurrentRow += 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(ps off)內容寫入失敗:" + ex.Message);
                return false;
            }
            #endregion

            #region Sheet(PSON)內容
            try
            {
                ISheet sheet_PSON = workbook.CreateSheet("ps on");
                iCurrentRow = 0;
                for (int o = 0; o < arrayList_PSON.Count; o++)
                {
                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 黑字體
                    XSSFCellStyle PSON_Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    PSON_Item_style.VerticalAlignment = VerticalAlignment.Center;
                    PSON_Item_style.Alignment = HorizontalAlignment.Center;
                    PSON_Item_style.BorderTop = BorderStyle.Medium;
                    PSON_Item_style.BorderBottom = BorderStyle.Medium;
                    PSON_Item_style.BorderLeft = BorderStyle.Medium;
                    PSON_Item_style.BorderRight = BorderStyle.Medium;
                    PSON_Item_style.WrapText = true;
                    PSON_Item_style.SetFont(normal_font);

                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 藍字體
                    XSSFCellStyle PSON_Data_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    PSON_Data_style.VerticalAlignment = VerticalAlignment.Center;
                    PSON_Data_style.Alignment = HorizontalAlignment.Center;
                    PSON_Data_style.BorderTop = BorderStyle.Medium;
                    PSON_Data_style.BorderBottom = BorderStyle.Medium;
                    PSON_Data_style.BorderLeft = BorderStyle.Medium;
                    PSON_Data_style.BorderRight = BorderStyle.Medium;
                    PSON_Data_style.WrapText = true;
                    PSON_Data_style.SetFont(normal_Blue_font);

                    string sPSON_Data = arrayList_PSON[o].ToString();
                    string[] arrayPSON_Data = sPSON_Data.Split('|');
                    IRow PSON_row1 = sheet_PSON.CreateRow(iCurrentRow);
                    PSON_row1.HeightInPoints = 30;  //設定每個儲存格列高

                    for (int p = 0; p < arrayPSON_Data.Length; p++)
                    {
                        ICell PSON_cell = PSON_row1.CreateCell(p);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_PSON.SetColumnWidth(p, 14 * 256);  //設定每個儲存格欄寬 
                        if (p < 2 || o == 0)
                        {
                            PSON_cell.CellStyle = PSON_Item_style;
                        }
                        else
                        {
                            PSON_cell.CellStyle = PSON_Data_style;
                        }

                        //判斷字串是否為Double數值
                        double dNumber;

                        bool isDouble = double.TryParse(arrayPSON_Data[p].ToString(), out dNumber);
                        if (isDouble)
                        {
                            /*ToString("0.##");在這個格式化字符串中，對於整數部分，'0' 表示一定要顯示，而對於小數部分，# 表示如果有數字則顯示，
                              否則不顯示，並自動將小數點後的零去掉。*/
                            PSON_cell.SetCellValue(dNumber.ToString("0.##"));
                        }
                        else
                        {
                            if (arrayPSON_Data[p].ToString().Contains("*"))
                            {
                                PSON_cell.SetCellValue(arrayPSON_Data[p].ToString().Replace(arrayPSON_Data[p].ToString(), "*"));
                            }
                            else
                            {
                                PSON_cell.SetCellValue(arrayPSON_Data[p].ToString());
                            }
                        }

                    }

                    iCurrentRow += 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(ps on)內容寫入失敗:" + ex.Message);
                return false;
            }
            #endregion

            #region Sheet(Eff_Noise)內容
            try
            {
                ISheet sheet_EffNoise = workbook.CreateSheet("Eff_Noise");
                iCurrentRow = 0;
                for (int q = 0; q < arrayList_EffNoise.Count; q++)
                {
                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 黑字體
                    XSSFCellStyle EffNoise_Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    EffNoise_Item_style.VerticalAlignment = VerticalAlignment.Center;
                    EffNoise_Item_style.Alignment = HorizontalAlignment.Center;
                    EffNoise_Item_style.BorderTop = BorderStyle.Medium;
                    EffNoise_Item_style.BorderBottom = BorderStyle.Medium;
                    EffNoise_Item_style.BorderLeft = BorderStyle.Medium;
                    EffNoise_Item_style.BorderRight = BorderStyle.Medium;
                    EffNoise_Item_style.WrapText = true;
                    EffNoise_Item_style.SetFont(normal_font);

                    //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 藍字體
                    XSSFCellStyle EffNoise_Data_style = (XSSFCellStyle)workbook.CreateCellStyle();
                    EffNoise_Data_style.VerticalAlignment = VerticalAlignment.Center;
                    EffNoise_Data_style.Alignment = HorizontalAlignment.Center;
                    EffNoise_Data_style.BorderTop = BorderStyle.Medium;
                    EffNoise_Data_style.BorderBottom = BorderStyle.Medium;
                    EffNoise_Data_style.BorderLeft = BorderStyle.Medium;
                    EffNoise_Data_style.BorderRight = BorderStyle.Medium;
                    EffNoise_Data_style.WrapText = true;
                    EffNoise_Data_style.SetFont(normal_Blue_font);

                    string sEffNoise_Data = arrayList_EffNoise[q].ToString();
                    string[] arrayEffNoise_Data = sEffNoise_Data.Split('|');
                    IRow EffNoise_row1 = sheet_EffNoise.CreateRow(iCurrentRow);
                    EffNoise_row1.HeightInPoints = 30;  //設定每個儲存格列高

                    for (int r = 0; r < arrayEffNoise_Data.Length; r++)
                    {
                        ICell EffNoise_cell = EffNoise_row1.CreateCell(r);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_EffNoise.SetColumnWidth(r, 14 * 256);  //設定每個儲存格欄寬 
                        if (r < 2 || q == 0)
                        {
                            EffNoise_cell.CellStyle = EffNoise_Item_style;
                        }
                        else
                        {
                            EffNoise_cell.CellStyle = EffNoise_Data_style;
                        }

                        //判斷字串是否為Double數值
                        double dNumber;

                        bool isDouble = double.TryParse(arrayEffNoise_Data[r].ToString(), out dNumber);
                        if (isDouble)
                        {
                            /*ToString("0.##");在這個格式化字符串中，對於整數部分，'0' 表示一定要顯示，而對於小數部分，# 表示如果有數字則顯示，
                              否則不顯示，並自動將小數點後的零去掉。*/
                            EffNoise_cell.SetCellValue(dNumber.ToString("0.##"));
                        }
                        else
                        {
                            if (arrayEffNoise_Data[r].ToString().Contains("*"))
                            {
                                EffNoise_cell.SetCellValue(arrayEffNoise_Data[r].ToString().Replace(arrayEffNoise_Data[r].ToString(), "*"));
                            }
                            else
                            {
                                EffNoise_cell.SetCellValue(arrayEffNoise_Data[r].ToString());
                            }
                        }

                    }

                    iCurrentRow += 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sheet(Eff_Noise)內容寫入失敗:" + ex.Message);
                return false;
            }
            #endregion

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
