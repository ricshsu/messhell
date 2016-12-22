using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Ext.Net;
using System.Threading;
using System.Collections;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using NYPCB.SPC;
using System.Xml;
using System.Runtime.InteropServices;
using System.Configuration;

public partial class CTF_Transfer : System.Web.UI.Page
{
    public struct ctfObj
    {
        public string Part_Id;
        public string Meas_Item;
        public string Lot_Id;
        public string rowsn;
        public string rowData;
    }

    private struct txtMain
    {
        public string part_id;
        public string lot_id;
        public string tool_id;
        public string stage_light;
        public string coax_light;
        public string start_time;
        public string end_time;
    }

    private struct txtRowData
    {
        public string unit_id;
        public string item;
        public string dvalue;
    }

    private struct mailResult
    {
        public string fileName;
        public string reason;
    }

    string gfile = null;
    string gMoveFolder = null;
    string gConnDriver = null;
    string gdateTime = null;
    string gRootPath = null;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!X.IsAjaxRequest)
        {
            this.Session.Remove("FileResul");
            this.Session.Remove("FileTotal");
            this.Session.Remove("LongActionProgress");
            this.ta_result.Text = "";
        }
    }
    
    // 執行背景執行序
    protected void StartLongAction(object sender, DirectEventArgs e)
    {
        this.Session["LongActionProgress"] = 0;
        ThreadPool.QueueUserWorkItem(LongAction);
        this.ResourceManager1.AddScript("{0}.startTask('longactionprogress');", this.TaskManager1.ClientID);
    }

    private void LongAction(object state)
    {
        gdateTime = System.DateTime.Now.ToString("yyyy-MM-dd");
        XmlNode node = null;
        XmlDocument xd = new XmlDocument();
        xd.Load(Server.MapPath(".") + "\\CTF\\param.xml");

        // Get Connection Driver
        node = xd.SelectSingleNode("//db_driver");
        gConnDriver = (node.InnerText);

        // Get File Path
        node = xd.SelectSingleNode("//file_path");
        gfile = (node.InnerText);

        node = xd.SelectSingleNode("//Movefile_path");
        gMoveFolder = (node.InnerText);

        node = xd.SelectSingleNode("//RootPath");
        gRootPath = (node.InnerText);

        // Get File list 
        FileInfo fileInfo = default(FileInfo);
        string fileName = null;
        string fileSName = null;
        string copyPath = null;
        DataTable detaildt = null;
        ArrayList rAry = new ArrayList();
        mailResult rltObj = default(mailResult);
        ArrayList fileAry = new ArrayList();
        int startSub = 0;
        int SubLength = 0;

        System.Diagnostics.Process proc = System.Diagnostics.Process.Start("net", "use \\Nypcb\\C823 A124549701 /user:N000142679");
        proc.WaitForExit();

        //AddConnection(@"\\Nypcb\\C823", "A124549701", "T:");
        
        try {
            DirSearch(gfile, ref fileAry);
        }
        catch (Exception error) { this.ta_result.Text = error.Message; }
        
        this.Session["FileTotal"] = (fileAry.Count).ToString();
        ArrayList resultAry = new ArrayList();
        
        for (int i = 0; i < (fileAry.Count); i++)
        {
            
            fileInfo = new FileInfo(fileAry[i].ToString());
            rltObj = new mailResult();
            detaildt = new DataTable();
            detaildt.Columns.Add("TxtData", Type.GetType("System.String"));
            fileSName = fileInfo.Name;
            fileName = fileInfo.FullName;

            try 
            {
                startSub = ((fileName.LastIndexOf(gRootPath) + gRootPath.Length) + 1);
                SubLength = (fileName.IndexOf(fileSName) - startSub - 1);

                if (SubLength < 0)
                {
                    copyPath = gdateTime + "\\";
                }
                else
                {
                    copyPath = gdateTime + "\\" + fileName.Substring(startSub, SubLength);
                }
                readFile(fileName, ref detaildt);

                txtMain mObj = new txtMain();

                // (放資料到DB) & (計算統計值) 
                if (insertDB(detaildt, ref mObj, fileName, copyPath, fileSName))
                {
                    insertDiameter(ref mObj);
                    CalStatistics(ref mObj);
                    System.IO.Directory.CreateDirectory(gMoveFolder + copyPath);
                    try
                    {
                        if (File.Exists((gMoveFolder + copyPath + "\\" + fileSName)))
                        {
                            File.Delete((gMoveFolder + copyPath + "\\" + fileSName));
                        }
                        System.IO.File.Move(fileName, (gMoveFolder + copyPath + "\\" + fileSName));
                    }
                    catch (Exception Error) { }
                    resultAry.Add((fileName + ":成功"));
                }
                else
                {
                    rltObj.fileName = fileName;
                    rAry.Add(rltObj);
                    resultAry.Add((fileName + ":失敗"));
                }
             
            } catch(Exception error) {
                rltObj.fileName = fileName;
                rAry.Add(rltObj);
                resultAry.Add((fileName + ":失敗"));
            }

            this.Session["LongActionProgress"] = i;
            this.Session["FileResul"] = resultAry;
        }

        //CancelConnection("T:", 0);
        proc = System.Diagnostics.Process.Start("net", "use \\Nypcb\\C823 /del");
        proc.WaitForExit();
        
        this.Session.Remove("FileTotal");
        this.Session.Remove("LongActionProgress");
    }

    protected void RefreshProgress(object sender, DirectEventArgs e)
    {
        ArrayList fileResul = (ArrayList)this.Session["FileResul"];
        object progress = this.Session["LongActionProgress"];

        if (progress != null)
        {
            try
            {
                double fileTotal = Convert.ToDouble(this.Session["FileTotal"]);
                Progress1.Hidden = false;
                this.Progress1.UpdateProgress((((int)progress) / (float)fileTotal), string.Format("Step {0} of {1} ...", progress.ToString(), fileTotal.ToString()));
                if (fileResul != null)
                {
                    this.ta_result.Text = "";
                    for (int i = 0; i < fileResul.Count; i++)
                    {
                        this.ta_result.Text += ((String)fileResul[i]) + "\n";
                    }
                }
            }
            catch (Exception error) {}
        }
        else
        {
            this.ResourceManager1.AddScript("{0}.stopTask('longactionprogress');", this.TaskManager1.ClientID);
            this.Progress1.UpdateProgress(1, "All finished!");
        }
    }

    #region "File 轉置"

    // Read TXT data
    private void readFile(string fileName, ref DataTable dt)
    {
        StreamReader file = null;
        string line = null;
        DataRow dtr = null;

        try
        {
            file = new StreamReader(fileName, System.Text.Encoding.Default);
            line = file.ReadLine();
            while (line != null)
            {
                dtr = dt.NewRow();
                dtr[0] = line;
                dt.Rows.Add(dtr);
                line = file.ReadLine();
            }
            file.Close();
        }
        catch (Exception ex)
        {
            dt = new DataTable();
        }

    }

    // Insert RawData to DB
    private bool insertDB(DataTable dt1, ref txtMain mObj, string fileName, string copyPath, string fileSName)
    {
        bool status = false;
        string lowStr = "";
        string tmpStr = "";
        string comStr = "";
        string sqlStr = "";
        SqlConnection conn = null;
        SqlCommand comm = null;
        SqlTransaction transaction = null;
        txtRowData dObj = default(txtRowData);
        ArrayList oAry = new ArrayList();

        try
        {
            string[] strAry = null;
            int jugInt = 0;
            int lint = 0;

            foreach (DataRow dtr in dt1.Rows)
            {
                jugInt = 0;
                tmpStr = dtr[0].ToString();
                lowStr = tmpStr.ToLower();

                jugInt = lowStr.IndexOf("lot_id");
                if ((jugInt >= 0))
                {
                    strAry = tmpStr.Split(new char[] { '=' });
                    mObj.lot_id = (strAry[1]).Trim().Replace("-", "");
                }

                jugInt = lowStr.IndexOf("part_number");
                if ((jugInt >= 0))
                {
                    strAry = tmpStr.Split(new char[] { '=' });
                    mObj.part_id = (strAry[1]).Trim();
                }

                jugInt = lowStr.IndexOf("tool");
                if ((jugInt >= 0))
                {
                    strAry = tmpStr.Split(new char[] { ':' });
                    mObj.tool_id = (strAry[1]).Trim();
                }

                jugInt = lowStr.IndexOf("stage");
                if ((jugInt >= 0))
                {
                    strAry = tmpStr.Split(new char[] { '=' });
                    mObj.stage_light = "0" + (strAry[1]).Trim();
                }

                jugInt = lowStr.IndexOf("coax");
                if ((jugInt >= 0))
                {
                    strAry = tmpStr.Split(new char[] { '=' });
                    mObj.coax_light = "0" + (strAry[1]).Trim();
                }

                jugInt = lowStr.IndexOf("meas. start");

                if ((jugInt >= 0))
                {
                    string timeStr = tmpStr;
                    lint = 0;
                    tmpStr = tmpStr.Substring(11, (tmpStr.Length - 11));
                    lint = tmpStr.LastIndexOf(":");
                    tmpStr = tmpStr.Substring(0, (lint + 3));
                    if (timeStr.IndexOf("下午") >= 0)
                    {
                        System.DateTime dateValue = default(System.DateTime);
                        if (DateTime.TryParse(tmpStr, out dateValue))
                        {
                            dateValue = dateValue.AddHours(12);
                            tmpStr = dateValue.ToString("yyyy/MM/dd HH:mm:ss");
                        }
                    }
                    mObj.start_time = tmpStr.Trim();
                }
                jugInt = lowStr.IndexOf("meas. end");

                if ((jugInt >= 0))
                {
                    string timeStr = tmpStr;
                    lint = 0;
                    tmpStr = tmpStr.Substring(9, (tmpStr.Length - 9));
                    lint = tmpStr.LastIndexOf(":");
                    tmpStr = tmpStr.Substring(0, (lint + 3));
                    if (timeStr.IndexOf("下午") >= 0)
                    {
                        System.DateTime dateValue = default(System.DateTime);
                        if (DateTime.TryParse(tmpStr, out dateValue))
                        {
                            dateValue = dateValue.AddHours(12);
                            tmpStr = dateValue.ToString("yyyy/MM/dd HH:mm:ss");
                        }
                    }
                    mObj.end_time = tmpStr.Trim();
                }

                // --- Row Data ---
                jugInt = tmpStr.IndexOf(",");
                if ((jugInt > 0))
                {
                    strAry = tmpStr.Split(new char[]{','});
                    if (strAry.Length >= 3)
                    {
                        dObj = new txtRowData();
                        dObj.unit_id = (strAry[0].Trim());
                        dObj.item = (strAry[1].Trim());
                        dObj.dvalue = (strAry[2].Trim());
                        oAry.Add(dObj);
                    }
                }

            }

        }
        catch (Exception ex)
        {
            return false;
        }

        // --- Insert Into DB ---
        comStr = "MERGE INTO CTF_Monitor_Performance_RawData as t_fv " +
                 "USING (select '{0}' PART_ID, '{1}' LOT_ID, '{2}' MACHINE_ID, '{3}' MEAS_ITEM, {4} UNIT_ID) s_fv  " +
                 "ON 1=1 " +
                 "AND t_fv.PART_ID = s_fv.PART_ID " +
                 "AND t_fv.LOT_ID = s_fv.LOT_ID " +
                 "AND t_fv.MACHINE_ID = s_fv.MACHINE_ID " +
                 "AND t_fv.MEAS_ITEM = s_fv.MEAS_ITEM " +
                 "AND t_fv.UNIT_ID = s_fv.UNIT_ID " +
                 "WHEN MATCHED THEN " +
                 "update set t_fv.Data_Value={5}, t_fv.stage_Light={6}, t_fv.Coax_Light={7}, t_fv.Lot_Meas_Start_DataTime='{8}', t_fv.Lot_Meas_End_DataTime='{9}', t_fv.updateTime='{10}' " +
                 "WHEN NOT MATCHED THEN " +
                 "insert(Part_Id, Lot_Id, Machine_Id, Lot_Meas_Start_DataTime, Lot_Meas_End_DataTime, Stage_Light, Coax_Light, Unit_Id, Meas_Item, Data_Value, updateTime) " +
                 "values('{11}', '{12}', '{13}', '{14}', '{15}', {16}, {17}, {18}, '{19}', {20}, '{21}'); ";

        if (oAry.Count > 0)
        {
            try
            {
                conn = new SqlConnection(gConnDriver);
                conn.Open();
                comm = conn.CreateCommand();
                transaction = conn.BeginTransaction();
                comm.Transaction = transaction;

                for (int i = 0; i <= (oAry.Count - 1); i++)
                {
                    dObj = (txtRowData)oAry[i];
                    sqlStr = string.Format(comStr, 
                    (mObj.part_id), (mObj.lot_id), (mObj.tool_id), (dObj.item), (dObj.unit_id),
                    (dObj.dvalue), (mObj.stage_light), (mObj.coax_light), (mObj.start_time), (mObj.end_time), gdateTime,
                    (mObj.part_id), (mObj.lot_id), (mObj.tool_id), (mObj.start_time), (mObj.end_time), 
                    (mObj.stage_light), (mObj.coax_light), (dObj.unit_id), (dObj.item),
                    (dObj.dvalue), gdateTime);
                    comm.CommandText = sqlStr;
                    comm.ExecuteNonQuery();
                }

                transaction.Commit();
                conn.Close();
                status = true;

            }
            catch (System.Data.SqlClient.SqlException sqlex)
            {
                if (sqlex.Number == 2627)
                {
                    try 
                    {
                        System.IO.Directory.CreateDirectory(gMoveFolder + copyPath);
                        if (File.Exists((gMoveFolder + copyPath + "\\" + fileSName)))
                        {
                            File.Delete((gMoveFolder + copyPath + "\\" + fileSName));
                        }
                        System.IO.File.Move(fileName, (gMoveFolder + copyPath + "\\" + fileSName));
                    } catch (Exception error) {}
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }

        }

        return status;
    }

    // 為了要在 Insert Diameter 開頭的資料, 所以再 Insert 一次.
    private void insertDiameter(ref txtMain mObj)
    {
        
        SqlConnection conn = null;
        SqlCommand comm = null;

        string delStr = "Delete from CTF_Monitor_Performance_RawData " +
                        "where 1=1 " +
                        "and Part_Id='{0}' " +
                        "and Lot_Id='{1}' " +
                        "and Machine_Id='{2}' " +
                        "and Meas_Item='Diameter' ";
        delStr = string.Format(delStr, (mObj.part_id), (mObj.lot_id), (mObj.tool_id));

        string sqlStr = "";
        sqlStr += "insert into CTF_Monitor_Performance_RawData ";
        sqlStr += "(Part_Id, Lot_Id, Machine_Id, Lot_Meas_Start_DataTime, Lot_Meas_End_DataTime, updateTime, Stage_Light, Coax_Light, ";
        sqlStr += "Meas_Item, Unit_Id, Data_Value) ";
        sqlStr += "select part_id, lot_id, machine_id, lot_meas_start_datatime, lot_meas_end_datatime, updateTime, ";
        sqlStr += "Stage_Light, Coax_Light, ";
        sqlStr += "'Diameter', Unit_Id, avg(Data_Value) ";
        sqlStr += "from dbo.CTF_Monitor_Performance_RawData ";
        sqlStr += "where 1=1 ";
        sqlStr += "and Part_Id='{0}' ";
        sqlStr += "and Lot_Id='{1}' ";
        sqlStr += "and Machine_Id='{2}' ";
        sqlStr += "and Meas_Item like 'Diameter%' ";
        sqlStr += "group by part_id, lot_id, machine_id, lot_meas_start_datatime, lot_meas_end_datatime, updateTime, ";
        sqlStr += "Stage_Light, Coax_Light, Unit_Id ";
        sqlStr = string.Format(sqlStr, (mObj.part_id), (mObj.lot_id), (mObj.tool_id));

        try
        {
            conn = new SqlConnection(gConnDriver);
            conn.Open();
            // Delete
            comm = conn.CreateCommand();
            comm.CommandText = delStr;
            comm.ExecuteNonQuery();
            // Insert 
            comm = conn.CreateCommand();
            comm.CommandText = sqlStr;
            comm.ExecuteNonQuery();
            conn.Close();
        }
        catch (Exception ex)
        { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

    }

    // 計算 Statistics
    private void CalStatistics(ref txtMain mObj)
    {
        SqlConnection connRead = null;
        SqlDataAdapter sAdpt = null;
        DataTable dtRaw = null;
        string sqlStr = "";

        try
        {
            sqlStr += "SELECT x.Data_Value, y.USL, y.CL, y.LSL, x.Meas_Item, x.Unit_Id ";
            sqlStr += "FROM ( ";
            sqlStr += "select a.lot_id, a.Machine_Id, a.Unit_Id, a.Meas_Item, a.Data_Value, a.Stage_Light, a.Coax_Light, ";
            sqlStr += "a.Lot_Meas_Start_DataTime, a.Lot_Meas_End_DataTime, a.Part_Id ";
            sqlStr += "from dbo.CTF_Monitor_Performance_RawData a ";
            sqlStr += "where 1 = 1 ";
            sqlStr += "and a.Part_Id = '{0}' ";
            sqlStr += "and a.lot_id = '{1}' ";
            sqlStr += ") x left join CTF_Monitor_Performance_SPEC y ";
            sqlStr += "ON x.Meas_Item = y.Meas_Item ";
            sqlStr += "AND x.Part_Id = y.Part_Id ";
            sqlStr += "order by Meas_Item, Unit_Id ";
            sqlStr = string.Format(sqlStr, mObj.part_id, mObj.lot_id);

            connRead = new SqlConnection(gConnDriver);
            connRead.Open();
            sAdpt = new SqlDataAdapter(sqlStr, connRead);
            dtRaw = new DataTable();
            sAdpt.Fill(dtRaw);
            connRead.Close();

            decimal[] arrSeries = null;
            decimal usl = default(decimal);
            decimal lsl = default(decimal);
            string FMeas_Item = "";
            string BMeas_Item = "";
            ArrayList rawAry = new ArrayList();


            for (int i = 0; i < (dtRaw.Rows.Count); i++)
            {
                if (i == 0)
                {
                    FMeas_Item = (String)dtRaw.Rows[i]["Meas_Item"];
                }
                BMeas_Item = (String)dtRaw.Rows[i]["Meas_Item"];


                if (FMeas_Item != BMeas_Item)
                {
                    // 計算統計值
                    arrSeries = new decimal[rawAry.Count];
                    for (int j = 0; j < (rawAry.Count); j++)
                    {
                        arrSeries[j] = Convert.ToDecimal(rawAry[j]);
                    }
                    if (!((FMeas_Item.ToLower() == "diameter_1") | (FMeas_Item.ToLower() == "diameter_2") | (FMeas_Item.ToLower() == "diameter_3") | (FMeas_Item.ToLower() == "diameter_4")))
                    {
                        InsertStatistics(ref mObj, ref arrSeries, FMeas_Item, usl, lsl);
                    }
                    FMeas_Item = BMeas_Item;
                    rawAry = new ArrayList();
                }

                if ((dtRaw.Rows[i]["Data_Value"]) != System.DBNull.Value)
                {
                    rawAry.Add(dtRaw.Rows[i]["Data_Value"]);
                }

                usl = -9999;
                if ((dtRaw.Rows[i]["USL"]) != System.DBNull.Value)
                {
                    usl = (Decimal)dtRaw.Rows[i]["USL"];
                }

                lsl = -9999;
                if ((dtRaw.Rows[i]["LSL"]) != System.DBNull.Value)
                {
                    lsl = (Decimal)dtRaw.Rows[i]["LSL"];
                }

            }

            // 計算統計值 -- 最後一筆
            arrSeries = new decimal[rawAry.Count];
            for (int j = 0; j < (rawAry.Count); j++)
            {
                arrSeries[j] = Convert.ToDecimal(rawAry[j]);
            }
            if (!((FMeas_Item.ToLower() == "diameter_1") | (FMeas_Item.ToLower() == "diameter_2") | (FMeas_Item.ToLower() == "diameter_3") | (FMeas_Item.ToLower() == "diameter_4")))
            {
                InsertStatistics(ref mObj, ref arrSeries, BMeas_Item, usl, lsl);
            }
        }
        catch (Exception ex)
        {
        }
        finally
        {
            if (connRead.State == ConnectionState.Open)
            {
                connRead.Close();
            }
        }
    }

    // Insert Statistics Data to DataBase
    private void InsertStatistics(ref txtMain mObj, ref decimal[] decimalAry, string MItem, decimal usl, decimal lsl)
    {
        
        SPCFormula objFormula = null;
        SPCStatisc objStatisc = null;
        SqlConnection conn = null;
        SqlCommand comm = null;
        string sqlStr = "";
        string decimalAryStr = "";

        for (int i = 0; i < (decimalAry.Length); i++)
        {
            decimalAryStr += Math.Round(decimalAry[i], 5).ToString() + "|";
        }
        decimalAryStr = decimalAryStr.Substring(0, decimalAryStr.Length - 1);

        try
        {
            if ((usl != -9999) || (lsl != -9999))
            {
                objFormula = new SPCFormula(decimalAry, usl, lsl, null);
                objStatisc = objFormula.Calculate();
                sqlStr += "MERGE INTO CTF_Monitor_Performance_Lot_Summary as t_fv ";
                sqlStr += "USING (select '{0}' PART_ID, '{1}' LOT_ID, '{2}' MACHINE_ID, '{3}' MEAS_ITEM) s_fv  " ;
                sqlStr += "ON 1=1 ";
                sqlStr += "AND t_fv.PART_ID = s_fv.PART_ID ";
                sqlStr += "AND t_fv.LOT_ID = s_fv.LOT_ID ";
                sqlStr += "AND t_fv.MACHINE_ID = s_fv.MACHINE_ID ";
                sqlStr += "AND t_fv.MEAS_ITEM = s_fv.MEAS_ITEM ";
                sqlStr += "WHEN MATCHED THEN ";
                sqlStr += "update set t_fv.Lot_Meas_Start_DataTime='{4}', t_fv.Lot_Meas_End_DataTime='{5}', t_fv.Stage_Light={6}, t_fv.Coax_Light={7}, t_fv.USL={8}, t_fv.LSL={9}, ";
                sqlStr += "t_fv.Mean_Value={10}, t_fv.Std_Value={11}, t_fv.Min_Value={12}, t_fv.Max_Value={13}, t_fv.Cp={14}, t_fv.Cpk={15}, t_fv.CSL={16}, t_fv.RowData='{17}', t_fv.UpdateTime='{18}' ";
                sqlStr += "WHEN NOT MATCHED THEN ";
                sqlStr += "insert ( ";
                sqlStr += "Part_Id, Lot_Id, Machine_Id, Lot_Meas_Start_DataTime, Lot_Meas_End_DataTime, UpdateTime, Meas_Item, Stage_Light, Coax_Light, USL, LSL, ";
                sqlStr += "Mean_Value, Std_Value, Min_Value, Max_Value, Cp, Cpk, CSL, RowData";
                sqlStr += ") values('{19}', '{20}', '{21}', '{22}', '{23}', '{24}', '{25}', {26}, {27}, {28}, {29}, ";
                sqlStr += "{30}, {31}, {32}, {33}, {34}, {35}, {36}, '{37}');";
                sqlStr = string.Format( sqlStr,
                (mObj.part_id), (mObj.lot_id), (mObj.tool_id), (MItem), 
                (mObj.start_time), (mObj.end_time), (mObj.stage_light), (mObj.coax_light), usl, lsl,
                objStatisc.xMean, objStatisc.Sigma, objStatisc.Minimum, objStatisc.Maximum, objStatisc.Cp, objStatisc.Cpk, objStatisc.Target, decimalAryStr, gdateTime,
                mObj.part_id, mObj.lot_id, mObj.tool_id, mObj.start_time, mObj.end_time, gdateTime, MItem, mObj.stage_light, mObj.coax_light,
                usl, lsl, objStatisc.xMean, objStatisc.Sigma, objStatisc.Minimum, objStatisc.Maximum, objStatisc.Cp, objStatisc.Cpk, objStatisc.Target, decimalAryStr
                );
            }
            else
            {
                sqlStr += "MERGE INTO CTF_Monitor_Performance_Lot_Summary as t_fv ";
                sqlStr += "USING (select '{0}' PART_ID, '{1}' LOT_ID, '{2}' MACHINE_ID, '{3}' MEAS_ITEM) s_fv  ";
                sqlStr += "ON 1=1 ";
                sqlStr += "AND t_fv.PART_ID = s_fv.PART_ID ";
                sqlStr += "AND t_fv.LOT_ID = s_fv.LOT_ID ";
                sqlStr += "AND t_fv.MACHINE_ID = s_fv.MACHINE_ID ";
                sqlStr += "AND t_fv.MEAS_ITEM = s_fv.MEAS_ITEM ";
                sqlStr += "WHEN MATCHED THEN ";
                sqlStr += "update set t_fv.Lot_Meas_Start_DataTime='{4}', t_fv.Lot_Meas_End_DataTime='{5}', t_fv.Stage_Light={6}, t_fv.Coax_Light={7}, t_fv.RowData='{8}', t_fv.UpdateTime='{9}' ";
                sqlStr += "WHEN NOT MATCHED THEN ";
                sqlStr += "insert (Part_Id, Lot_Id, Machine_Id, Lot_Meas_Start_DataTime, Lot_Meas_End_DataTime, UpdateTime, Meas_Item, Stage_Light, Coax_Light, RowData) ";
                sqlStr += "values('{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17}, {18}, '{19}');";
                
                sqlStr = string.Format(sqlStr,
                (mObj.part_id), (mObj.lot_id), (mObj.tool_id), (MItem),
                (mObj.start_time), (mObj.end_time), (mObj.stage_light), (mObj.coax_light), decimalAryStr, gdateTime,
                (mObj.part_id), (mObj.lot_id), (mObj.tool_id), (mObj.start_time), (mObj.end_time), gdateTime, MItem, (mObj.stage_light), (mObj.coax_light), decimalAryStr
                );
            }

            try
            {
                conn = new SqlConnection(gConnDriver);
                conn.Open();
                comm = conn.CreateCommand();
                comm.CommandText = sqlStr;
                comm.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex){}
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }
        catch (Exception ex){}

    }

    // 檔案尋找
    private void DirSearch(string sDir, ref ArrayList fileAry)
    {

        try
        {
            DirectoryInfo dInfo = new DirectoryInfo(sDir);
            if ((dInfo.Name.ToUpper()) == "FINISH")
            {
                return;
            }
            // --- Read Folder ---
            foreach (string d in Directory.GetDirectories(sDir))
            {
                DirSearch(d, ref fileAry);
            }
            // --- Read File ---
            foreach (string f in Directory.GetFiles(sDir, "*.txt"))
            {
                fileAry.Add(f);
            }
        }
        catch (Exception ex)
        {
        }
    }

    #endregion


    #region "連線網路磁碟機"
    
    [DllImport("mpr.dll", EntryPoint = "WNetAddConnectionA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
    public static extern long WNetAddConnection(string lpszNetPath, string lpszPassword, string lpszLocalName);

    [DllImport("mpr.dll", EntryPoint = "WNetCancelConnectionA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
    public static extern long WNetCancelConnection(string lpszName, long bForce);

    public long AddConnection(string path, string pwd, string pathNo)
    {
        try
        {
            return WNetAddConnection(path, pwd, pathNo);
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    public long CancelConnection(string pathNo, int Force)
    {
        try
        {
            return WNetCancelConnection(pathNo, Force);
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    public enum ResourceScope
    {
        RESOURCE_CONNECTED = 1,
        RESOURCE_GLOBALNET,
        RESOURCE_REMEMBERED,
        RESOURCE_RECENT,
        RESOURCE_CONTEXT
    }

    public enum ResourceType
    {
        RESOURCETYPE_ANY,
        RESOURCETYPE_DISK,
        RESOURCETYPE_PRINT,
        RESOURCETYPE_RESERVED
    }

    public enum ResourceUsage
    {
        RESOURCEUSAGE_CONNECTABLE = 0x00000001,
        RESOURCEUSAGE_CONTAINER = 0x00000002,
        RESOURCEUSAGE_NOLOCALDEVICE = 0x00000004,
        RESOURCEUSAGE_SIBLING = 0x00000008,
        RESOURCEUSAGE_ATTACHED = 0x00000010,
        RESOURCEUSAGE_ALL = (RESOURCEUSAGE_CONNECTABLE | RESOURCEUSAGE_CONTAINER | RESOURCEUSAGE_ATTACHED),
    }

    public enum ResourceDisplayType
    {
        RESOURCEDISPLAYTYPE_GENERIC,
        RESOURCEDISPLAYTYPE_DOMAIN,
        RESOURCEDISPLAYTYPE_SERVER,
        RESOURCEDISPLAYTYPE_SHARE,
        RESOURCEDISPLAYTYPE_FILE,
        RESOURCEDISPLAYTYPE_GROUP,
        RESOURCEDISPLAYTYPE_NETWORK,
        RESOURCEDISPLAYTYPE_ROOT,
        RESOURCEDISPLAYTYPE_SHAREADMIN,
        RESOURCEDISPLAYTYPE_DIRECTORY,
        RESOURCEDISPLAYTYPE_TREE,
        RESOURCEDISPLAYTYPE_NDSCONTAINER
    }

    [StructLayout(LayoutKind.Sequential)]
    private class NETRESOURCE
    {
        public ResourceScope dwScope = 0;
        public ResourceType dwType = 0;
        public ResourceDisplayType dwDisplayType = 0;
        public ResourceUsage dwUsage = 0;
        public string lpLocalName = null;
        public string lpRemoteName = null;
        public string lpComment = null;
        public string lpProvider = null;
    }

    #endregion

    // 資料轉置, 由 EDA TO X60
    protected void but_Transfer_Click(object sender, EventArgs e)
    {
        SqlConnection conn = null;
        SqlDataAdapter sadp = null;
        SqlCommand comm = null;
        SqlTransaction transaction = null;
        ArrayList objAry = new ArrayList();
        ctfObj ctfobj;

        // --- 先抓取資料 ---
        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();
            DataTable dt = new DataTable();
            String sqlStr = "select a.Part_Id, a.Meas_Item, a.Lot_Id, b.rowsn, a.RowData from ";
            sqlStr += "( select Part_Id, Meas_Item, Lot_Id, RowData, ROW_NUMBER() Over (Partition By Part_Id, Meas_Item, Lot_Id Order By UpdateTime Desc) As One ";
            sqlStr += "from CTF_Monitor_Performance_Lot_Summary ) a, [x60LinkInsert].[x60].[dbo].[Nypcb5_CTF1] b ";
            sqlStr += "where a.One = 1 ";
            sqlStr += "and a.Part_Id = b.part ";
            sqlStr += "and a.Meas_Item = b.Mitem ";
            sqlStr += "GROUP BY a.Part_Id, a.Meas_Item, a.Lot_Id, b.rowsn, a.RowData ";
            sqlStr += "ORDER BY a.Part_Id, a.Lot_Id, a.Meas_Item";
            sadp = new SqlDataAdapter(sqlStr, conn);
            sadp.Fill(dt);
            conn.Close();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ctfobj = new ctfObj();
                try
                {
                    ctfobj.Part_Id = dt.Rows[i]["Part_Id"].ToString();
                    ctfobj.Meas_Item = dt.Rows[i]["Meas_Item"].ToString();
                    ctfobj.Lot_Id = dt.Rows[i]["Lot_Id"].ToString();
                    ctfobj.rowsn = dt.Rows[i]["rowsn"].ToString();

                    // 處理 Insert Mester 資料
                    String[] rowAry = (dt.Rows[i]["RowData"].ToString()).Split('|');
                    String tmpStr = "";
                    for (int j = 0; j < rowAry.Length; j++)
                    {
                        tmpStr += (j + 1) + ":" + (rowAry[j].ToString()) + ";";
                    }
                    tmpStr = tmpStr.Substring(0, (tmpStr.Length - 1));
                    ctfobj.rowData = tmpStr;
                }
                catch (Exception inError) { }
                objAry.Add(ctfobj);
            }
        }
        catch (Exception error){}
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

        String insertDetail = "insert into [x60LinkInsert].[x60].[dbo].[CTF_Detail] ";
        insertDetail += "SELECT a.Lot_Id, a.Unit_Id, a.Part_Id, b.rowsn, b.rowsn, '', '', a.Unit_Id, a.Data_Value, null, 9, '' ";
        insertDetail += "FROM dbo.CTF_Monitor_Performance_RawData a, [x60LinkInsert].[x60].[dbo].[Nypcb5_CTF1] b ";
        insertDetail += "where 1=1 ";
        insertDetail += "and a.Part_Id = b.Part ";
        insertDetail += "and a.Meas_Item = b.Mitem ";
        insertDetail += "and a.Part_Id = '{0}' ";
        insertDetail += "and a.Meas_Item = '{1}' ";
        insertDetail += "and b.rowsn = '{2}' ";
        insertDetail += "and a.Lot_Id = '{3}' ";

        String insertMester = "Insert into [x60LinkInsert].[x60].[dbo].[CTF_Master](Lot, Part, Rowsn, trtm, LotMean, LotStd, LotMax, LotMin, LotR, Cp, Cpk, LSL, USL, Panel, msMan, msMch, Mvalue) ";
        insertMester += "SELECT a.Lot_Id, a.Part_Id, b.ROWSN, A.Lot_Meas_End_DataTime, a.Mean_Value, a.Std_Value, a.Max_Value, a.Min_Value, a.Stage_Light, ";
        insertMester += "a.Cp, a.Cpk, a.LSL, a.USL, 9, '', '', '{0}' ";
        insertMester += "FROM dbo.CTF_Monitor_Performance_Lot_Summary a, [x60LinkInsert].[x60].[dbo].[Nypcb5_CTF1] b ";
        insertMester += "WHERE 1=1 ";
        insertMester += "and a.Part_Id = b.Part ";
        insertMester += "and a.Meas_Item = b.Mitem ";
        insertMester += "AND a.Part_Id='{1}' ";
        insertMester += "AND a.Meas_Item='{2}' ";
        insertMester += "and b.rowsn='{3}' ";
        insertMester += "AND Lot_Id='{4}'";
        
        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();
            String sqlStr = "";
            for (int i = 0; i < objAry.Count; i++)
            {
                ctfobj = (ctfObj)objAry[i];
                try 
                {
                    // --- Insert to Deatil ---
                    sqlStr = "";
                    sqlStr = String.Format(insertDetail, ctfobj.Part_Id, ctfobj.Meas_Item, ctfobj.rowsn, ctfobj.Lot_Id);
                    comm = conn.CreateCommand();
                    comm.CommandText = sqlStr;
                    comm.ExecuteNonQuery();
                    // --- Insert to Master ---
                    sqlStr = "";
                    sqlStr = String.Format(insertMester, ctfobj.rowData, ctfobj.Part_Id, ctfobj.Meas_Item, ctfobj.rowsn, ctfobj.Lot_Id);
                    comm = conn.CreateCommand();
                    comm.CommandText = sqlStr;
                    comm.ExecuteNonQuery();
                } catch(Exception inErr){}
            }
            transaction.Commit();
            conn.Close();
        }
        catch (Exception error){}
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }
}