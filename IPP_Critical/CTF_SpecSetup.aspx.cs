using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using Ext.Net;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Collections;
using NYPCB.SPC;
using System.Text;

public partial class CTF_SpecSetup : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        
        if (!this.IsPostBack) 
        {
            //ViewState["partid"] = "JYKB33";
            if (Request["PART"] != null)
            {
                this.lab_part.Text = Request["PART"].ToString();
            }
        }

        if (!X.IsAjaxRequest)
        {
            this.Session["TestPersons"] = null;
        }
        this.BindData();
    }

    public class MeasSpec
    {
        public int? Id
        {
            get;
            set;
        }

        public String part_id
        {
            get;
            set;
        }

        public String meas_item
        {
            get;
            set;
        }

        public double USL
        {
            get;
            set;
        }

        public double CL
        {
            get;
            set;
        }

        public double LSL
        {
            get;
            set;
        }

        public String SPEC_TYPE
        {
            get;
            set;
        }

    }

    //----------------Page------------------------   
    private List<MeasSpec> getDataSource() 
    {

        List<MeasSpec> specObj = new List<MeasSpec>();
        SqlConnection conn = null;
        SqlDataAdapter sAdpt = null;
        MeasSpec mObj;
        String sqlStr = "select p_id, part_id, meas_item, USL, CL, LSL, SPEC_TYPE ";
        sqlStr += "From CTF_Monitor_Performance_SPEC ";
        sqlStr += "where part_id='" + (this.lab_part.Text) + "'";
        DataTable dt;

        try
        {

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();
            sAdpt = new SqlDataAdapter(sqlStr, conn);
            dt = new DataTable();
            sAdpt.Fill(dt);
            conn.Close();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                mObj = new MeasSpec();
                mObj.Id = Convert.ToInt32(dt.Rows[i]["p_id"]);
                mObj.part_id = (dt.Rows[i]["part_id"]).ToString();
                mObj.meas_item = (dt.Rows[i]["meas_item"]).ToString();
                mObj.USL = Convert.ToDouble(dt.Rows[i]["USL"]);
                mObj.CL = Convert.ToDouble(dt.Rows[i]["CL"]);
                mObj.LSL = Convert.ToDouble(dt.Rows[i]["LSL"]);
                mObj.SPEC_TYPE = (dt.Rows[i]["SPEC_TYPE"]).ToString();
                specObj.Add(mObj);
            }

        }
        catch (Exception error) { }
        finally 
        { 
          if (conn.State == ConnectionState.Open) {
              conn.Close();
          }
        }

        return specObj;  
    }

    private static int curId = 10000;
    
    private static object lockObj = new object();

    private int NewId
    {
        get
        {
            return System.Threading.Interlocked.Increment(ref curId);
        }
    }

    private List<MeasSpec> CurrentData
    {
        get
        {
            object persons = this.Session["TestPersons"];
            if (persons == null)
            {
                persons = getDataSource();
                this.Session["TestPersons"] = persons;
            }
            return (List<MeasSpec>)persons;
        }
    }

    private int? AddPerson(MeasSpec person)
    {
        lock (lockObj)
        {
            List<MeasSpec> persons = this.CurrentData;
            person.Id = this.NewId;
            persons.Add(person);
            this.Session["TestPersons"] = persons;
            return person.Id;
        }
    }

    private void DeletePerson(int id)
    {
        lock (lockObj)
        {
            List<MeasSpec> persons = this.CurrentData;
            MeasSpec person = null;
            foreach (MeasSpec p in persons)
            {
                if (p.Id == id)
                {
                    person = p;
                    break;
                }
            }
            if (person == null)
            {
                throw new Exception("TestPerson not found");
            }
            persons.Remove(person);
            this.Session["TestPersons"] = persons;
        }
    }

    private void UpdatePerson(MeasSpec person)
    {
        lock (lockObj)
        {
            List<MeasSpec> persons = this.CurrentData;
            MeasSpec updatingPerson = null;

            foreach (MeasSpec p in persons)
            {
                if (p.Id == person.Id)
                {
                    updatingPerson = p;
                    break;
                }
            }

            if (updatingPerson == null)
            {
                throw new Exception("TestPerson not found");
            }

            updatingPerson.meas_item = person.meas_item;
            updatingPerson.USL = person.USL;
            updatingPerson.CL = person.CL;
            updatingPerson.LSL = person.LSL;

            this.Session["TestPersons"] = persons;
        }
    }

    private void BindData()
    {
        if (X.IsAjaxRequest)
        {
            return;
        }

        GridPanel1.Title = this.lab_part.Text;
        this.Store1.DataSource = this.CurrentData;
        this.Store1.DataBind();
    }

    protected void SaveClick(object sender, DirectEventArgs e)
    {

        SqlConnection conn = null;
        ChangeRecords<MeasSpec> persons = new StoreDataHandler(e.ExtraParams["data"]).BatchObjectData<MeasSpec>();

        try
        {
            
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();

            foreach (MeasSpec created in persons.Created)
            {
                this.AddPerson(created);
                CTF_Spec(ref conn, created);
            }

            foreach (MeasSpec updated in persons.Updated)
            {
                this.UpdatePerson(updated);
                ModelProxy record = Store1.GetById(updated.Id.Value);
                record.Commit();
                CTF_Spec(ref conn, updated);
            }

            foreach (MeasSpec deleted in persons.Deleted)
            {
                this.DeletePerson(deleted.Id.Value);
                Store1.CommitRemoving(deleted.Id.Value);
                DeleteSPEC(ref conn, deleted);
            }

            conn.Close();
            X.Js.Call("backOpener");
        }
        catch (Exception error) 
        { 
            //exeMessage(error.Message); 
        }
        finally 
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
            this.Store1.DataSource = this.CurrentData;
            this.Store1.DataBind();
        }
    }

    #region "ReCaculate 重算統計值"
    /**
     * Step1. 更新表格 CTF_Monitor_Performance_SPEC
     * Step2. 取得資料 CTF_Monitor_Performance_RawData 算統計
     * Step3. 更新表格 CTF_Monitor_Performance_Lot_Summary 
     **/
    private void CTF_Spec(ref SqlConnection conn, MeasSpec measObj)
    {
        List<MeasSpec> specObj = new List<MeasSpec>();
        SqlCommand comm = null;

        if (measObj.part_id == "") 
        {
            measObj.part_id = GridPanel1.Title.Trim();
            measObj.part_id = this.lab_part.Text;
        }

        String part_id = measObj.part_id;
        String measItem = measObj.meas_item;
        String cl = measObj.CL.ToString();
        String usl = measObj.USL.ToString();
        String lsl = measObj.LSL.ToString();
        String sideType = measObj.SPEC_TYPE;
        String sqlStr = "";
        String updateStr = "";
        String insertStr = "";

        if (sideType.Equals("T")) { //雙邊
            updateStr = "update set SPEC_TYPE='{0}', USL={1}, LSL={2}, CL={3} ";
            insertStr = "insert (Part_id, Meas_item, SPEC_TYPE, USL, LSL, CL) values('{0}','{1}','{2}', {3}, {4}, {5});";
            updateStr = String.Format(updateStr, "T", usl, lsl, cl);
            insertStr = String.Format(insertStr, part_id, measItem, "T", usl, lsl, cl);
        } else if (sideType.Equals("U")) { // 單邊[上]
            updateStr = "update set SPEC_TYPE='{0}', USL={1} ";
            insertStr = "insert (Part_id, Meas_item, SPEC_TYPE, USL) values('{0}','{1}','{2}', {3});";
            updateStr = String.Format(updateStr, "U", usl);
            insertStr = String.Format(insertStr, part_id, measItem, "U", usl);
        }  else if (sideType.Equals("L")) { // 單邊[下]
            updateStr = "update set SPEC_TYPE='{0}', LSL={1} ";
            insertStr = "insert (Part_id, Meas_item, SPEC_TYPE, LSL) values('{0}','{1}','{2}', {3});";
            updateStr = String.Format(updateStr, "L", lsl);
            insertStr = String.Format(insertStr, part_id, measItem, "L", lsl);
        }

        sqlStr += "MERGE INTO CTF_Monitor_Performance_SPEC as t_fv ";
        sqlStr += "USING (select '{0}' PART_ID, '{1}' MEAS_ITEM) as s_fv ";
        sqlStr += "ON t_fv.PART_ID = s_fv.PART_ID and t_fv.MEAS_ITEM = s_fv.MEAS_ITEM ";
        sqlStr += "WHEN MATCHED THEN ";
        sqlStr += updateStr;
        sqlStr += "WHEN NOT MATCHED THEN ";
        sqlStr += insertStr;
        sqlStr = String.Format(sqlStr, part_id, measItem);

        try 
        {
            comm = conn.CreateCommand();
            comm.CommandText = sqlStr;
            comm.ExecuteNonQuery();
            CalStatistics(ref conn, measObj);
        }
        catch (Exception Error) { throw Error; }
    }

    private void CalStatistics(ref SqlConnection conn, MeasSpec measObj)
    {
        SqlDataAdapter sAdpt = null;
        DataTable dtRaw;
        String sqlStr = "";
        String part_id = measObj.part_id;
        String measItem = measObj.meas_item;

        try
        {
            sqlStr += "select Lot_Id, Machine_Id, Data_Value ";
            sqlStr += "from dbo.CTF_Monitor_Performance_RawData ";
            sqlStr += "where 1 = 1 ";
            sqlStr += "and Part_Id='{0}' ";
            sqlStr += "and Meas_Item='{1}' ";
            sqlStr += "group by Lot_Id, Machine_Id, Data_Value ";
            sqlStr += "order by Lot_Id, Machine_Id ";
            sqlStr = String.Format(sqlStr, part_id, measItem);

            sAdpt = new SqlDataAdapter(sqlStr, conn);
            dtRaw = new DataTable();
            sAdpt.Fill(dtRaw);

            if (dtRaw.Rows.Count > 0)
            {
                Decimal[] arrSeries = null;
                bool noLineSpc = false;
                String FMACHINE = "";
                String BMACHINE = "";
                String FLOT = "";
                String BLOT = "";
                ArrayList rawAry = new ArrayList();

                for (int i = 0; i < dtRaw.Rows.Count; i++)
                {
                    if (i == 0)
                    {
                        FLOT = (String)dtRaw.Rows[i]["LOT_ID"];
                        FMACHINE = (String)dtRaw.Rows[i]["Machine_Id"];
                    }
                    BLOT = (String)dtRaw.Rows[i]["LOT_ID"];
                    BMACHINE = (String)dtRaw.Rows[i]["Machine_Id"];

                    if (FLOT != BLOT)
                    {
                        //計算統計值
                        Array.Resize<Decimal>(ref arrSeries, (rawAry.Count));
                        for (int j = 0; j < rawAry.Count; j++)
                        {
                            arrSeries[j] = Convert.ToDecimal(rawAry[j]);
                        }
                        InsertStatistics(ref conn, ref arrSeries, FLOT, FMACHINE, measObj);
                        FLOT = BLOT;
                        FMACHINE = BMACHINE;
                        rawAry = new ArrayList();
                    }

                    if ((dtRaw.Rows[i]["Data_Value"]) != null)
                    {
                        rawAry.Add(dtRaw.Rows[i]["Data_Value"]);
                    }
                }

                // 計算統計值 - 最後一筆
                Array.Resize<Decimal>(ref arrSeries, (rawAry.Count));
                for (int j = 0; j < rawAry.Count; j++)
                {
                    arrSeries[j] = Convert.ToDecimal(rawAry[j]);
                }
                InsertStatistics(ref conn, ref arrSeries, BLOT, BMACHINE, measObj);
            }

        }
        catch (Exception Error) { throw Error; }
    }

    private void InsertStatistics(ref SqlConnection conn, ref Decimal[] decimalAry, String lot_id, String machine_id, MeasSpec measObj)
    {
        SPCFormula objFormula = null;
        SPCStatisc objStatisc = null;
        SqlCommand comm;
        String sqlStr  = "";
        Decimal usl ;
        Decimal lsl ;
        String decimalAryStr = "";
        String partID = measObj.part_id;

        for (int i = 0 ; i < (decimalAry.Length); i++) 
        {
          decimalAryStr += Math.Round(decimalAry[i], 5).ToString() + "|";
        }
        decimalAryStr = decimalAryStr.Substring(0, (decimalAryStr.Length - 1));

        if ((measObj.SPEC_TYPE).Equals("T"))
        { //雙邊
            usl = Convert.ToDecimal(measObj.USL);
            lsl = Convert.ToDecimal(measObj.LSL);
            objFormula = new SPCFormula(decimalAry, usl, lsl, null);
            objStatisc = objFormula.Calculate();
        }
        else if ((measObj.SPEC_TYPE).Equals("U"))
        { // 單邊[上]
            usl = Convert.ToDecimal(measObj.USL);
            objFormula = new SPCFormula(decimalAry, usl, null, null);
            objStatisc = objFormula.Calculate();
        }
        else if ((measObj.SPEC_TYPE).Equals("L"))
        { // 單邊[下]
            lsl = Convert.ToDecimal(measObj.LSL);
            objFormula = new SPCFormula(decimalAry, null, lsl, null);
            objStatisc = objFormula.Calculate();
        }

        sqlStr = "";
        sqlStr += "update CTF_Monitor_Performance_Lot_Summary ";
        sqlStr += "set Mean_Value={0}, Std_Value={1}, Min_Value={2}, Max_Value={3}, CP={4}, CPK={5}, USL={6}, LSL={7}, CSL={8}, rowdata='{9}' ";
        sqlStr += "where 1=1 ";
        sqlStr += "and part_id='{10}' ";
        sqlStr += "and meas_item='{11}' ";
        sqlStr += "and lot_id='{12}' ";
        sqlStr += "and machine_id='{13}' ";

        sqlStr = String.Format(sqlStr,
                               Math.Round(objStatisc.xMean, 5).ToString(),
                               Math.Round(objStatisc.Sigma, 5).ToString(),
                               Math.Round(objStatisc.Minimum, 5).ToString(),
                               Math.Round(objStatisc.Maximum, 5).ToString(),
                               (objStatisc.Cp).ToString(),
                               (objStatisc.Cpk).ToString(),
                               (measObj.USL).ToString(),
                               (measObj.LSL).ToString(),
                               (objStatisc.Target).ToString(),
                               decimalAryStr,
                               partID,
                               (measObj.meas_item),
                               lot_id,
                               machine_id);
        try 
        { 
            comm = conn.CreateCommand();
            comm.CommandText = sqlStr;
            comm.ExecuteNonQuery();
        } catch(Exception error) {
            throw error;
        }

    }

    private void DeleteSPEC(ref SqlConnection conn, MeasSpec measObj) 
    {
        
        SqlCommand comm;
        String sqlStr = "Delete from CTF_Monitor_Performance_SPEC where 1=1 pard_id='{0}' and Meas_item='{1}' ";
        sqlStr = String.Format(sqlStr, (measObj.part_id), (measObj.meas_item));

        try
        {
            comm = conn.CreateCommand();
            comm.CommandText = sqlStr;
            comm.ExecuteNonQuery();
        }
        catch (Exception error)
        {
            throw error;
        }

    }

    #endregion

    // 導回原頁
    private void exeScript() 
    { 
      StringBuilder sb = new StringBuilder();
      sb.Append("<script language='javascript'>");
      sb.Append("alert(window.opener.document.getElementById('but_Execute'));");
      //sb.Append("window.opener.document.getElementById('but_Execute').click();");
      //sb.Append("this.close();");
      sb.Append("</script>");
      ClientScriptManager myCSManager = this.ClientScript;
      myCSManager.RegisterStartupScript(this.GetType(), "SetStatusScript", sb.ToString());
    }

    // Message
    private void exeMessage(String msg)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("<script language='javascript'>");
        sb.Append("alert('" + msg + "');");
        sb.Append("</script>");
        ClientScriptManager myCSManager = this.ClientScript;
        myCSManager.RegisterStartupScript(this.GetType(), "SetStatusScript", sb.ToString());
    }

    //
    protected void TransferAll_Click(object sender, EventArgs e)
    {
        SqlConnection conn = null;
        SqlDataAdapter sqlAdp;
        String sqlStr = "";
        DataTable specDT;
        DataTable rowDT;
        MeasSpec measObj;
        
        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();
            sqlStr = "";
            sqlStr += "select P_ID, Part_Id, Meas_Item, USL, CL, LSL, SPEC_TYPE ";
            sqlStr += "from dbo.CTF_Monitor_Performance_SPEC ";
            sqlAdp = new SqlDataAdapter(sqlStr, conn);
            specDT = new DataTable();
            sqlAdp.Fill(specDT);

            for (int x = 0; x < specDT.Rows.Count; x++)
            {
                measObj = new MeasSpec();
                measObj.Id = (int)specDT.Rows[x]["P_ID"];
                measObj.part_id = ((String)specDT.Rows[x]["Part_Id"]).Trim();
                measObj.meas_item = ((String)specDT.Rows[x]["Meas_Item"]).Trim();
                measObj.USL = (double)specDT.Rows[x]["USL"];
                measObj.CL = (double)specDT.Rows[x]["CL"];
                measObj.LSL = (double)specDT.Rows[x]["LSL"];
                measObj.SPEC_TYPE = ((String)specDT.Rows[x]["SPEC_TYPE"]).Trim();

                // --- SPEC START ---
                sqlStr = "";
                sqlStr += "select Lot_Id, Machine_Id, Data_Value ";
                sqlStr += "from dbo.CTF_Monitor_Performance_RawData ";
                sqlStr += "where 1 = 1 ";
                sqlStr += "and Part_Id='{0}' ";
                sqlStr += "and Meas_Item='{1}' ";
                sqlStr += "group by Lot_Id, Machine_Id, Data_Value ";
                sqlStr += "order by Lot_Id, Machine_Id ";
                sqlStr = String.Format(sqlStr, ((String)specDT.Rows[x]["Part_Id"]).Trim(), ((String)specDT.Rows[x]["Meas_Item"]).Trim());

                sqlAdp = new SqlDataAdapter(sqlStr, conn);
                rowDT = new DataTable();
                sqlAdp.Fill(rowDT);

                if (rowDT.Rows.Count > 0)
                {
                    Decimal[] arrSeries = null;
                    bool noLineSpc = false;
                    String FMACHINE = "";
                    String BMACHINE = "";
                    String FLOT = "";
                    String BLOT = "";
                    ArrayList rawAry = new ArrayList();

                    for (int i = 0; i < rowDT.Rows.Count; i++)
                    {
                        if (i == 0)
                        {
                            FLOT = (String)rowDT.Rows[i]["LOT_ID"];
                            FMACHINE = (String)rowDT.Rows[i]["Machine_Id"];
                        }
                        BLOT = (String)rowDT.Rows[i]["LOT_ID"];
                        BMACHINE = (String)rowDT.Rows[i]["Machine_Id"];

                        if (FLOT != BLOT)
                        {
                            //計算統計值
                            Array.Resize<Decimal>(ref arrSeries, (rawAry.Count));
                            for (int j = 0; j < rawAry.Count; j++)
                            {
                                arrSeries[j] = Convert.ToDecimal(rawAry[j]);
                            }
                            InsertStatistics(ref conn, ref arrSeries, FLOT, FMACHINE, measObj);
                            FLOT = BLOT;
                            FMACHINE = BMACHINE;
                            rawAry = new ArrayList();
                        }

                        if ((rowDT.Rows[i]["Data_Value"]) != null)
                        {
                            rawAry.Add(rowDT.Rows[i]["Data_Value"]);
                        }
                    }

                    // 計算統計值 - 最後一筆
                    Array.Resize<Decimal>(ref arrSeries, (rawAry.Count));
                    for (int j = 0; j < rawAry.Count; j++)
                    {
                        arrSeries[j] = Convert.ToDecimal(rawAry[j]);
                    }
                    InsertStatistics(ref conn, ref arrSeries, BLOT, BMACHINE, measObj);
                }

            }
            conn.Close();
        }
        catch (Exception error)
        {
            exeMessage(error.Message); 
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }
}