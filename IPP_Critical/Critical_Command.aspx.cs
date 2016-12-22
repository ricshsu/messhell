using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Ext.Net;
using System.Data;
using System.Collections;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;

public partial class Critical_Command : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!X.IsAjaxRequest)
        {
            this.Session["TestPersons"] = null;
            this.UserForm.Title = this.lab_user.Text;
        }

        if (!this.IsPostBack)
        {
            //this.lab_part.Text = "FCS077";
            //this.lab_lot.Text = "N2BA90510131";
            //this.lab_item.Text = "Ni Thickness";
            //this.lab_user.Text = "N000142679";
            //pageInit();

            if (Request["partID"] != null)
            {
                this.lab_part.Text = Request["partID"].ToString();
                this.lab_lot.Text = Request["lotID"].ToString();
                this.lab_item.Text = Request["item"].ToString();
                this.lab_user.Text = Request["userID"].ToString();
                this.lab_fType.Text = Request["fType"].ToString();
                pageInit();
                this.BindData();
            }
        }
        else 
        {
            this.BindData();
        }
    }

    private void pageInit() 
    {
        SqlConnection conn = null;
        SqlDataAdapter sAdpt = null;
        String sqlStr = "";
        DataTable dt;

        if ((this.lab_fType.Text.Trim()).Equals("Critical_Lot"))
        {
            sqlStr += "select 1 as p_id, part as part_id, lot, Parametric_Measurement as meas_item,  ";
            sqlStr += "plant, mchno as tool, Convert(char(19), trtm, 120) as trtm ";
            //sqlStr += "'N000142679' as userid, Convert(char(19), trtm, 120) as updatetime, 'Handle' as comment ";
            sqlStr += "from dbo.Critical_LOT_Data ";
            sqlStr += "where 1=1 and MeasureCount = 1 ";
            sqlStr += "and part='" + (this.lab_part.Text) + "' ";
            sqlStr += "and Lot='" + (this.lab_lot.Text) + "' ";
            sqlStr += "and Parametric_Measurement='" + (this.lab_item.Text) + "' ";
        }
        else if ((this.lab_fType.Text.Trim()).Equals("Critical_KPP"))
        {
            sqlStr += "select 1 as p_id, part_id, lot, Parametric_Measurement as meas_item,  ";
            sqlStr += "plant, EQPID as tool, Convert(char(19), trtm, 120) as trtm ";
            sqlStr += "from view_IPP_Process_CriticalItem_Monitor ";
            sqlStr += "where 1=1 ";
            sqlStr += "and part_id='" + (this.lab_part.Text) + "' ";
            sqlStr += "and Lot='" + (this.lab_lot.Text) + "' ";
            sqlStr += "and Parametric_Measurement='" + (this.lab_item.Text) + "' ";
        } 
        else
        {
            sqlStr += "select 1 as p_id, part_id, lot, Parametric_Measurement as meas_item,  ";
            sqlStr += "plant, mchno as tool, Convert(char(19), trtm, 120) as trtm ";
            sqlStr += "from view_IPP_CriticalItem_Monitor ";
            sqlStr += "where 1=1 ";
            sqlStr += "and part_id='" + (this.lab_part.Text) + "' ";
            sqlStr += "and Lot='" + (this.lab_lot.Text) + "' ";
            sqlStr += "and Parametric_Measurement='" + (this.lab_item.Text) + "' ";
        }
        
        
        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();
            sAdpt = new SqlDataAdapter(sqlStr, conn);
            dt = new DataTable();
            sAdpt.Fill(dt);
            conn.Close();

            for (int i = 0; i < 1; i++)
            {
                this.tf_part.Value = (dt.Rows[i]["part_id"]).ToString();
                this.tf_lot.Text = (dt.Rows[i]["lot"]).ToString();
                this.tf_item.Value = (dt.Rows[i]["meas_item"]).ToString();
                this.tf_plant.Value = (dt.Rows[i]["plant"]).ToString();
                this.tf_machine.Value = (dt.Rows[i]["tool"]).ToString();
                this.tf_trtm.Value = (dt.Rows[i]["trtm"]).ToString();
                this.tf_user.Value = this.lab_user.Text;
                this.tf_uTime.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            }

        }
        catch (Exception error) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    
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

        public String lot
        {
            get;
            set;
        }

        public String meas_item
        {
            get;
            set;
        }

        public String plant
        {
            get;
            set;
        }

        public String tool
        {
            get;
            set;
        }

        public String trtm
        {
            get;
            set;
        }

        public String comment
        {
            get;
            set;
        }

        public String user
        {
            get;
            set;
        }

        public String updatetime
        {
            get;
            set;
        }

        public String mean
        {
            get;
            set;
        }

        public String std
        {
            get;
            set;
        }

        public String max
        {
            get;
            set;
        }

        public String min
        {
            get;
            set;
        }

        public String cp
        {
            get;
            set;
        }

        public String cpk
        {
            get;
            set;
        }

    }

    //----------------Page------------------------   
    private List<MeasSpec> getDataSource()
    {
        List<MeasSpec> specObj = new List<MeasSpec>();
        DataTable dt;
        SqlConnection conn = null;
        SqlDataAdapter sAdpt = null;
        MeasSpec mObj;
        String sqlStr = "";

        if ((this.lab_fType.Text.Trim()).Equals("Critical_Lot"))
        {
            sqlStr += "select CASE WHEN b.p_id is null then '1' else b.p_id end as p_id, ";
            sqlStr += "a.part as part_id, a.lot, a.Parametric_Measurement as meas_item, ";
            sqlStr += "a.plant, a.mchno as tool, Convert(char(19), a.trtm, 120) as trtm, ";
            sqlStr += "a.meanval, a.std, a.maxval, a.minval, a.cp, a.cpk, ";
            sqlStr += "CASE WHEN b.userid is null then 'N12345678' else b.userid end as userid, ";
            sqlStr += "CASE WHEN b.updatetime is null then '' else Convert(char(19), b.updatetime, 120) end as updatetime,  ";
            sqlStr += "CASE WHEN b.comment is null then '' else b.comment end as comment ";
            sqlStr += "from dbo.Critical_LOT_Data a, dbo.Critical_Lot_Control b ";
            sqlStr += "where 1=1 and a.MeasureCount = 1 ";
            sqlStr += "and a.part = b.part ";
            sqlStr += "and a.Lot = b.lot ";
            sqlStr += "and a.Parametric_Measurement = b.meas_item ";
            sqlStr += "and a.part='" + (this.lab_part.Text) + "' ";
            sqlStr += "and a.Lot='" + (this.lab_lot.Text) + "' ";
            sqlStr += "and a.Parametric_Measurement='" + (this.lab_item.Text) + "' ";
            sqlStr += "order by b.updatetime desc";
        }
        else if ((this.lab_fType.Text.Trim()).Equals("Critical_KPP"))
        {
            sqlStr += "select CASE WHEN b.p_id is null then '1' else b.p_id end as p_id, ";
            sqlStr += "a.part_id, a.lot, a.Parametric_Measurement as meas_item, a.plant,";
            sqlStr += "CASE WHEN a.EqpID is null then '' else a.EqpID end as tool, ";
            sqlStr += "Convert(char(19), a.trtm, 120) as trtm, ";
            sqlStr += "a.meanval, a.std, a.maxval, a.minval, a.cp, a.cpk, ";
            sqlStr += "CASE WHEN b.userid is null then 'N12345678' else b.userid end as userid, ";
            sqlStr += "CASE WHEN b.updatetime is null then '' else Convert(char(19), b.updatetime, 120) end as updatetime,  ";
            sqlStr += "CASE WHEN b.comment is null then '' else b.comment end as comment ";
            sqlStr += "from dbo.view_IPP_Process_CriticalItem_Monitor a, dbo.Critical_Lot_Control b ";
            sqlStr += "where 1=1 ";
            sqlStr += "and a.part_id = b.part ";
            sqlStr += "and a.Lot = b.lot ";
            sqlStr += "and a.Parametric_Measurement = b.meas_item ";
            sqlStr += "and a.part_id='" + (this.lab_part.Text) + "' ";
            sqlStr += "and a.Lot='" + (this.lab_lot.Text) + "' ";
            sqlStr += "and a.Parametric_Measurement='" + (this.lab_item.Text) + "' ";
            sqlStr += "order by b.updatetime desc";
        }
        else 
        {
            sqlStr += "select CASE WHEN b.p_id is null then '1' else b.p_id end as p_id, ";
            sqlStr += "a.part_id, a.lot, a.Parametric_Measurement as meas_item, a.plant,";
            sqlStr += "CASE WHEN a.mchno is null then '' else a.mchno end as tool, ";
            sqlStr += "Convert(char(19), a.trtm, 120) as trtm, ";
            sqlStr += "a.meanval, a.std, a.maxval, a.minval, a.cp, a.cpk, ";
            sqlStr += "CASE WHEN b.userid is null then 'N12345678' else b.userid end as userid, ";
            sqlStr += "CASE WHEN b.updatetime is null then '' else Convert(char(19), b.updatetime, 120) end as updatetime,  ";
            sqlStr += "CASE WHEN b.comment is null then '' else b.comment end as comment ";
            sqlStr += "from dbo.view_IPP_CriticalItem_Monitor a, dbo.Critical_Lot_Control b ";
            sqlStr += "where 1=1 ";
            sqlStr += "and a.part_id = b.part ";
            sqlStr += "and a.Lot = b.lot ";
            sqlStr += "and a.Parametric_Measurement = b.meas_item ";
            sqlStr += "and a.part_id='" + (this.lab_part.Text) + "' ";
            sqlStr += "and a.Lot='" + (this.lab_lot.Text) + "' ";
            sqlStr += "and a.Parametric_Measurement='" + (this.lab_item.Text) + "' ";
            sqlStr += "order by b.updatetime desc";
        }
        
        
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
                mObj.lot = (dt.Rows[i]["lot"]).ToString();
                mObj.meas_item = (dt.Rows[i]["meas_item"]).ToString();
                mObj.plant = (dt.Rows[i]["plant"]).ToString();
                mObj.tool = (dt.Rows[i]["tool"]).ToString();
                mObj.trtm = (dt.Rows[i]["trtm"]).ToString();
                mObj.user = (dt.Rows[i]["userid"]).ToString();
                mObj.updatetime = (dt.Rows[i]["updatetime"]).ToString();
                mObj.comment = (dt.Rows[i]["comment"]).ToString();
                mObj.mean = (dt.Rows[i]["meanval"]).ToString();
                mObj.std = (dt.Rows[i]["std"]).ToString();
                mObj.max = (dt.Rows[i]["maxval"]).ToString();
                mObj.min = (dt.Rows[i]["minval"]).ToString();
                mObj.cp = (dt.Rows[i]["cp"]).ToString();
                mObj.cpk = (dt.Rows[i]["cpk"]).ToString();
                specObj.Add(mObj);
            }

        }
        catch (Exception error) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
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
            updatingPerson.part_id = person.part_id;
            updatingPerson.lot = person.lot;
            updatingPerson.comment = person.comment;

            this.Session["TestPersons"] = persons;
        }
    }

    private void BindData()
    {
        if (X.IsAjaxRequest)
        {
            return;
        }

        GridPanel1.Title = "Comment History";
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
                insertControlTable(ref conn, created);
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

    private void insertControlTable(ref SqlConnection conn, MeasSpec measObj) 
    {
        SqlCommand comm = null;
        String sqlStr = "insert into Critical_Lot_Control(part, lot, meas_item, userid, comment, updatetime) ";
        sqlStr += "values('{0}', '{1}', '{2}', '{3}', '{4}', getdate())";
        sqlStr = String.Format(sqlStr, measObj.part_id, measObj.lot, measObj.meas_item, (this.tf_user.Value), measObj.comment);
        
        try
        {
            comm = conn.CreateCommand();
            comm.CommandText = sqlStr;
            comm.ExecuteNonQuery();
        }
        catch (Exception Error) { throw Error; }
    }

}