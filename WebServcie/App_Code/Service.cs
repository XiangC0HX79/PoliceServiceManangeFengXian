using System;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml; 
using System.Xml.Linq;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;

using System.Text;
using System.Data;
using System.Data.OleDb;

using SharedNamespace;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消对下行的注释。
// [System.Web.Script.Services.ScriptService]
public class Service : System.Web.Services.WebService
{
    private String strConn = "Data Source=10.18.168.39;user id=sa;password=1111;initial catalog=fxqw;min pool size=5;persist security info=true";
    private String strMapUrl = "http://10.18.168.38/ArcGISNEW/services/FXGAXM/MapServer";
    private String strSDEPredix = "[sde].[sde].";

    //小区名单，默认小区后缀为 小区/街坊/公寓/村/苑/宅/里，黑名单里面名称为小区，白名单里面名称为单位
    private String[] quarterPredixList = { "小区","街坊","公寓","村","苑","宅","里" };
    private String[] quarterBlackList = { "泰来坊","美墅","中远两湾城","阳光威尼斯" };
    private String[] quarterWhiteList = { "上海纤云美容美发苑", "文沁苑", "黎金苑","悠诗阁酒店公寓","新俪公寓","新华源公寓","华东师范大学中江路学生公寓","华东师范大学留学生公寓","江南春食苑","金银花足浴保健苑","毛毛休闲茶苑","上海康康足部保健苑","上海纤云美容美发苑","西部秀苑","晓英音乐苑" };
    
    public Service () {

        //如果使用设计的组件，请取消注释以下行 
        //InitializeComponent(); 
    }

    [WebMethod]
    public string HelloWorld() {
        return "Hello World";
    }

    /// <summary>
    /// 初始化：系统配置
    /// </summary>
    /// <returns></returns>
    [WebMethod]
    public DataTable getSysConfig()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "PARAID, " +
                        "PARANAME, " +
                        "PARAVALUE " + 
                        "FROM T_QW_SYS_PARA";
        return clsGetData.GetTable(sql);
    }

    /// <summary>
    /// 初始化：系统字典
    /// </summary>
    /// <returns></returns>
    [WebMethod]
    public DataTable getSysDic()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "[DICID] " +
                        ",[PDICID] " +
                        ",[DICVALUE] " +
                        ",[ISUSE] " +
                        ",[NOTE] " +
                        ",[ORDERNUM] " +
                        "FROM T_QW_SYS_DIC";
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getDepartmentInfo()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "DEPID, " +
                        "PX, " +
                        "JC, " +
                        "DWMC, " +
                        "ZB " +
                        " FROM V_DWXX";
        DataTable deptTable = clsGetData.GetTable(sql);

        return deptTable;
    }

    [WebMethod]
    public DataTable getAlarmInfo(String deptName, String type, String beginTime, String endTime)
    {
        String sql = "SELECT " +
                        "JQDH id, " +
                        "T_QW_JQ.JQLB type, " +
                        "BJSJ time, " +
                        "BJDZ address, " +
                        "BJRXM name, " +
                        "JQBT title, " +
                        "BJDH phone, " +
                        "LXDH contactphone, " +
                        "CJR police, " +
                        "JQNR info, " +
                        "SSXQ dept, " +
                        "X x, " +
                        "Y y " +
                          ",NEWX " +
                          ",NEWY " +
                          ",ISMUST " +
                          ",MAINCATEGORY maintype " +
                        "FROM T_QW_JQ LEFT JOIN T_QW_JQLBFZ ON T_QW_JQ.JQLB = T_QW_JQLBFZ.JQLB WHERE ";

        if (deptName != "所有单位")
            sql += "SSXQ = '" + deptName + "' AND ";

        if (type != "所有类别")
            sql += "MAINCATEGORY = '" + type + "' AND ";

        if (beginTime == "")
            sql += "BJSJ > DATEADD(HH,-2,GETDATE()) ";
        else
            sql += "CONVERT(VARCHAR,BJSJ,120) > '" + beginTime + "' ";

        if (endTime == "")
            sql += "AND BJSJ <= GETDATE() ";
        else
            sql += "AND CONVERT(VARCHAR,BJSJ,120) <= '" + endTime + "' ";

        sql += "ORDER BY BJSJ";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        DataTable dataTable = clsGetData.GetTable(sql);

        return dataTable;
    }


    [WebMethod]
    public DataTable getAlarmTypeInfo()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT DISTINCT " +
                        "[DESC] AS ID, " +
                        "MAINCATEGORY AS JQNAME " +
                        " FROM T_QW_JQLBFZ";
        DataTable deptTable = clsGetData.GetTable(sql);

        return deptTable;
    }
    
    [WebMethod]
    public DataTable getAlarmPolice(String alarmID)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "T_JQ_Process.ID , " +
                        "JQID , " +
                        "JYID , " +
                        "TypeDateTime , " +
                        "pType " +
                        "FROM T_JQ_Process ,( " +
                        "SELECT MAX(ID) ID " +
                        "FROM T_JQ_Process " +
                        "GROUP BY JQID,JYID) MAXID " +
                        "WHERE T_JQ_Process.ID = MAXID.ID " +
                        "AND T_JQ_Process.JQID = '" + alarmID + "'";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getCallInfo(String policeID,String stid)
    {
        DateTime now = DateTime.Now; 
        String beginTime = (DateTime.Now.AddDays(-1)).ToString("yyyy-MM-dd HH:mm:ss");

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "ID , " +
                        "DMDate , " +
                        "IsResponse , " +
                        "ResponseDate , " +
                        "BDMNAME , " +
                        "BDMDepName , " +
                        "BDMJYID , " +
                        "BDMDEPID , " +
                        "DMNAME , " +
                        "DMJYID , " +
                        "DMDEPNAME , " +
                        "MEMO , " +
                        "ISCONFIRM " +
                        "FROM T_QW_DM WHERE BDMJYID = '" + policeID + "' AND STID = '" + stid + "' " +
                        "AND CONVERT(VARCHAR,DMDate,120) > '" + beginTime + "' " +
                        "ORDER BY DMDate DESC";
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getServiceDic()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "QWTYPEID, " +
                        "QWTYPENAME, " +
                        "ISGISSHOW, " +
                        "ImageName " +
                        "FROM T_QW_SYS_QWTYPE WHERE ISUSE = 1 ORDER BY QWTYPEID";
        return clsGetData.GetTable(sql);
    }
    
    [WebMethod]
    public DataTable getGPSRealTimeInfo(String GPSID)
    {
        String sql = "SELECT " +
                    "''         GPSID, " +
                    "''    GPSSIMCARD, " +
                    "''          TYPE, " +
                    "''       GPSTYPE, " +
                    "''       GPSNAME, " +
                    "''      GPSDEPID, " +
                    "''    GPSDEPNAME, " +
                    "''      POLICENO, " +
                    "''         PHONE, " +
                    "''     LONGITUDE, " +
                    "''      LATITUDE, " +
                    "GETDATE()      DATARECORDTIME, " +
                    "''          QWID, " +
                    "''      QWTYPE, " +
                    "''  QWSTATUS, " +
                    "''       QWSTATUSNAME, " +
                    "''     GPSSTATUS, " +
                    "''         CARNO, " +
                    "''        USERID, " +
                    "''          RADIONO, " +
                    "''            CALLNO, " +
                    "''        ZONENM, " +
                    "''    PATROLTYPE, " +
                    "''       RUNNAME, " +
                    "''     STARTTIME, " +
                    "''       ENDTIME, " +
                    "''      DUTYNOTE, " +
                    "''            SEX, " +
                    "''     STATECHANGETIME ";

        sql += "UNION SELECT " +
                     "GPSID         GPSID, " +
                     "GPSSIMCARD    GPSSIMCARD, " +
                     "TYPE          TYPE, " +
                     "GPSTYPE       GPSTYPE, " +
                     "GPSNAME       GPSNAME, " +
                     "GPSDEPID      GPSDEPID, " +
                     "GPSDEPNAME    GPSDEPNAME, " +
                     "POLICENO      POLICENO, " +
                     "PHONE         PHONE, " +
                     "LONGITUDE     LONGITUDE, " +
                     "LATITUDE      LATITUDE, " +
                     "DATARECORDTIME      DATARECORDTIME, " +
                     "QWID          QWID, " +
                     "DUTYTYPE      QWTYPE, " +
                     "CURRENTSTATE  QWSTATUS, " +
                     "STATENM       QWSTATUSNAME, " +
                     "GPSSTATUS     GPSSTATUS, " +
                     "CARNO         CARNO, " +
                     "USERID        USERID, " +
                     "STID          RADIONO, " +
                     "HH            CALLNO, " +
                     "ZONENM        ZONENM, " +
                     "PATROLTYPE    PATROLTYPE, " +
                     "RUNNAME       RUNNAME, " +
                     "STARTTIME     STARTTIME, " +
                     "ENDTIME       ENDTIME, " +
                     "DUTYNOTE      DUTYNOTE, " +
                     "XB            SEX, " +
                     "STATECHANGETIME     STATECHANGETIME ";
        if (GPSID != "")
            sql += "FROM VIEW_GPSPLAN WHERE GPSSIMCARD is not null and GPSSIMCARD!='' AND GPSID > " + GPSID;
        else
            sql += "FROM VIEW_GPSPLAN WHERE GPSSIMCARD is not null and GPSSIMCARD!=''";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getServiceStatic()
    {
        String sql = "SELECT " +
                     "[QWID] " +
      ",[STARTTIME] " +
      ",[ENDTIME] " +
      ",[DUTYTYPE]      QWTYPE" +
      ",[USERID] " +
      ",[POLICEUSER]    GPSNAME" +
      ",[DEPID]         GPSDEPID" +
      ",[DEPNAME]       GPSDEPNAME" +
      ",[CURRENTSTATE]  QWSTATUS" +
      ",[STATECHANGETIME] STATECHANGETIME" +
      ",[GPSSTATUS] " +
      ",[GPSSTATUSCHANGETIME] DATARECORDTIME" +
      ",[STATENM]       QWSTATUSNAME" +
      ",[JH]            POLICENO" +
      ",[STID]          RADIONO" +
      ",RUNNAME       RUNNAME " +
      ",NOTE      DUTYNOTE " +
      "  FROM View_MaxPlanDuty";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getGPSListTrackHis(String deptID, String name, String beginTime, String endTime)
    {
        //定义存放的数据表
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                     "DISTINCT GPSSIMCARD " +
                     "FROM T_QW_SYS_GPSDATA LEFT JOIN V_DWXX ON GPSDEPID = V_DWXX.DEPID WHERE  GPSSIMCARD is not null and GPSSIMCARD!='' ";

        if (deptID == "-1")
        {
        }
        else if (deptID == "-2")
        {
            sql += "AND V_DWXX.ZB = 125 ";
        }
        else
        {
            sql += "AND GPSDEPID = '" + deptID + "' ";
        }

        if (name != "")
            sql += "AND ("
                + "(GPSTYPE = 1 AND (GPSNAME LIKE '%" + name + "%' OR GPSSIMCARD LIKE '%" + name + "%'))"
                + "OR (GPSTYPE <> 1 AND (GPSNAME LIKE '%" + name + "%' OR POLICENO LIKE '%" + name + "%' OR GPSSIMCARD LIKE '%" + name + "%'))"
                + ") ";

        sql += "AND CONVERT(VARCHAR,DATARECORDTIME,120) > '" + beginTime + "' " +
                     "AND CONVERT(VARCHAR,DATARECORDTIME,120) <= '" + endTime + "'";
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getGPSTrackLineBack(String deptID, String beginTime, String endTime)
    {
        //定义存放的数据表
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                     "GPSID         GPSID, " +
                     "GPSSIMCARD    GPSSIMCARD, " +
                     "GPSTYPE       GPSTYPE, " +
                     "GPSNAME       GPSNAME, " +
                     "GPSDEPID      GPSDEPID, " +
                     "GPSDEPNAME    GPSDEPNAME, " +
                     "POLICENO      POLICENO, " +
                     "PHONE         PHONE, " +
                     "LONGITUDE     LONGITUDE, " +
                     "LATITUDE      LATITUDE, " +
                     "DATARECORDTIME      DATARECORDTIME, " +
                     "DUTYTYPE      QWTYPE, " +
                     "CURRENTSTATE  QWSTATUS, " +
                     "STATENM       QWSTATUSNAME, " +
                     "PATROLTYPE    PATROLTYPE, " +
                     "USERID        USERID, " +
                     "STID          STID, " +
                     "RUNNAME       RUNNAME, " +
                     "STARTTIME     STARTTIME, " +
                     "ENDTIME       ENDTIME " +
                    "FROM T_QW_SYS_GPSDATA LEFT JOIN " +
            //过滤值班DUTYTYPE = 1
                    "(SELECT STARTTIME,ENDTIME,JH,CURRENTSTATE,STATENM,MAX(PATROLTYPE) PATROLTYPE,USERID,STID,RUNNAME,MAX(DUTYTYPE) DUTYTYPE " +
                    "FROM View_PLANDUTYSTATE " +
                    "GROUP BY STARTTIME,ENDTIME,JH,CURRENTSTATE,STATENM,USERID,STID,RUNNAME) PLANDUTYSTATE ON "
            //过滤值班DUTYTYPE = 1
                    + " PLANDUTYSTATE.JH = T_QW_SYS_GPSDATA.POLICENO "
                    + "AND PLANDUTYSTATE.STARTTIME <= T_QW_SYS_GPSDATA.DATARECORDTIME "
                    + "AND PLANDUTYSTATE.ENDTIME >= T_QW_SYS_GPSDATA.DATARECORDTIME "
                    + "LEFT JOIN V_DWXX ON GPSDEPID = V_DWXX.DEPID ";
        if (deptID == "-2")
            sql += "WHERE V_DWXX.ZB = 125 ";
        else
            sql += "WHERE GPSDEPID = '" + deptID + "' ";

        sql += "AND CONVERT(VARCHAR,DATARECORDTIME,120) > '" + beginTime + "' "
                    + "AND CONVERT(VARCHAR,DATARECORDTIME,120) <= '" + endTime + "' "
                    + "ORDER BY DATARECORDTIME";
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getGPSTrackHis(String gpssimcard, String beginTime,String endTime)
    {                
        //定义存放的数据表
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                     "GPSID         GPSID, " +
                     "GPSSIMCARD    GPSSIMCARD, " +
                     "GPSTYPE       GPSTYPE, " +
                     "GPSNAME       GPSNAME, " +
                     "GPSDEPID      GPSDEPID, " +
                     "GPSDEPNAME    GPSDEPNAME, " +
                     "POLICENO      POLICENO, " +
                     "PHONE         PHONE, " +
                     "LONGITUDE     LONGITUDE, " +
                     "LATITUDE      LATITUDE, " +
                     "DATARECORDTIME      DATARECORDTIME, " +
                     "DUTYTYPE      QWTYPE, " +
                     "CURRENTSTATE  QWSTATUS, " +
                     "STATENM       QWSTATUSNAME, " +
                     "PATROLTYPE    PATROLTYPE, " +
                     "USERID        USERID, " +
                     "STID          STID, " +
                     "RUNNAME       RUNNAME, " +
                     "STARTTIME     STARTTIME, " +
                     "ENDTIME       ENDTIME " +
                        "FROM T_QW_SYS_GPSDATA LEFT JOIN " + 
                        //过滤值班DUTYTYPE = 1
                        "(SELECT STARTTIME,ENDTIME,JH,CURRENTSTATE,STATENM,MAX(PATROLTYPE) PATROLTYPE,USERID,STID,RUNNAME,MAX(DUTYTYPE) DUTYTYPE " +
                        "FROM View_PLANDUTYSTATE " +
                        "GROUP BY STARTTIME,ENDTIME,JH,CURRENTSTATE,STATENM,USERID,STID,RUNNAME) PLANDUTYSTATE ON "
                        //过滤值班DUTYTYPE = 1
                        + " PLANDUTYSTATE.JH = T_QW_SYS_GPSDATA.POLICENO "
                        + "AND PLANDUTYSTATE.STARTTIME <= T_QW_SYS_GPSDATA.DATARECORDTIME "
                        + "AND PLANDUTYSTATE.ENDTIME >= T_QW_SYS_GPSDATA.DATARECORDTIME "
                        + "WHERE GPSSIMCARD = '" + gpssimcard + "' "
                        + "AND CONVERT(VARCHAR,DATARECORDTIME,120) > '" + beginTime + "' "
                        + "AND CONVERT(VARCHAR,DATARECORDTIME,120) <= '" + endTime + "' "
                        + "ORDER BY DATARECORDTIME";
        return clsGetData.GetTable(sql);
    }
    
    [WebMethod]
    public DataTable getPatrolZone()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "KEYID              KEYID, " +
                        "DEPID              DEPID, " +
                        "ZONENM             ZONENM, " +
                        "RANGE              RANGE, " +
                        "LASTUPDATEUSER     LASTUPDATEUSER, " +
                        "LASTUPDATETIME     LASTUPDATETIME, " +
                        "ZONENEAME          ZONENEAME, " +
                        "ZONEGPSRANGE       ZONEGPSRANGE, " +
                        "LASTUPDATETIME     LASTUPDATETIME, " +
                        "ZONENEAME          ZONENEAME " +
                        "FROM T_QW_PATROLZONE";
        return clsGetData.GetTable(sql);
    }


    [WebMethod]
    public DataTable getPatrolPoint()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "ID              ID, " +
                        "MUSTNAME        MUSTNAME, " +
                        "ZONEID          ZONEID, " +
                        "Address         Address, " +
                        "X               X, " +
                        "Y               Y, " +
                        "Type            Type, " +
                        "StartTime       StartTime, " +
                        "endTime         endTime, " +
                        "[DESC]          [DESC], " +
                        "DICVALUE        DICVALUE " +
                        "FROM T_QW_MUSTTIME, " +
                        "(SELECT * FROM T_QW_SYS_DIC WHERE PDICID = 187) DIC " +
                        "WHERE T_QW_MUSTTIME.Type = DIC.DICID";
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getPoliceCrossing(int interval)
    {
        //定义存放的数据表
        //String endTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        String beginTime = (DateTime.Now.AddMinutes(-interval)).ToString("yyyy-MM-dd HH:mm:ss");
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                     "GPSID         GPSID, " +
                     "GPSSIMCARD    GPSSIMCARD, " +
                     "GPSTYPE       GPSTYPE, " +
                     "GPSNAME       GPSNAME, " +
                     "GPSDEPID      GPSDEPID, " +
                     "GPSDEPNAME    GPSDEPNAME, " +
                     "POLICENO      POLICENO, " +
                     "PHONE         PHONE, " +
                     "LONGITUDE     LONGITUDE, " +
                     "LATITUDE      LATITUDE, " +
                     "DATARECORDTIME      DATARECORDTIME, " +
                     "DUTYTYPE      QWTYPE, " +
                     "CURRENTSTATE  QWSTATUS, " +
                     "STATENM       QWSTATUSNAME " +
                        "FROM T_QW_SYS_GPSDATA, " +
            //过滤值班DUTYTYPE = 1
                        "(SELECT STARTTIME,ENDTIME,JH,CURRENTSTATE,STATENM,MAX(DUTYTYPE) DUTYTYPE " +
                        "FROM View_PLANDUTYSTATE " +
                        "GROUP BY STARTTIME,ENDTIME,JH,CURRENTSTATE,STATENM) PLANDUTYSTATE "
            //过滤值班DUTYTYPE = 1
                        + "WHERE GPSTYPE <> 1 "
                        + "AND PLANDUTYSTATE.JH = T_QW_SYS_GPSDATA.POLICENO "
                        + "AND PLANDUTYSTATE.STARTTIME <= T_QW_SYS_GPSDATA.DATARECORDTIME "
                        + "AND PLANDUTYSTATE.ENDTIME >= T_QW_SYS_GPSDATA.DATARECORDTIME "
                        + "AND CONVERT(VARCHAR,DATARECORDTIME,120) > '" + beginTime + "' "
                        + "ORDER BY DATARECORDTIME";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getAuth(String policeID)
    {
        //定义存放的数据表
         ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
         String sql = "";
         if (policeID != "")
         {
             sql = "SELECT " +
                     "DWXX.DEPID    GPSDEPID, " +
                     "DWXX.DWMC     GPSDEPNAME, " +
                     "JYXX.ID       USERID, " +
                     "JYXX.RYXM     GPSNAME " +
                     "FROM V_QWGLXT_JYXX_1 JYXX,V_DWXX DWXX " +
                     "WHERE JYXX.ID = '" + policeID + "' AND JYXX.SSGZZ = DWXX.DEPID";
         }
         else
         {
             sql = "SELECT " +
                     "DWXX.DEPID    GPSDEPID, " +
                     "DWXX.DWMC     GPSDEPNAME, " +
                     "JYXX.ID       USERID, " +
                     "JYXX.RYXM     GPSNAME " +
                     "FROM V_QWGLXT_JYXX_1 JYXX,V_DWXX DWXX " +
                     "WHERE JYXX.RYXM = 'admin' AND JYXX.SSGZZ = DWXX.DEPID";
         }
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable setCallInfo(String isresponse, String isconfirm, String frequency, String stid, String jyname, String depname, String jyid, String depid, String username, String userid, String userdepname, String userdepid, String note)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "insert into T_QW_DM ([DMDate],[STID],[IsResponse],[ISCONFIRM],[frequency],[BDMNAME],[BDMDepName],[BDMJYID],[BDMDEPID] " +
                        ",[DMNAME],[DMJYID],[DMDEPNAME],[DMDEPID],[MEMO]) " +
                        "values ('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + stid + "','" + isresponse + "'," + isconfirm + "," + frequency + ",'" + jyname + "','" + depname + "'," + jyid + ",'" + depid + "','" + username + "','" + userid + "','" + userdepname + "','" + userdepid + "','" + note + "')";

        clsGetData.SetTable(sql);

        sql = "SELECT TOP 1 " +
                       "ID , " +
                       "DMDate , " +
                       "IsResponse , " +
                       "ResponseDate , " +
                       "BDMNAME , " +
                       "BDMDepName , " +
                       "BDMJYID , " +
                       "BDMDEPID , " +
                       "DMNAME , " +
                       "DMJYID , " +
                       "DMDEPNAME , " +
                       "MEMO , " +
                       "ISCONFIRM " +
                       "FROM T_QW_DM WHERE BDMJYID = '" + jyid + "' AND STID = '" + stid + "' " +
                       "ORDER BY DMDate DESC";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable setServiceState(String stateid, String statenm, String jyid, String statusdes, String username,String gpsTime)
    {
        DateTime now = DateTime.Now;

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        Int32 count = Convert.ToInt32(clsGetData.GetValue("SELECT COUNT(*) FROM T_QW_SYS_JYAPPENDINFO WHERE JYID = '" + jyid + "'"));

        String sql = "";

        if (count > 0)
        {
            sql = "update T_QW_SYS_JYAPPENDINFO set " +
                             "currentstate = '" + stateid +
                             "',STATECHANGETIME ='" + now.ToString("yyyy-MM-dd HH:mm:ss") +
                             "',STATENM= '" + statenm + "' where JYID = '" + jyid + "'";
        }
        else
        {
            sql = "insert into T_QW_SYS_JYAPPENDINFO (JYID,CURRENTSTATE " +
                    ",STATUDES,STATECHANGETIME,GPSSTATUS,GPSSTATUSCHANGETIME,STATENM)  " +
                    "values ('" + jyid + "','" + stateid + "','', getdate(), 1, '" + gpsTime + "', '" + statenm + "')";
        }

        sql += ";insert into T_QW_STATUSCHANGELOG(STATUS,CHANGETIME,STATUSDES,CHANGEUSER,JYID) "+
                "values ('" + stateid + "','" + now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + statusdes + "','" + username + "','" + jyid + "')";

        if (statenm == "异常")
        {
            DataTable gpsTable = clsGetData.GetTable("SELECT * FROM View_GPSPLAN WHERE USERID = " + jyid);
            if(gpsTable.Rows.Count > 0)
            {
                DataRow row = gpsTable.Rows[0];

                sql += ";INSERT INTO T_QW_UNNORMALQW(UnNormalType,GpsIDOrZoneID,GPSNameOrZoneName,DepID,DepName,ReportDateTime,X,Y) " +
                    "VALUES (6,'" + row["GPSSIMCARD"].ToString() + "','" 
                    +  row["GPSNAME"].ToString() + "','"
                    + row["GPSDEPID"].ToString() + "','"
                    + row["GPSDEPNAME"].ToString() + "','" 
                    + now.ToString("yyyy-MM-dd HH:mm:ss") + "','" 
                    + row["LONGITUDE"].ToString() + "','" 
                    + row["LATITUDE"].ToString() + "')";
            }
        }

        clsGetData.ExcuteNoQuery(sql);

        sql = "SELECT " +
                     "GPSID         GPSID, " +
                     "GPSSIMCARD    GPSSIMCARD, " +
                     "GPSTYPE       GPSTYPE, " +
                     "GPSNAME       GPSNAME, " +
                     "GPSDEPID      GPSDEPID, " +
                     "GPSDEPNAME    GPSDEPNAME, " +
                     "POLICENO      POLICENO, " +
                     "PHONE         PHONE, " +
                     "LONGITUDE     LONGITUDE, " +
                     "LATITUDE      LATITUDE, " +
                     "DATARECORDTIME      DATARECORDTIME, " +
                     "QWID          QWID, " +
                     "DUTYTYPE      QWTYPE, " +
                     "CURRENTSTATE  QWSTATUS, " +
                     "STATENM       QWSTATUSNAME, " +
                     "CARNO         CARNO, " +
                     "USERID        USERID, " +
                     "STID          STID, " +
                     "ZONENM        ZONENM, " +
                     "PATROLTYPE    PATROLTYPE, " +
                     "RUNNAME       RUNNAME, " +
                     "STARTTIME     STARTTIME, " +
                     "ENDTIME       ENDTIME, " +
                     "STATECHANGETIME     STATECHANGETIME " +
                     "FROM VIEW_GPSPLAN WHERE USERID = '" + jyid + "'";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public String setAlarmFocus(String JQID, String ISMUST)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "Update T_QW_JQ  set ISMUST = " + ISMUST + " where JQDH = '" + JQID + "'";

        clsGetData.SetTable(sql);

       return clsGetData.ErrorString;
    }

    [WebMethod]
    public DataTable setAlarmPolice(String JQID, String JYID,String Content,String UpdateUser,String Memo)
    {
        DateTime now = DateTime.Now;

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        
        String sql = "insert into T_JQ_DealJYRela (JQID,JYID) values ('"+JQID +"','"+JYID +"')";

        sql += ";insert into T_JQ_Process (JQID,JYID,pTYPE,FromType,TypeDateTime,Content,UpdateUser,Memo) " +
                "values ('" + JQID + "','" + JYID + "',0,2,'" + now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Content + "','" + UpdateUser + "','" + Memo + "')";

       clsGetData.ExcuteNoQuery(sql);

       sql = "SELECT " +
                      "T_JQ_Process.ID , " +
                      "JQID , " +
                      "JYID , " +
                      "TypeDateTime , " +
                      "pType " +
                      "FROM T_JQ_Process ,( " +
                      "SELECT MAX(ID) ID " +
                      "FROM T_JQ_Process " +
                      "GROUP BY JQID,JYID) MAXID " +
                      "WHERE T_JQ_Process.ID = MAXID.ID " +
                      "AND T_JQ_Process.JQID = '" + JQID + "'";

       return clsGetData.GetTable(sql);
    }
    
    [WebMethod]
    public DataTable deleteAlarmPolice(String JQID, String JYID)
    {        
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "delete T_JQ_DealJYRela where JQID = '" + JQID + "' and JYID = '" + JYID + "'";

        sql += ";delete T_JQ_Process where JQID = '" + JQID + "' and JYID = '" + JYID + "'";

       clsGetData.ExcuteNoQuery(sql);

       sql = "SELECT " +
                      "T_JQ_Process.ID , " +
                      "JQID , " +
                      "JYID , " +
                      "TypeDateTime , " +
                      "pType " +
                      "FROM T_JQ_Process ,( " +
                      "SELECT MAX(ID) ID " +
                      "FROM T_JQ_Process " +
                      "GROUP BY JQID,JYID) MAXID " +
                      "WHERE T_JQ_Process.ID = MAXID.ID " +
                      "AND T_JQ_Process.JQID = '" + JQID + "'";

       return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public String setAlarmPoliceType(String JQID, String JYID,String pTYPE, String UpdateUser, String Memo)
    {
        DateTime now = DateTime.Now;

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "update T_JQ_Process set pTYPE =  " + pTYPE  +
                        ",FromType = 1,TypeDateTime = '" + now.ToString("yyyy-MM-dd HH:mm:ss") + "',UpdateUser = '" + UpdateUser + "',Memo= '" + Memo + "' " +
                        "where JQID='" + JQID + "' and JYID='" + JYID + "'";

        clsGetData.SetTable(sql);

        return clsGetData.ErrorString;
    }

    [WebMethod]
    public String setAlarmPos(String JQID, String X, String Y)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        clsGetData.SetTable("Update T_QW_JQ set NEWX= '" + X + "',NEWY='" + Y + "' where JQDH = '" + JQID + "'");

        return clsGetData.ErrorString;
    }

    private int getFieldID(MapServerInfo mapinfo, String layerName,String fieldName)
    {
        MapLayerInfo[] maplayerinfos = mapinfo.MapLayerInfos;

        int geometryfieldid = 0;

        foreach (MapLayerInfo layerinfo in maplayerinfos)
        {
            if ((layerinfo.Name == layerName) && (layerinfo.LayerType == "Feature Layer"))
            {
                Field[] fields = layerinfo.Fields.FieldArray;
                foreach (Field field in fields)
                {
                    if (fieldName.ToUpper() == field.Name.ToUpper())
                    {
                        break;
                    }

                    geometryfieldid++;
                }
                break;
            }
        }
        return geometryfieldid;
    }

    private int getLayerID(MapServerInfo mapinfo,String layerName)
    {
        MapDescription mapdesc = mapinfo.DefaultMapDescription;
        MapLayerInfo[] maplayerinfos = mapinfo.MapLayerInfos;

        int layerid = 0;
        string geometryfieldname = string.Empty;
        foreach (MapLayerInfo layerinfo in maplayerinfos)
        {
            if ((layerinfo.Name == layerName) && (layerinfo.LayerType == "Feature Layer"))
            {
                layerid = layerinfo.LayerID;
                
                Field[] fields = layerinfo.Fields.FieldArray;

                foreach (Field field in fields)
                {
                    if (field.Type == esriFieldType.esriFieldTypeGeometry)
                        geometryfieldname = field.Name;
                }
                break;
            }
        }

        return layerid;
    }

    private PointN getRoadPoint(PolylineN polyline,Double roadlen, Int32 door, Int32 l_f_door, Int32 l_t_door)
    {
        PointN pointN = null;
        Double len = roadlen;

        if (door <= l_f_door)
        {
            SharedNamespace.Path path = polyline.PathArray[polyline.PathArray.Length - 1];
            pointN = path.PointArray[path.PointArray.Length - 1] as PointN;
        }
        else if (door >= l_t_door)
        {
            pointN = polyline.PathArray[0].PointArray[0] as PointN;
        }
        else
        {
            len = len * ((Double)(l_t_door - door) / (l_t_door - l_f_door));

            foreach (SharedNamespace.Path path in polyline.PathArray)
            {
                PointN prePoint = null;
                foreach (PointN p in path.PointArray)
                {
                    if (prePoint == null)
                    {
                        prePoint = p;
                    }
                    else
                    {
                        Double dx = p.X - prePoint.X;
                        Double dy = p.Y - prePoint.Y;
                        Double curLen = Math.Sqrt(dx * dx + dy * dy);
                        if (curLen > len)
                        {
                            pointN = new PointN();
                            pointN.X = prePoint.X + dx * len / curLen;
                            pointN.Y = prePoint.Y + dy * len / curLen;
                        }

                        prePoint = p;
                        len -= curLen;
                    }

                    if (len < 0)
                    {           
                        break;
                    }
                }

                if (len < 0)
                    break;
            }
        }

        return pointN;
    }

    private DataTable createAddressTable()
    {
        DataTable responseTable = new DataTable("Table");
        DataColumn column = new DataColumn();
        column.ColumnName = "Label";
        column.DataType = Type.GetType("System.String");
        responseTable.Columns.Add(column);

        column = new DataColumn();
        column.ColumnName = "X";
        column.DataType = Type.GetType("System.String");
        responseTable.Columns.Add(column);

        column = new DataColumn();
        column.ColumnName = "Y";
        column.DataType = Type.GetType("System.String");
        responseTable.Columns.Add(column);

        return responseTable;
    }

    private Boolean IsQuartor(String name)
    {
        if (name.Trim() == "")
            return false;

        foreach (String quartPredix in quarterPredixList)
        {
            if (name.LastIndexOf(quartPredix) == name.Length - quartPredix.Length)
            {
                return !quarterWhiteList.Contains(name);
            }
        }

        return quarterBlackList.Contains(name);
    }

    [WebMethod]
    public DataTable getCrossRoad(String road)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "select a.名称,a.全名首拼,a.拼音全称,a.类别 " +
                    "from " + strSDEPredix + "道路中心线 a," + strSDEPredix + "交叉路口 b " +
                    "where b.名称 = '" + road + "' and a.名称 <> '" + road + "' and CHARINDEX(a.名称,b.交叉道路1) > 0";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getCrossPoint(String road1,String road2)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "SELECT '"+road1+road2+"' Label,POINT_X X,POINT_Y Y FROM "+strSDEPredix+"交叉路口 "
            + "WHERE 名称 = '" + road1 + "' AND 交叉道路1 LIKE '%" + road2 + "%'";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable addressLocator(String adddress)
    {
        MHGAXM_MapServer mapservice = new MHGAXM_MapServer();
        mapservice.Url = strMapUrl;

        String mapname = mapservice.GetDefaultMapName();
        MapServerInfo serverInfo = mapservice.GetServerInfo(mapname);

        //先从各个点图层Find，一共14个图层
        int[] layerIDs = new int[14];
        layerIDs[0] = getLayerID(serverInfo, "主要大厦");
        layerIDs[1] = getLayerID(serverInfo, "火车站");
        layerIDs[2] = getLayerID(serverInfo, "金融单位");
        layerIDs[3] = getLayerID(serverInfo, "娱乐场所");
        layerIDs[4] = getLayerID(serverInfo, "学校");
        layerIDs[5] = getLayerID(serverInfo, "小区");
        layerIDs[6] = getLayerID(serverInfo, "政府机关");
        layerIDs[7] = getLayerID(serverInfo, "派出所");
        layerIDs[8] = getLayerID(serverInfo, "公交站点");
        layerIDs[9] = getLayerID(serverInfo, "加油站");
        layerIDs[10] = getLayerID(serverInfo, "医院");
        layerIDs[11] = getLayerID(serverInfo, "区公安局");
        layerIDs[12] = getLayerID(serverInfo, "地铁车站");
        layerIDs[13] = getLayerID(serverInfo, "消防中队");
        //layerIDs[14] = getLayerID(serverInfo, "门牌号");

        MapServerFindResult[] findResults = mapservice.Find(serverInfo.DefaultMapDescription, new ImageDisplay(), adddress, true, "名称", esriFindOption.esriFindAllLayers, layerIDs);

        DataTable responseTable;
        if (findResults.Length > 0)
        {
            responseTable = createAddressTable();

            foreach (MapServerFindResult findResult in findResults)
            {
                PointN pointN = findResult.Shape as PointN;

                DataRow row = responseTable.NewRow();
                row["Label"] = findResult.Value;
                row["X"] = pointN.X;
                row["Y"] = pointN.Y;
                responseTable.Rows.Add(row);
            }

            return responseTable;
        }

        Regex regRoad = new Regex(@"(\w+?[路道街])(.*)$", RegexOptions.Singleline);
        Match matchRoad = regRoad.Match(adddress);

        Regex regQuartor = new Regex(@"^(.*?)(\d+)(号|$)", RegexOptions.Singleline);
        Match matchQuartor = regQuartor.Match(adddress);

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql;

        //道路匹配
        if (matchRoad.Success)
        {
            String road1 = matchRoad.Groups[1].Value;
            adddress = (matchRoad.Groups[2].Value as String).Trim();

            QueryFilter queryfilter = new QueryFilter();
            queryfilter.WhereClause = "名称 = '" + road1 + "'";
            RecordSet recordset = mapservice.QueryFeatureData(mapname, getLayerID(serverInfo, "道路中心线"), queryfilter);

            //道路匹配-找到道路
            if ((recordset != null) && (recordset.Records.Length > 0))
            {
                Record record = recordset.Records[0];

                PolylineN polyline = record.Values[getFieldID(serverInfo, "道路中心线", "SHAPE")] as PolylineN;
                Int32 l_f_door = Convert.ToInt32(record.Values[getFieldID(serverInfo, "道路中心线", "左起门牌")]);
                Int32 l_t_door = Convert.ToInt32(record.Values[getFieldID(serverInfo, "道路中心线", "左止门牌")]);
                Double len = Convert.ToDouble(record.Values[getFieldID(serverInfo, "道路中心线", "SHAPE.len")]);
                if ((l_f_door == 0)
                    && (l_t_door == 0))
                {
                    l_f_door = 1;
                    l_t_door = Convert.ToInt32(len / 5);
                }

                Regex regCrossRoad = new Regex(@"近?(\w+?[路道街])", RegexOptions.Singleline);
                Match matchCrossRoad = regCrossRoad.Match(adddress);

                Regex regAlley = new Regex(@"(\d+)弄", RegexOptions.Singleline);
                Match matchAlley = regAlley.Match(adddress);

                Regex regDoorplate = new Regex(@"^(.*?)(\d+)(号|$)", RegexOptions.Singleline);
                Match matchDoorplate = regDoorplate.Match(adddress);

                String doorplate = "";
                String quartor = "";
                if (matchDoorplate.Success)
                {
                    quartor = matchDoorplate.Groups[1].Value;
                    doorplate = matchDoorplate.Groups[2].Value;
                }

                //道路匹配-找到道路-交叉路
                if (matchCrossRoad.Success)
                {
                    String road2 = matchCrossRoad.Groups[1].Value;

                    return getCrossPoint(road1, road2);
                }
                //道路匹配-找到道路-弄
                else if (matchAlley.Success)
                {
                    String alley = matchAlley.Groups[1].Value;

                    if (doorplate != "")
                    {
                        //sql = "SELECT 道路名+门牌号+'(' + 名称 + ')' Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        //  + " WHERE 道路名 = '" + road1 + "' AND 弄号 = '" + alley + "弄' AND 门牌号 = '" + doorplate + "号'";
                        sql = "SELECT 道路名 + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                          + " WHERE 道路名 = '" + road1 + "' AND 门牌号 = '" + alley + "弄" + doorplate + "号'";
                        responseTable = clsGetData.GetTable(sql);

                        if (responseTable.Rows.Count > 0)
                        {
                            return responseTable;
                        }
                    }

                    //道路匹配-找到道路-弄-未匹配到门牌号
                    sql = "SELECT 道路名 + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        + " WHERE 道路名 = '" + road1 + "' AND 门牌 = '" + alley + "弄'";
                    responseTable = clsGetData.GetTable(sql);

                    if (responseTable.Rows.Count > 0)
                    {
                        PointN pointN = new PointN();
                        pointN.X = 0;
                        pointN.Y = 0;
                        foreach (DataRow iRow in responseTable.Rows)
                        {
                            pointN.X += Convert.ToDouble(iRow["X"]);
                            pointN.Y += Convert.ToDouble(iRow["Y"]);
                        }
                        pointN.X = pointN.X / responseTable.Rows.Count;
                        pointN.Y = pointN.Y / responseTable.Rows.Count;

                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + alley + "弄" + (doorplate != ""?doorplate + "号'":"") + "(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.InsertAt(row,0);
                    }

                    return responseTable;
                }
                //道路匹配-找到道路-小区
                else if (quartor != "")
                {
                    if (doorplate != "")
                    {
                        sql = "SELECT 道路名 + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                            + " WHERE 道路名 = '" + road1 + "' AND 名称 = '" + quartor + "' AND 门牌号 = LTrim(RTrim(门牌)) + '" + doorplate + "号'";
                        responseTable = clsGetData.GetTable(sql);

                        if (responseTable.Rows.Count > 0)
                        {
                            return responseTable;
                        }
                    }

                    //道路匹配-找到道路-小区-未匹配到门牌号
                    sql = "SELECT 道路名 + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        + " WHERE 道路名 = '" + road1 + "' AND 名称 = '" + quartor + "'";
                    responseTable = clsGetData.GetTable(sql);

                    if (responseTable.Rows.Count > 0)
                    {
                        PointN pointN = new PointN();
                        pointN.X = 0;
                        pointN.Y = 0;
                        foreach (DataRow iRow in responseTable.Rows)
                        {
                            pointN.X += Convert.ToDouble(iRow["X"]);
                            pointN.Y += Convert.ToDouble(iRow["Y"]);
                        }
                        pointN.X = pointN.X / responseTable.Rows.Count;
                        pointN.Y = pointN.Y / responseTable.Rows.Count;

                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + quartor + (doorplate != "" ? doorplate + "号'" : "") + "(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.InsertAt(row, 0);
                    }

                    return responseTable;
                }
                //道路匹配-找到道路-门牌号
                else if (doorplate != "")
                {
                    //sql = "SELECT 道路名 + 门牌号 + CASE WHEN a.单位名称 = '' THEN '' ELSE '(' + a.单位名称 + ')' END Label,POINT_X X,POINT_Y Y "
                    //        + "FROM " + strSDEPredix + "门牌号码 a LEFT JOIN("
                    //        + "SELECT COUNT(*) C_UNIT,单位名称 "
                    //        + "FROM " + strSDEPredix + "门牌号码 "
                    //        + "WHERE 单位名称 <> '' "
                    //        + "GROUP BY 单位名称 "
                    //        + ") b ON a.单位名称 = b.单位名称 "
                    //        + "WHERE ISNULL(C_UNIT,0) <= 5 AND 道路名 = '" + road1 + "' AND 门牌号 = '" + doorplate + "号' AND 弄号 = '' AND 小区名 = '' "
                    //        + "ORDER BY ISNULL(C_UNIT,0)";

                    sql = "SELECT 道路名 + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,名称,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        + " WHERE 道路名 = '" + road1 + "' AND 门牌号 = '" + doorplate + "号'";
                    DataTable tempTable = clsGetData.GetTable(sql);
                    responseTable = createAddressTable();

                    //小区门牌号过滤
                    foreach (DataRow tempRow in tempTable.Rows)
                    {
                        if(IsQuartor(tempRow["名称"].ToString()))
                            continue;

                        DataRow row = responseTable.NewRow();
                        row["Label"] = tempRow["Label"];
                        row["X"] = tempRow["X"];
                        row["Y"] = tempRow["Y"];
                        responseTable.Rows.Add(row);
                    }

                    if (responseTable.Rows.Count > 0)
                    {
                        return responseTable;
                    }

                    //道路匹配-找到道路-门牌号-未匹配到门牌号
                    sql = "SELECT 门牌号,名称,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        + " WHERE 道路名 = '" + road1 + "' AND LTrim(RTrim(门牌)) = ''";
                    responseTable = clsGetData.GetTable(sql);

                    DataRow maxItem = null, minItem = null;
                    Int32 maxDoor = Int32.MaxValue, minDoor = Int32.MinValue;
                    Int32 nDoorplate = Convert.ToInt32(doorplate);

                    foreach (DataRow row in responseTable.Rows)
                    {
                        PointN point = new PointN();
                        point.X = Convert.ToDouble(row["X"]);
                        point.Y = Convert.ToDouble(row["Y"]);

                        if (IsQuartor(row["名称"].ToString()))
                            continue;

                        if (ptopointset(point, polyline) > 50)
                            continue;

                        String curDoor = row["门牌号"].ToString().Trim();
                        //String curUnit = row["单位名称"].ToString().Trim();

                        Int32 nCurdoor;
                        Int32 nIndex = curDoor.IndexOf("号");
                        if ((nIndex < 0)
                            || (!Int32.TryParse(curDoor.Substring(0, nIndex), out nCurdoor)))
                            continue;

                        if ((nDoorplate % 2) != (nCurdoor % 2))
                            continue;

                        if (nCurdoor > nDoorplate)
                        {
                            if ((maxItem == null) || (maxDoor > nCurdoor))
                            {
                                maxItem = row;
                                maxDoor = nCurdoor;
                            }
                        }

                        if (nCurdoor < nDoorplate)
                        {
                            if ((minItem == null) || (nCurdoor > minDoor))
                            {
                                minItem = row;
                                minDoor = nCurdoor;
                            }
                        }
                    }

                    if ((minItem != null) && (maxItem != null))
                    {
                        PointN minPointN = new PointN();
                        minPointN.X = Convert.ToDouble(minItem["X"]);
                        minPointN.Y = Convert.ToDouble(minItem["Y"]);

                        PointN maxPointN = new PointN();
                        maxPointN.X = Convert.ToDouble(maxItem["X"]);
                        maxPointN.Y = Convert.ToDouble(maxItem["Y"]);

                        Double scale = ((Double)(nDoorplate - minDoor) / (maxDoor - minDoor));

                        PointN pointN = new PointN();
                        pointN.X = minPointN.X + (maxPointN.X - minPointN.X) * scale;
                        pointN.Y = minPointN.Y + (maxPointN.Y - minPointN.Y) * scale;

                        responseTable = createAddressTable();
                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + doorplate + "号(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.Add(row);
                        return responseTable;
                    }
                    else if ((minItem != null) && (maxItem == null))
                    {
                        PointN minPointN = new PointN();
                        minPointN.X = Convert.ToDouble(minItem["X"]);
                        minPointN.Y = Convert.ToDouble(minItem["Y"]);

                        PointN maxPointN = polyline.PathArray[0].PointArray[0] as PointN;
                        //maxPointN.X = Convert.ToDouble(maxItem["X"]);
                        //maxPointN.X = Convert.ToDouble(maxItem["Y"]);

                        Double scale = (l_t_door > nDoorplate)?((Double)(nDoorplate - minDoor) / (l_t_door - minDoor)):1;
                        ;

                        PointN pointN = new PointN();
                        pointN.X = minPointN.X + (maxPointN.X - minPointN.X) * scale;
                        pointN.Y = minPointN.Y + (maxPointN.Y - minPointN.Y) * scale;

                        responseTable = createAddressTable();
                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + doorplate + "号(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.Add(row);
                        return responseTable;
                    }
                    else if ((maxItem != null) && (minItem == null))
                    {
                        SharedNamespace.Path path = polyline.PathArray[polyline.PathArray.Length - 1];

                        PointN minPointN = path.PointArray[path.PointArray.Length - 1] as PointN;
                        //minPointN.X = Convert.ToDouble(minItem["X"]);
                        //minPointN.X = Convert.ToDouble(minItem["Y"]);

                        PointN maxPointN = polyline.PathArray[0].PointArray[0] as PointN;
                        maxPointN.X = Convert.ToDouble(maxItem["X"]);
                        maxPointN.Y = Convert.ToDouble(maxItem["Y"]);

                        Double scale = (l_f_door < nDoorplate)?((Double)(nDoorplate - l_f_door) / (maxDoor - l_f_door)):0;

                        PointN pointN = new PointN();
                        pointN.X = minPointN.X + (maxPointN.X - minPointN.X) * scale;
                        pointN.Y = minPointN.Y + (maxPointN.Y - minPointN.Y) * scale;

                        responseTable = createAddressTable();
                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + doorplate + "号(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.Add(row);
                        return responseTable;
                    }
                    else
                    {
                        PointN pointN = getRoadPoint(polyline, len, Convert.ToInt32(doorplate), l_f_door, l_t_door);

                        responseTable = createAddressTable();
                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + doorplate + "号(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.Add(row);
                        return responseTable;
                    }
                }
                //道路匹配-找到道路-没有门牌号
                else if (adddress != "")
                {
                    sql = "SELECT 道路名 + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        + " WHERE 道路名 = '" + road1 + "' AND (名称 = '" + adddress + "' OR 门牌 = '" + adddress + "')";
                    responseTable = clsGetData.GetTable(sql);

                    if (responseTable.Rows.Count > 0)
                    {
                        PointN pointN = new PointN();
                        pointN.X = 0;
                        pointN.Y = 0;
                        foreach (DataRow iRow in responseTable.Rows)
                        {
                            pointN.X += Convert.ToDouble(iRow["X"]);
                            pointN.Y += Convert.ToDouble(iRow["Y"]);
                        }
                        pointN.X = pointN.X / responseTable.Rows.Count;
                        pointN.Y = pointN.Y / responseTable.Rows.Count;

                        DataRow row = responseTable.NewRow();
                        row["Label"] = road1 + adddress + "(近似地点)";
                        row["X"] = pointN.X;
                        row["Y"] = pointN.Y;
                        responseTable.Rows.InsertAt(row, 0);
                    }

                    return responseTable;
                }
                else
                {
                    return createAddressTable();
                }
            }
            //道路匹配-找不到道路
            else
            {
                return createAddressTable();
            }
        }
        //小区匹配
        else if (matchQuartor.Success)
        {
            String quartor = matchQuartor.Groups[1].Value;
            String doorplate = matchQuartor.Groups[2].Value;

            if (doorplate != "")
            {
                sql = "SELECT LTrim(RTrim(道路名)) + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                    + " WHERE 名称 = '" + quartor + "' AND 门牌号 = LTrim(RTrim(门牌)) + '" + doorplate + "号'";
                responseTable = clsGetData.GetTable(sql);

                if (responseTable.Rows.Count > 0)
                {
                    return responseTable;
                }
            }

            //道路匹配-找到道路-小区-未匹配到门牌号
            sql = "SELECT LTrim(RTrim(道路名)) + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                + " WHERE 名称 = '" + quartor + "'";
            responseTable = clsGetData.GetTable(sql);

            if (responseTable.Rows.Count > 0)
            {
                PointN pointN = new PointN();
                pointN.X = 0;
                pointN.Y = 0;
                foreach (DataRow iRow in responseTable.Rows)
                {
                    pointN.X += Convert.ToDouble(iRow["X"]);
                    pointN.Y += Convert.ToDouble(iRow["Y"]);
                }
                pointN.X = pointN.X / responseTable.Rows.Count;
                pointN.Y = pointN.Y / responseTable.Rows.Count;

                DataRow row = responseTable.NewRow();
                row["Label"] = quartor + (doorplate != "" ? doorplate + "号'" : "") + "(近似地点)";
                row["X"] = pointN.X;
                row["Y"] = pointN.Y;
                responseTable.Rows.InsertAt(row, 0);
            }

            return responseTable;
        }
        //其他匹配
        else
        {
            sql = "SELECT LTrim(RTrim(道路名)) + 门牌号 + CASE 名称 WHEN '' THEN '' ELSE '(' + 名称 + ')' END Label,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                + " WHERE 名称 LIKE '%" + adddress + "%'";

            responseTable = clsGetData.GetTable(sql);

            return responseTable;
        }
    }

    [WebMethod]
    public DataTable getExceptService(String deptID)
    {
        //String sql = "SELECT TOP 20 UnNormalType,GpsIDOrZoneID,GPSNameOrZoneName,DepID,DepName,ReportDateTime,UnNormalDesc,X,Y,Memo FROM T_QW_UNNORMALQW";
        String sql = "SELECT T_QW_UNNORMALQW.* "
            + "FROM T_QW_UNNORMALQW LEFT JOIN V_DWXX ON T_QW_UNNORMALQW.DepID = V_DWXX.DEPID"
            + " WHERE ((UnNormalType <> 5 AND UnNormalType <> 6 AND ReportDateTime > DATEADD(MI,-30,GETDATE()))"
            + " OR ((UnNormalType = 5 OR UnNormalType = 6) AND ReportDateTime > DATEADD(HH,-12,GETDATE())))";

        if (deptID == "-1")
        {
        }
        else if (deptID == "-2")
        {
            sql += " AND V_DWXX.ZB = 125 ";
        }
        else
            sql += " AND T_QW_UNNORMALQW.DepID = " + deptID;

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        return clsGetData.GetTable(sql);
    }
    
          // 判断点象限函数
     private Int32 getQuad(Double x,Double y) 
     {
        return (x >= 0) ? ((y >= 0) ? 0 : 3) : ((y >= 0) ? 1 : 2);
     }

    private Boolean polygonContainPoint(String polygonCoords,Double x,Double y)
    {
        String[] poly = polygonCoords.Split(';');

        // 检查顶点数
        if(poly.Length < 3)
            return true;
    
        // 平移多边形，使point在新坐标系中为原点
        Double[,] vPoly = new Double[poly.Length,2];
        for(int i=0;i<poly.Length;i++)
        {
            Double tempX,tempY;
            if(poly[i].IndexOf(' ') >= 0)
            {
                tempX = Convert.ToDouble(poly[i].Split(' ')[0]);
                tempY = Convert.ToDouble(poly[i].Split(' ')[1]);
            }
            else
            {
                tempX = Convert.ToDouble(poly[i].Split(',')[0]);
                tempY = Convert.ToDouble(poly[i].Split(',')[1]);
            }
            vPoly[i,0] = tempX - x;
            vPoly[i,1] = tempY - y;
        }

        Int32 sum = 0;//弧长和*2/π
        Int32 q1, q2;//相邻两节点的象限
        Double ep = 0;//用来存放外积
        Boolean dq = false;//用来存放两点是否在相对象限中

        q1 = getQuad(vPoly[0,0],vPoly[0,1]);
        for (int i = 1; i < poly.Length; i++)
        {
             // point为多边形的一个节点
            if(vPoly[i,0] == 0 && vPoly[i,1] == 0)
                return true;

             // 计算两点向量外积，用来判断点是否在多边形边上（即三点共线）
            ep = vPoly[i,1] * vPoly[i-1,0] - vPoly[i,0] * vPoly[i-1,1];
            dq = (vPoly[i-1,0] * vPoly[i,0] <= 0) && (vPoly[i-1,1] * vPoly[i,1] <= 0);
            if ((ep==0) && dq)
                 return true;

            // 计算象限判断相邻两点的象限关系，并修改sum
            q2 = getQuad(vPoly[i,0],vPoly[i,1]);
             if (q2 == (q1+1)%4) 
             {
                 sum = sum + 1;
             }
             else if (q2 == (q1+3)%4) 
             {
                  sum = sum - 1;
            } 
             else if (q2 == (q1+2)%4) 
             {
               if (ep > 0) 
                    sum = sum + 2;
               else 
                    sum = sum - 2;
            }
             q1 = q2;
        }
        if (sum > 0)
             return true;

        return false;
    }
    /* 计算点到折线集的最近距离,并返回最近点.
注意：调用的是ptolineseg()函数 */
double ptopointset(PointN point,PolylineN road)
{    
    Double cd =Double.MaxValue;
    for(int i=0;i<road.PathArray.Length;i++)
    {
        SharedNamespace.Path path = road.PathArray[i];
        for(int j=0;j<path.PointArray.Length - 1;j++)
        {
            Double td = ptolinesegdist(point,path.PointArray[j] as PointN,path.PointArray[j+1] as PointN);
            cd = Math.Min(cd,td);
        }
    }
return cd;
}
    /* 求点p到线段l的最短距离,并返回线段上距该点最近的点np
注意：np是线段l上到点p最近的点，不一定是垂足 */
double ptolinesegdist(PointN p, PointN s, PointN e)
{
double r=relation(p,s,e);
if(r<0)
{
return dist(p,s);
}
if(r>1)
{
return dist(p,e);
}

return dist(p,  perpendicular(p, s,e));
}

// 求点C到线段AB所在直线的垂足 P
PointN perpendicular(PointN p, PointN s,PointN e)
{
    double r = relation(p,s,e);
    PointN tp = new PointN();
    tp.X = s.X + r * (e.X -s.X);
    tp.Y = s.Y + r * (e.Y -s.Y);
    return tp;
} 
double relation(PointN p, PointN s,PointN e)
{
    return dotmultiply(p, e, s) / (dist(s, e) * dist(s, e));
}
/*****************************************************************************
**
r=dotmultiply(p1,p2,op),得到矢量(p1-op)和(p2-op)的点积，如果两个矢量都非零矢量

r<0:两矢量夹角为锐角；r=0：两矢量夹角为直角；r>0:两矢量夹角为钝角
******************************************************************************
*/
double dotmultiply(PointN p1, PointN p2, PointN p0)
{
    return ((p1.X - p0.X) * (p2.X - p0.X) + (p1.Y - p0.Y) * (p2.Y - p0.Y));

}
double dist(PointN p1, PointN p2) // 返回两点之间欧氏距离
{
    return (Math.Sqrt((p1.X - p2.X) * (p1.X - p2.X) + (p1.Y - p2.Y) * (p1.Y - p2.Y)));
} 

    [WebMethod]
    public DataTable getWarningArea(String deptID)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "T_QW_PoliceZone.[ID] " +
                        ",[LEVEL] " +
                        ",T_QW_PoliceZone.[DEPID] " +
                        ",[NAME] " +
                        ",[GPSRANGE] " +
                        ",DICVALUE " +
                        "FROM [T_QW_PoliceZone] " +
                        "LEFT JOIN [T_QW_SYS_DIC] ON  [T_QW_PoliceZone].[LEVEL] = T_QW_SYS_DIC.DICID " +
                        "LEFT JOIN V_DWXX ON T_QW_PoliceZone.DEPID = V_DWXX.DEPID ";

        if (deptID == "-1")
        {
        }
        else if (deptID == "-2")
        {
            sql += " WHERE V_DWXX.ZB = 125 ";
        }
        else
            sql += " WHERE T_QW_PoliceZone.DepID = " + deptID;

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getNewTask()
    {
        String sql = "SELECT "
            + "A.TASKNAME "
            + ",A.SCENECOMMAND "
            + ",A.STARTDATE "
            + ",A. COMMCHANEL "
            + ",A.CONTACTWAY "
            + ",A.TASKID "
            + ",A.COMMANDER "
            + "FROM  T_QW_NewTask  A "
            + "WHERE (A.STARTDATE <= GETDATE()) AND (ISNULL(A.ENDDATE, GETDATE()) <= GETDATE()) AND (A.CURRENTSTATUS = 2)";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getNewTaskPolice(String taskID)
    {
        String sql = "SELECT DISTINCT "
            + "A.TASKID "
            + ", A.TASKNAME "
            + ", D.UERID AS USERID "
            + ",D.POLICENAME "
            + ",E.STID "
            + ",E.HH "
            + ",E.XB "
            + ",B.GROUPNAME "
            + ",B.POSTNAME "
            + ",F.DEPID "
            + ",F.DWMC "
            + ",D.ISPLAINCLOTHES "
            + ",E.JH FROM "
            + "T_QW_NEWTASK A "
            + "LEFT JOIN DBO.T_QW_TASK_GROUPSTEP B ON A.TASKID=B.TASKID "
            + "LEFT JOIN T_QW_TASK_ISSUEDDEP C ON B.STAGEID=C.STAGEID "
            + "LEFT JOIN T_QW_TASK_POLICESET D ON C.ISSUEDID=D.ISSUEDID "
            + "LEFT JOIN V_QWGLXT_JYXX_1 E ON D.UERID=E.ID "
            + "LEFT JOIN V_DWXX F ON F.DEPID=C.DEPID "
            + "WHERE A.STARTDATE<=GETDATE() AND ISNULL(A.ENDDATE,GETDATE())<=GETDATE() AND  A.CURRENTSTATUS=2"
            + " AND A.TASKID = " + taskID;

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public String ErrorRoad()
    {
        String result = "";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        MHGAXM_MapServer mapservice = new MHGAXM_MapServer();
        mapservice.Url = strMapUrl;

        String mapname = mapservice.GetDefaultMapName();
        MapServerInfo serverInfo = mapservice.GetServerInfo(mapname);

        QueryFilter queryfilter = new QueryFilter();
        queryfilter.WhereClause = "1 = 1";
        RecordSet recordset = mapservice.QueryFeatureData(mapname, getLayerID(serverInfo, "道路中心线"), queryfilter);
        
        if ((recordset != null) && (recordset.Records.Length > 0))
        {
            for (int i = 0; i < recordset.Records.Length; i++)
            {
                Record record = recordset.Records[i];

                PolylineN polyline = record.Values[getFieldID(serverInfo, "道路中心线", "SHAPE")] as PolylineN;
                Int32 l_f_door = Convert.ToInt32(record.Values[getFieldID(serverInfo, "道路中心线", "左起门牌")]);
                Int32 l_t_door = Convert.ToInt32(record.Values[getFieldID(serverInfo, "道路中心线", "左止门牌")]);
                Double len = Convert.ToDouble(record.Values[getFieldID(serverInfo, "道路中心线", "SHAPE.len")]);
                String roadName = Convert.ToString(record.Values[getFieldID(serverInfo, "道路中心线", "名称")]);

                if ((l_f_door > 0) && (l_t_door > 0))
                {
                    String sql = "SELECT 门牌号,名称,POINT_X X,POINT_Y Y FROM " + strSDEPredix + "门牌号码"
                        + " WHERE 道路名 = '" + roadName + "' AND LTrim(RTrim(门牌)) = ''";

                    DataTable dataTable = clsGetData.GetTable(sql);
                    PointN maxItem = null, minItem = null;
                    Int32 maxDoor = Int32.MaxValue, minDoor = Int32.MinValue;

                    foreach (DataRow row in dataTable.Rows)
                    {
                        PointN point = new PointN();
                        point.X = Convert.ToDouble(row["X"]);
                        point.Y = Convert.ToDouble(row["Y"]);

                        if (IsQuartor(row["名称"].ToString()))
                            continue;

                        if (ptopointset(point, polyline) > 50)
                            continue;

                        String curDoor = row["门牌号"].ToString().Trim();
                        //String curUnit = row["单位名称"].ToString().Trim();

                        Int32 nCurdoor;
                        Int32 nIndex = curDoor.IndexOf("号");
                        if ((nIndex < 0)
                            || (!Int32.TryParse(curDoor.Substring(0, nIndex), out nCurdoor)))
                            continue;

                        if ((maxItem == null) || (maxDoor < nCurdoor))
                        {
                            maxItem = point;
                            maxDoor = nCurdoor;
                        }

                        if ((minItem == null) || (nCurdoor < minDoor))
                        {
                            minItem = point;
                            minDoor = nCurdoor;
                        }
                    }
                    
                    
                    PointN maxPointN = polyline.PathArray[0].PointArray[0] as PointN;
                
                    SharedNamespace.Path path = polyline.PathArray[polyline.PathArray.Length - 1];

                    PointN minPointN = path.PointArray[path.PointArray.Length - 1] as PointN;

                    if ((minItem != null) && (maxItem != null))
                    {
                        if ((dist(minItem, minPointN) > dist(maxItem, minPointN))
                            && (dist(minItem, maxPointN) < dist(maxItem, maxPointN)))
                        result += roadName + ";";
                    }
                }
            }
        }

        return result;
    }

    //奉贤接口
    [WebMethod]
    public DataTable getAlarmInfoFX(String deptName, String typeid, String beginTime, String endTime)
    {
        String sql = "SELECT " +
                        "JQDH id, " +
                        "JQLB type, " +
                        "NEWJQLB newType, " +
                        "BJSJ time, " +
                        "BJDZ address, " +
                        "BJRXM name, " +
                        "JQLB type, " +
                        "JQBT title, " +
                        "BJDH phone, " +
                        "LXDH contactphone, " +
                        "CJR police, " +
                        "JQNR info, " +
                        "SSXQ dept, " +
                        "X x, " +
                        "Y y " +
                        ",NEWX " +
                        ",NEWY " +
                        ",ISMUST " +
                      ",LEVEL typelevel" +
                      ",PID typepid" +
                      ",ID typeid" +
                      ",CODE typecode " +
                      ",COLOR color " +
                        "FROM T_QW_JQ LEFT JOIN T_QW_JQLBTREE ON T_QW_JQ.JQLB = T_QW_JQLBTREE.JQNAME WHERE ";

        if (deptName != "所有单位")
            sql += "SSXQ = '" + deptName + "' AND ";

        //if (type != "所有类别")
        //    sql += "JQLB = '" + type + "' AND ";

        if (beginTime == "")
            sql += "BJSJ > DATEADD(HH,-2,GETDATE()) ";
        else
            sql += "CONVERT(VARCHAR,BJSJ,120) > '" + beginTime + "' ";

        if (endTime == "")
            sql += "AND BJSJ <= GETDATE() ";
        else
            sql += "AND CONVERT(VARCHAR,BJSJ,120) <= '" + endTime + "' ";

        sql += "ORDER BY BJSJ";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        DataTable dataTable = clsGetData.GetTable(sql);

        if (typeid != "0")
        {
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = dataTable.Rows[i];
                String[] ids = row["typecode"].ToString().Split(',');
                if (ids.Count(id => id == typeid) == 0)
                    dataTable.Rows.Remove(row);
            }
        }

        return dataTable;
    }

    [WebMethod]
    public DataTable getAlarmTypeInfoFX()
    {
        String sql = "SELECT [ID]" +
                       ",[PID]" +
                       ",[JQNAME]" +
                       ",[LEVEL]" +
                       ",[CODE] " +
                       ",[COLOR] " +
                       "FROM T_QW_JQLBTREE " +
                       "ORDER BY ID";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        DataTable dataTable = clsGetData.GetTable(sql);

        return dataTable;
    }
    
    [WebMethod]
    public DataTable getBaseSTIDDMG()
    {
        String sql = "SELECT " +
                     "'5'           GPSTYPE," +
                     "STID          GPSSIMCARD, " +
                     "HH            CALLNO, " +
                     "NAME + '(' + HH + ')'         GPSNAME, " +
                     "DEPID         GPSDEPID " +
            "FROM T_QW_BASESTIDMG WHERE Isuse = 1";

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        DataTable dataTable = clsGetData.GetTable(sql);

        return dataTable;
    }

    [WebMethod]
    public String setAlarmNewType(String JQID, String NEWJQLB)
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "Update T_QW_JQ set NEWJQLB = '" + NEWJQLB + "' where JQDH = '" + JQID + "'";

        clsGetData.SetTable(sql);

       return clsGetData.ErrorString;
    }

    [WebMethod]
    public DataTable getElePolice()
    {
        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);

        String sql = "SELECT ID ID," +
                     "DEPID         DEPID," +
                     "CODE          CODE," +
                     "NAME          NAME, " +
                     "X             X, " +
                     "Y             Y, " +
                     "TYPE          TYPE " +
            "FROM T_QW_VIDEOPOINT ";

        return clsGetData.GetTable(sql);
    }

    [WebMethod]
    public DataTable getCallInfo_Today()
    {
        String beginTime = DateTime.Now.ToString("yyyyMMdd");

        ClsGetData clsGetData = new ClsGetData("System.Data.SqlClient", strConn);
        String sql = "SELECT " +
                        "BDMJYID USERID," +
                        "STID  RADIONO " +
                        "FROM T_QW_DM " +
                        "WHERE CONVERT(VARCHAR,DMDate,112) = '" + beginTime + "' " +
                        "ORDER BY DMDate DESC";
        return clsGetData.GetTable(sql);
    }
}
