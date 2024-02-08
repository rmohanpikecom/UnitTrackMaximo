using System.Data;
using System.Threading.Tasks;
using Microsoft.ApplicationBlocks.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace UnitTrackMaximo
{
    public class clsDAL
    {
        public static string dynConnectionString = System.Configuration.ConfigurationManager.AppSettings["DynamicsConnection"]!.ToString();

        public static IConfigurationRoot Config()
        {
            var configuration = new ConfigurationBuilder()
               .AddJsonFile("DBQueries.json", optional: false, reloadOnChange: true)
               .AddEnvironmentVariables()
               .Build();
            return configuration;
        }

        #region 

        #region Dynamics_ProjectTask_GetList
        public static DataSet Dynamics_ProjectTask_GetList()
        {
            var configuration = Config();
            string BatchSize = System.Configuration.ConfigurationManager.AppSettings["BatchSize"]!.ToString();
            string dsn = dynConnectionString;
            string cmd = configuration.GetSection("qryProjectTaskGetList").Value!;
            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.Text, cmd);
            return ds;
        }
        #endregion

        #region SQL_ProjectTask_Create
        public static int SQL_ProjectTask_Create(string Parent_Task_Number, string Parent_Task_Id, string Project_Number, string Project_Id, string BusinessUnitName, string ProcessDate)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_DYN_WORKORDERS_CREATE";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@Parent_Task_Number",SqlDbType.NVarChar,100),
                new SqlParameter("@Parent_Task_Id",SqlDbType.NVarChar,100),
                new SqlParameter("@Project_Number",SqlDbType.NVarChar,100),
                new SqlParameter("@Project_Id",SqlDbType.NVarChar,100),
                new SqlParameter("@BusinessUnitName",SqlDbType.NVarChar,100),
                new SqlParameter("@ProcessDate",SqlDbType.Date)
            };
            commandParameters[0].Value = Parent_Task_Number;
            commandParameters[1].Value = Parent_Task_Id;
            commandParameters[2].Value = Project_Number;
            commandParameters[3].Value = Project_Id;
            commandParameters[4].Value = BusinessUnitName;
            commandParameters[5].Value = ProcessDate;


            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region SQL_MaximoUnits_Create
        public static int SQL_MaximoUnits_Create(string Station_Detail_ID, string Payment_Status, string Work_Function, string Estimate_Quantity, string Compatible_Unit_Quantity, string Completed_Date, string Compatible_Unit_Changed_By, string Compatible_Unit_Changed_Date, string Compatible_Unit_Station, string Parent_Compatible_Unit_Name, string Parent_Compatible_Unit_Description, string Payment_Type, string Field_To_From, string Field_ID_To, string Serial_Number, string Manufacturer, string Vendor_Id, int WorkOrder_id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_MAXIMO_COMPATIBLE_UNITS_CREATE";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@Station_Detail_ID ",SqlDbType.NVarChar,50),
                new SqlParameter("@Payment_Status",SqlDbType.NVarChar,50),
                new SqlParameter("@Work_Function",SqlDbType.NVarChar,50),
                new SqlParameter("@Estimate_Quantity",SqlDbType.NVarChar,50),
                new SqlParameter("@Compatible_Unit_Quantity",SqlDbType.NVarChar,50),
                new SqlParameter("@Completed_Date",SqlDbType.DateTime),
                new SqlParameter("@Compatible_Unit_Changed_By",SqlDbType.NVarChar,50),
                new SqlParameter("@Compatible_Unit_Changed_Date",SqlDbType.DateTime),
                new SqlParameter("@Compatible_Unit_Station",SqlDbType.NVarChar,50),
                new SqlParameter("@Parent_Compatible_Unit_Name",SqlDbType.NVarChar,50),
                new SqlParameter("@Parent_Compatible_Unit_Description",SqlDbType.NVarChar,500),
                new SqlParameter("@Payment_Type",SqlDbType.NVarChar,50),
                new SqlParameter("@Field_To_From",SqlDbType.NVarChar,50),
                new SqlParameter("@Field_ID_To",SqlDbType.NVarChar,50),
                new SqlParameter("@Serial_Number",SqlDbType.NVarChar,50),
                new SqlParameter("@Manufacturer",SqlDbType.NVarChar,50),
                new SqlParameter("@Vendor_Id",SqlDbType.NVarChar,50),
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int)
            };
            commandParameters[0].Value = Station_Detail_ID;
            commandParameters[1].Value = Payment_Status;
            commandParameters[2].Value = Work_Function;
            commandParameters[3].Value = Estimate_Quantity;
            commandParameters[4].Value = Compatible_Unit_Quantity;
            commandParameters[5].Value = Completed_Date;
            commandParameters[6].Value = Compatible_Unit_Changed_By;
            commandParameters[7].Value = Compatible_Unit_Changed_Date;
            commandParameters[8].Value = Compatible_Unit_Station;
            commandParameters[9].Value = Parent_Compatible_Unit_Name;
            commandParameters[10].Value = Parent_Compatible_Unit_Description;
            commandParameters[11].Value = Payment_Type;
            commandParameters[12].Value = Field_To_From;
            commandParameters[13].Value = Field_ID_To;
            commandParameters[14].Value = Serial_Number;
            commandParameters[15].Value = Manufacturer;
            commandParameters[16].Value = Vendor_Id;
            commandParameters[17].Value = WorkOrder_id;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region SQL_Update_ServiceItem
        public static int SQL_Update_ServiceItem()
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_UPDATE_SERVICE_ITEM";

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd);
        }
        #endregion



        #region SQL_ProjectTask_GetList
        public static DataSet SQL_ProjectTask_GetList(int Status)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_PROJECT_TASK_GETLIST";

            SqlParameter[] commandParameters =
           {
                new SqlParameter("@Status_Id ",SqlDbType.Int)
            };
            commandParameters[0].Value = Status;

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion


        #region SQL_MaximoUnits_Create
        public static int Update_SubTaskDetails(string WorkOrder_Id, string Sub_Task_Number, string Sub_Task_Name, string Sub_Task_Id, string Sub_Task_Billable_Flag, string Sub_Task_Crew_Leader, string Sub_Task_Project_Flag)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_SUB_TASK_UPDATE";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id ",SqlDbType.NVarChar,100),
                new SqlParameter("@Sub_Task_Number",SqlDbType.NVarChar,100),
                new SqlParameter("@Sub_Task_Name",SqlDbType.NVarChar,100),
                new SqlParameter("@Sub_Task_Id",SqlDbType.NVarChar,100),
                new SqlParameter("@Sub_Task_Billable_Flag",SqlDbType.NVarChar,100),
                new SqlParameter("@Sub_Task_Crew_Leader",SqlDbType.NVarChar,100),
                new SqlParameter("@Sub_Task_Project_Flag",SqlDbType.NVarChar,100)
            };
            commandParameters[0].Value = WorkOrder_Id;
            commandParameters[1].Value = Sub_Task_Number;
            commandParameters[2].Value = Sub_Task_Name;
            commandParameters[3].Value = Sub_Task_Id;
            commandParameters[4].Value = Sub_Task_Billable_Flag;
            commandParameters[5].Value = Sub_Task_Crew_Leader;
            commandParameters[6].Value = Sub_Task_Project_Flag;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion


        #region SQL_ProjectTask_StatusUpdate
        public static int SQL_ProjectTask_StatusUpdate(int WorkOrder_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_PROJECT_TASK_UPDATE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int),
            };
            commandParameters[0].Value = WorkOrder_Id;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion



        #region SQL_NLR_GetList
        public static DataSet SQL_NLR_GetList()
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_NLR_GETLIST";

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion



        #region Oracle_ProjectTask_NLR_Create
        public static int Oracle_ProjectTask_NLR_Create(string PARAM_PROJ_TO_BILL, string PARAM_WONO_TO_BILL, string PARAM_SUBTASK_TO_BILL, string PARAM_UNIT_TO_BILL, string PARAM_QUANTITY, string PARAM_WE_DATE, string ERROR_CODES, string ORACLE_TASK_NO, string ORACLE_SUBTASK, string PARENT_BILLABLE_FLAG, string SUBTASK_BILLABLE_CHARGEABLE_FLAG, string RATE_SCHEDUL_ERROR, string PROJ_TASK_DETAILS, string PROJECT_NUMBER, string PROJECT_NAME, string SUBTASK_PROJECT, string PARENT_WO_NUMBER, string PARENT_WO_NAME, string PARENT_BILLABLE, string PARENT_CHARGEABLE, string TOP_TASK_ID, string SUBTASK_WO_NUMBER, string SUBTASK_WO_NAME, string SUBTASK_WO_BILLABLE, string SUBTASK_WO_CHARGEABLE, string CREW_LEADER, string RATE_SCHEDULE_DETAILS, string EXP_NAME, string RATE_SCHEDULE_NAME, string UNIT_NAME, string RATE, string UNIT_OF_MEASURE, string RATE_START_DATE, string RATE_END_DATE, string FBDILOADER, string EXPENDITUREDATE, string PERSONNAME, string PERSONNUMBER, string HUMANRESOURCEASSIGNMENT, string PROJECTNAME, string PROJECTNUMBER, string TASK_NAME, string TASK_NUMBER, string EXPENDITURETYPE, string EXPENDITUREORGANIZATION, string CONTRACTNUMBER, string FUNDINGSOURCENUMBER, string NONLABORRESOURCE, string NONLABORRESOURCEORGANIZATION, string QUANTITY, string WORKTYPE, string ADDITIONALINFO, string PROJID, string NLRID, string TASKID, string EXISTINGQTYINPPM, string REGION, string LEGALENTITY, string BUSINESSUNIT)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_ORACLE_PROJECT_TASK_NLR_DATA_CREATE";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@PARAM_PROJ_TO_BILL",SqlDbType.NVarChar,100),
                new SqlParameter("@PARAM_WONO_TO_BILL",SqlDbType.NVarChar,100),
                new SqlParameter("@PARAM_SUBTASK_TO_BILL",SqlDbType.NVarChar,100),
                new SqlParameter("@PARAM_UNIT_TO_BILL",SqlDbType.NVarChar,100),
                new SqlParameter("@PARAM_QUANTITY",SqlDbType.NVarChar,100),
                new SqlParameter("@PARAM_WE_DATE",SqlDbType.NVarChar,100),
                new SqlParameter("@ERROR_CODES",SqlDbType.NVarChar,100),
                new SqlParameter("@ORACLE_TASK_NO",SqlDbType.NVarChar,100),
                new SqlParameter("@ORACLE_SUBTASK",SqlDbType.NVarChar,100),
                new SqlParameter("@PARENT_BILLABLE_FLAG",SqlDbType.NVarChar,100),
                new SqlParameter("@SUBTASK_BILLABLE_CHARGEABLE_FLAG",SqlDbType.NVarChar,100),
                new SqlParameter("@RATE_SCHEDUL_ERROR",SqlDbType.NVarChar,100),
                new SqlParameter("@PROJ_TASK_DETAILS",SqlDbType.NVarChar,100),
                new SqlParameter("@PROJECT_NUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@PROJECT_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@SUBTASK_PROJECT",SqlDbType.NVarChar,100),
                new SqlParameter("@PARENT_WO_NUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@PARENT_WO_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@PARENT_BILLABLE",SqlDbType.NVarChar,100),
                new SqlParameter("@PARENT_CHARGEABLE",SqlDbType.NVarChar,100),
                new SqlParameter("@TOP_TASK_ID",SqlDbType.NVarChar,100),
                new SqlParameter("@SUBTASK_WO_NUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@SUBTASK_WO_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@SUBTASK_WO_BILLABLE",SqlDbType.NVarChar,100),
                new SqlParameter("@SUBTASK_WO_CHARGEABLE",SqlDbType.NVarChar,100),
                new SqlParameter("@CREW_LEADER",SqlDbType.NVarChar,100),
                new SqlParameter("@RATE_SCHEDULE_DETAILS",SqlDbType.NVarChar,100),
                new SqlParameter("@EXP_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@RATE_SCHEDULE_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@UNIT_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@RATE",SqlDbType.NVarChar,100),
                new SqlParameter("@UNIT_OF_MEASURE",SqlDbType.NVarChar,100),
                new SqlParameter("@RATE_START_DATE",SqlDbType.NVarChar,100),
                new SqlParameter("@RATE_END_DATE",SqlDbType.NVarChar,100),
                new SqlParameter("@FBDILOADER",SqlDbType.NVarChar,100),
                new SqlParameter("@EXPENDITUREDATE",SqlDbType.NVarChar,100),
                new SqlParameter("@PERSONNAME",SqlDbType.NVarChar,100),
                new SqlParameter("@PERSONNUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@HUMANRESOURCEASSIGNMENT",SqlDbType.NVarChar,100),
                new SqlParameter("@PROJECTNAME",SqlDbType.NVarChar,100),
                new SqlParameter("@PROJECTNUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@TASK_NAME",SqlDbType.NVarChar,100),
                new SqlParameter("@TASK_NUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@EXPENDITURETYPE",SqlDbType.NVarChar,100),
                new SqlParameter("@EXPENDITUREORGANIZATION",SqlDbType.NVarChar,100),
                new SqlParameter("@CONTRACTNUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@FUNDINGSOURCENUMBER",SqlDbType.NVarChar,100),
                new SqlParameter("@NONLABORRESOURCE",SqlDbType.NVarChar,100),
                new SqlParameter("@NONLABORRESOURCEORGANIZATION",SqlDbType.NVarChar,100),
                new SqlParameter("@QUANTITY",SqlDbType.NVarChar,100),
                new SqlParameter("@WORKTYPE",SqlDbType.NVarChar,100),
                new SqlParameter("@ADDITIONALINFO",SqlDbType.NVarChar,100),
                new SqlParameter("@PROJID",SqlDbType.NVarChar,100),
                new SqlParameter("@NLRID",SqlDbType.NVarChar,100),
                new SqlParameter("@TASKID",SqlDbType.NVarChar,100),
                new SqlParameter("@EXISTINGQTYINPPM",SqlDbType.NVarChar,100),
                new SqlParameter("@REGION",SqlDbType.NVarChar,100),
                new SqlParameter("@LEGALENTITY",SqlDbType.NVarChar,100),
                new SqlParameter("@BUSINESSUNIT",SqlDbType.NVarChar,100)

            };
            commandParameters[0].Value = PARAM_PROJ_TO_BILL;
            commandParameters[1].Value = PARAM_WONO_TO_BILL;
            commandParameters[2].Value = PARAM_SUBTASK_TO_BILL;
            commandParameters[3].Value = PARAM_UNIT_TO_BILL;
            commandParameters[4].Value = PARAM_QUANTITY;
            commandParameters[5].Value = PARAM_WE_DATE;
            commandParameters[6].Value = ERROR_CODES;
            commandParameters[7].Value = ORACLE_TASK_NO;
            commandParameters[8].Value = ORACLE_SUBTASK;
            commandParameters[9].Value = PARENT_BILLABLE_FLAG;
            commandParameters[10].Value = SUBTASK_BILLABLE_CHARGEABLE_FLAG;
            commandParameters[11].Value = RATE_SCHEDUL_ERROR;
            commandParameters[12].Value = PROJ_TASK_DETAILS;
            commandParameters[13].Value = PROJECT_NUMBER;
            commandParameters[14].Value = PROJECT_NAME;
            commandParameters[15].Value = SUBTASK_PROJECT;
            commandParameters[16].Value = PARENT_WO_NUMBER;
            commandParameters[17].Value = PARENT_WO_NAME;
            commandParameters[18].Value = PARENT_BILLABLE;
            commandParameters[19].Value = PARENT_CHARGEABLE;
            commandParameters[20].Value = TOP_TASK_ID;
            commandParameters[21].Value = SUBTASK_WO_NUMBER;
            commandParameters[22].Value = SUBTASK_WO_NAME;
            commandParameters[23].Value = SUBTASK_WO_BILLABLE;
            commandParameters[24].Value = SUBTASK_WO_CHARGEABLE;
            commandParameters[25].Value = CREW_LEADER;
            commandParameters[26].Value = RATE_SCHEDULE_DETAILS;
            commandParameters[27].Value = EXP_NAME;
            commandParameters[28].Value = RATE_SCHEDULE_NAME;
            commandParameters[29].Value = UNIT_NAME;
            commandParameters[30].Value = RATE;
            commandParameters[31].Value = UNIT_OF_MEASURE;
            commandParameters[32].Value = RATE_START_DATE;
            commandParameters[33].Value = RATE_END_DATE;
            commandParameters[34].Value = FBDILOADER;
            commandParameters[35].Value = EXPENDITUREDATE;
            commandParameters[36].Value = PERSONNAME;
            commandParameters[37].Value = PERSONNUMBER;
            commandParameters[38].Value = HUMANRESOURCEASSIGNMENT;
            commandParameters[39].Value = PROJECTNAME;
            commandParameters[40].Value = PROJECTNUMBER;
            commandParameters[41].Value = TASK_NAME;
            commandParameters[42].Value = TASK_NUMBER;
            commandParameters[43].Value = EXPENDITURETYPE;
            commandParameters[44].Value = EXPENDITUREORGANIZATION;
            commandParameters[45].Value = CONTRACTNUMBER;
            commandParameters[46].Value = FUNDINGSOURCENUMBER;
            commandParameters[47].Value = NONLABORRESOURCE;
            commandParameters[48].Value = NONLABORRESOURCEORGANIZATION;
            commandParameters[49].Value = QUANTITY;
            commandParameters[50].Value = WORKTYPE;
            commandParameters[51].Value = ADDITIONALINFO;
            commandParameters[52].Value = PROJID;
            commandParameters[53].Value = NLRID;
            commandParameters[54].Value = TASKID;
            commandParameters[55].Value = EXISTINGQTYINPPM;
            commandParameters[56].Value = REGION;
            commandParameters[57].Value = LEGALENTITY;
            commandParameters[58].Value = BUSINESSUNIT;



            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #endregion

    }
}
