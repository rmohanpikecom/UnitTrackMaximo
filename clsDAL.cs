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

        #region Dynamics_WorkOrders_GetList
        public static DataSet Dynamics_WorkOrders_GetList()
        {
            var configuration = Config();
            string BatchSize = System.Configuration.ConfigurationManager.AppSettings["BatchSize"]!.ToString();
            string dsn = dynConnectionString;
            string cmd = configuration.GetSection("qryDynWorkOrderGetList").Value!;
            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.Text, cmd);
            return ds;
        }
        #endregion

        #region Dynamics_WorkOrder_Create
        public static int Dynamics_WorkOrder_Create(string Parent_Task_Number, string Parent_Task_Id, string Project_Number, string Project_Id, string BusinessUnitName, string ProcessDate, int Project_SubTask_Flag)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDERS_CREATE";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@Parent_Task_Number",SqlDbType.NVarChar,100),
                new SqlParameter("@Parent_Task_Id",SqlDbType.NVarChar,100),
                new SqlParameter("@Project_Number",SqlDbType.NVarChar,100),
                new SqlParameter("@Project_Id",SqlDbType.NVarChar,100),
                new SqlParameter("@BusinessUnitName",SqlDbType.NVarChar,100),
                new SqlParameter("@ProcessDate",SqlDbType.Date),
                new SqlParameter("@Project_SubTask_Flag",SqlDbType.Int),
                
            };
            commandParameters[0].Value = Parent_Task_Number;
            commandParameters[1].Value = Parent_Task_Id;
            commandParameters[2].Value = Project_Number;
            commandParameters[3].Value = Project_Id;
            commandParameters[4].Value = BusinessUnitName;
            commandParameters[5].Value = ProcessDate;
            commandParameters[6].Value = Project_SubTask_Flag;


            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region WorkOrder_GetList_Maximo
        public static DataSet WorkOrder_GetList_Maximo(DateTime ProcessDate)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_GETLIST_MAXIMO_DATA";

            SqlParameter[] commandParameters =
           {
                new SqlParameter("@ProcessDate",SqlDbType.Date)
            };
            commandParameters[0].Value = ProcessDate;

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #region WorkOrder_GetList_SubTask
        public static DataSet WorkOrder_GetList_SubTask(DateTime ProcessDate)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_GETLIST_SUBTASK";

            SqlParameter[] commandParameters =
           {
                new SqlParameter("@ProcessDate",SqlDbType.Date)
            };
            commandParameters[0].Value = ProcessDate;

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #region WorkOrder_StatusUpdate
        public static int WorkOrder_StatusUpdate(int WorkOrder_Id, int Status)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_STATUS_UPDATE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int),
                new SqlParameter("@Status",SqlDbType.Int)
            };
            commandParameters[0].Value = WorkOrder_Id;
            commandParameters[1].Value = Status;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region DukeMaximoUnits_Delete
        public static int DukeMaximoUnits_Delete(int WorkOrder_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_MAXIMO_UNITS_DELETE";
            SqlParameter[] commandParameters =
            {
                    new SqlParameter("@WorkOrder_Id",SqlDbType.Int)
                };
            commandParameters[0].Value = WorkOrder_Id;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region DukeMaximoUnits_Create
        public static int DukeMaximoUnits_Create(string ReportDate, string Site, string Vendor, string Vendor_Name, string CPR, string Payment_Type, string Contract, string Status, string CurrentCprStatusDate, string Total_Cost, string Vendor_Invoice_Num, string CPR_Created_Date, string CPR_Submit_Date, string CPR_Type, string Derived_Contract, string WorkStartDate, string Line, string WO_Task_Num, string Service_Item, string PrLine_Description, string PRLineQuantity, string Order_Unit, string Unit_Cost, string Line_Cost, string GL_Debit_Account, int WorkOrder_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_MAXIMO_UNITS_CREATE";
            SqlParameter[] commandParameters =
            {
            new SqlParameter("@REPORT_DATE",SqlDbType.NVarChar,100),
            new SqlParameter("@SITEID ",SqlDbType.NVarChar,100),
            new SqlParameter("@VENDOR ",SqlDbType.NVarChar,100),
            new SqlParameter("@VENDOR_NAME ",SqlDbType.NVarChar,100),
            new SqlParameter("@CPR ",SqlDbType.NVarChar,100),
            new SqlParameter("@PAYMENT_TYPE ",SqlDbType.NVarChar,100),
            new SqlParameter("@CONTRACT_NUM ",SqlDbType.NVarChar,100),
            new SqlParameter("@CURRENT_CPR_STATUS ",SqlDbType.NVarChar,100),
            new SqlParameter("@CURRENT_CPR_STATUS_DATE ",SqlDbType.NVarChar,100),
            new SqlParameter("@CPR_TOTAL_COST ",SqlDbType.NVarChar,100),
            new SqlParameter("@VENDOR_INVOICE ",SqlDbType.NVarChar,100),
            new SqlParameter("@CPR_CREATED_DATE ",SqlDbType.NVarChar,100),
            new SqlParameter("@CPR_SUBMITTED_DATE ",SqlDbType.NVarChar,100),
            new SqlParameter("@CPR_Type ",SqlDbType.NVarChar,100),
            new SqlParameter("@DERIVED_CONTRACT ",SqlDbType.NVarChar,100),
            new SqlParameter("@WORK_START_DATE ",SqlDbType.NVarChar,100),
            new SqlParameter("@LINE ",SqlDbType.NVarChar,100),
            new SqlParameter("@TASK_WO ",SqlDbType.NVarChar,100),
            new SqlParameter("@ITEMNUM ",SqlDbType.NVarChar,100),
            new SqlParameter("@DESCRIPTION ",SqlDbType.NVarChar,100),
            new SqlParameter("@ORDERQTY ",SqlDbType.NVarChar,100),
            new SqlParameter("@ORDERUNIT ",SqlDbType.NVarChar,100),
            new SqlParameter("@UNITCOST ",SqlDbType.NVarChar,100),
            new SqlParameter("@LINECOST ",SqlDbType.NVarChar,100),
            new SqlParameter("@GL_DEBIT_ACCOUNT ",SqlDbType.NVarChar,100),
            new SqlParameter("@WorkOrder_Id ",SqlDbType.NVarChar,100),

        };

            commandParameters[0].Value = ReportDate;
            commandParameters[1].Value = Site;
            commandParameters[2].Value = Vendor;
            commandParameters[3].Value = Vendor_Name;
            commandParameters[4].Value = CPR;
            commandParameters[5].Value = Payment_Type;
            commandParameters[6].Value = Contract;
            commandParameters[7].Value = Status;
            commandParameters[8].Value = CurrentCprStatusDate;
            commandParameters[9].Value = Total_Cost;
            commandParameters[10].Value = Vendor_Invoice_Num;
            commandParameters[11].Value = CPR_Created_Date;
            commandParameters[12].Value = CPR_Submit_Date;
            commandParameters[13].Value = CPR_Type;
            commandParameters[14].Value = Derived_Contract;
            commandParameters[15].Value = WorkStartDate;
            commandParameters[16].Value = Line;
            commandParameters[17].Value = WO_Task_Num;
            commandParameters[18].Value = Service_Item;
            commandParameters[19].Value = PrLine_Description;
            commandParameters[20].Value = PRLineQuantity;
            commandParameters[21].Value = Order_Unit;
            commandParameters[22].Value = Unit_Cost;
            commandParameters[23].Value = Line_Cost;
            commandParameters[24].Value = GL_Debit_Account;
            commandParameters[25].Value = WorkOrder_Id;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region WorkOrder_GetList
        public static DataSet WorkOrder_GetList(int Status, DateTime ProcessDate)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_GETLIST";

            SqlParameter[] commandParameters =
           {
                new SqlParameter("@Status_Id ",SqlDbType.Int),
                new SqlParameter("@ProcessDate",SqlDbType.Date)
            };
            commandParameters[0].Value = Status;
            commandParameters[1].Value = ProcessDate;

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #region WorkOrder_Update_SubTaskDetails
        public static int WorkOrder_Update_SubTaskDetails(int WorkOrder_Id, string Sub_Task_Number, string Sub_Task_Name, string Sub_Task_Id, string Sub_Task_Billable_Flag, string Sub_Task_Crew_Leader, string Sub_Task_Project_Flag)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_SUBTASK_UPDATE";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id ",SqlDbType.Int),
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

        #region WorkOrder_StatusUpdate_Error
        public static int WorkOrder_StatusUpdate_Error(int WorkOrder_Id, string SubTaskErrorMessage, string SubTaskError)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_STATUS_UPDATE_ERROR";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int),
                new SqlParameter("@SubTaskErrorMessage",SqlDbType.NVarChar,100),
                new SqlParameter("@SubTaskError",SqlDbType.NVarChar,500),
            };
            commandParameters[0].Value = WorkOrder_Id;
            commandParameters[1].Value = SubTaskErrorMessage;
            commandParameters[2].Value = SubTaskError;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion



        #region NLR_GetList
        public static DataSet NLR_GetList(string WorkOrder_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_NLR_GETLIST";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int)
            };
            commandParameters[0].Value = Convert.ToInt32(WorkOrder_Id);

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #region NLR_Data_Delete
        public static int NLR_Data_Delete(string Compatible_Unit_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_NLR_DATA_DELETE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@Compatible_Unit_Id",SqlDbType.NVarChar,100)
             };
            commandParameters[0].Value = Compatible_Unit_Id;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region NLR_Create
        public static int NLR_Create(string PARAM_PROJ_TO_BILL, string PARAM_WONO_TO_BILL, string PARAM_SUBTASK_TO_BILL, string PARAM_UNIT_TO_BILL, string PARAM_QUANTITY, string PARAM_WE_DATE, string ERROR_CODES, string ORACLE_TASK_NO, string ORACLE_SUBTASK, string PARENT_BILLABLE_FLAG, string SUBTASK_BILLABLE_CHARGEABLE_FLAG, string RATE_SCHEDUL_ERROR, string PROJ_TASK_DETAILS, string PROJECT_NUMBER, string PROJECT_NAME, string SUBTASK_PROJECT, string PARENT_WO_NUMBER, string PARENT_WO_NAME, string PARENT_BILLABLE, string PARENT_CHARGEABLE, string TOP_TASK_ID, string SUBTASK_WO_NUMBER, string SUBTASK_WO_NAME, string SUBTASK_WO_BILLABLE, string SUBTASK_WO_CHARGEABLE, string CREW_LEADER, string RATE_SCHEDULE_DETAILS, string EXP_NAME, string RATE_SCHEDULE_NAME, string UNIT_NAME, string RATE, string UNIT_OF_MEASURE, string RATE_START_DATE, string RATE_END_DATE, string FBDILOADER, string EXPENDITUREDATE, string PERSONNAME, string PERSONNUMBER, string HUMANRESOURCEASSIGNMENT, string PROJECTNAME, string PROJECTNUMBER, string TASK_NAME, string TASK_NUMBER, string EXPENDITURETYPE, string EXPENDITUREORGANIZATION, string CONTRACTNUMBER, string FUNDINGSOURCENUMBER, string NONLABORRESOURCE, string NONLABORRESOURCEORGANIZATION, string QUANTITY, string WORKTYPE, string ADDITIONALINFO, string PROJID, string NLRID, string TASKID, string EXISTINGQTYINPPM, string REGION, string LEGALENTITY, string BUSINESSUNIT, string Compatible_Unit_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_NLR_DATA_CREATE";

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
                new SqlParameter("@BUSINESSUNIT",SqlDbType.NVarChar,100),
                new SqlParameter("@Compatible_Unit_Id",SqlDbType.NVarChar,100)

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
            commandParameters[59].Value = Compatible_Unit_Id;



            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region NLR_DataUpdate
        public static int NLR_DataUpdate(string TransactionNumber, string OracleStatus, string OracleMessage, string UnprocessTransactionId, string CompatibleUnitId)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_NLR_DATA_UPDATE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@TransactionNumber",SqlDbType.NVarChar,100),
                new SqlParameter("@OracleStatus",SqlDbType.NVarChar,100),
                new SqlParameter("@OracleMessage",SqlDbType.NVarChar,500),
                new SqlParameter("@UnprocessTransactionId",SqlDbType.NVarChar,100),
                new SqlParameter("@CompatibleUnitId",SqlDbType.NVarChar,100),
            };
            commandParameters[0].Value = TransactionNumber;
            commandParameters[1].Value = OracleStatus;
            commandParameters[2].Value = OracleMessage;
            commandParameters[3].Value = UnprocessTransactionId;
            commandParameters[4].Value = CompatibleUnitId;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion
        


        #region WorkOrder_GetList_Oracle
        public static DataSet WorkOrder_GetList_Oracle(DateTime ProcessDate)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_GETLIST_ORACLE";

            SqlParameter[] commandParameters =
           {
                new SqlParameter("@ProcessDate",SqlDbType.Date)
            };
            commandParameters[0].Value = ProcessDate;

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #region Oracle_PPM_GetList
        public static DataSet Oracle_PPM_GetList(int WorkOrderId)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_ORACLE_PPM_GETLIST";

            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id ",SqlDbType.Int)
            };
            commandParameters[0].Value = WorkOrderId;


            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);

            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion


        #region GetMaximoData_Mailing
        public static DataSet GetMaximoData_Mailing(DateTime RecordDate)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_MAXIMO_GET_DATA_MAILING";

            SqlParameter[] commandParameters =
           {
                new SqlParameter("@RECORD_CREATED_DATE",SqlDbType.Date)
            };
            commandParameters[0].Value = RecordDate;

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd, commandParameters);            
            return ds;
        }
        #endregion

        #region WorkOrder_MailFlag_Update
        public static int WorkOrder_MailFlag_Update(int WorkOrder_Id, int MailFlag)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_WORKORDER_MAILING_FLAG_UPDATE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int),
                new SqlParameter("@Mail_Flag",SqlDbType.Int)
            };
            commandParameters[0].Value = WorkOrder_Id;
            commandParameters[1].Value = MailFlag;

            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion
        #endregion

    }
}
