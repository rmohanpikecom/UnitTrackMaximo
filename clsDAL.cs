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
            string cmd = "SP_DUKE_DYN_WORKORDERS_CREATE";

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

        #region SQL_WorkOrder_GetList
        public static DataSet SQL_WorkOrder_GetList(int Status)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_WORKORDER_GETLIST";

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

        #region Maximo_MaximoUnits_Create
        public static int Maximo_MaximoUnits_Create(string Station_Detail_ID, string Payment_Status, string Work_Function, string Estimate_Quantity, string Compatible_Unit_Quantity, string Completed_Date, string Compatible_Unit_Changed_By, string Compatible_Unit_Changed_Date, string Compatible_Unit_Station, string Parent_Compatible_Unit_Name, string Parent_Compatible_Unit_Description, string Payment_Type, string Field_To_From, string Field_ID_To, string Serial_Number, string Manufacturer, string Vendor_Id, int WorkOrder_id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_MAXIMO_UNITS_CREATE";

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

        #region SQL_WorkOrder_StatusUpdate
        public static int SQL_WorkOrder_StatusUpdate(int WorkOrder_Id)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_WORKORDER_UPDATE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@WorkOrder_Id",SqlDbType.Int)
            };
            commandParameters[0].Value = WorkOrder_Id;

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

        #region SQL_Update_SubTaskDetails
        public static int SQL_Update_SubTaskDetails(string WorkOrder_Id, string Sub_Task_Number, string Sub_Task_Name, string Sub_Task_Id, string Sub_Task_Billable_Flag, string Sub_Task_Crew_Leader, string Sub_Task_Project_Flag)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_SUBTASK_UPDATE";

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

        #region Oracle_NLR_Create
        public static int Oracle_NLR_Create(string PARAM_PROJ_TO_BILL, string PARAM_WONO_TO_BILL, string PARAM_SUBTASK_TO_BILL, string PARAM_UNIT_TO_BILL, string PARAM_QUANTITY, string PARAM_WE_DATE, string ERROR_CODES, string ORACLE_TASK_NO, string ORACLE_SUBTASK, string PARENT_BILLABLE_FLAG, string SUBTASK_BILLABLE_CHARGEABLE_FLAG, string RATE_SCHEDUL_ERROR, string PROJ_TASK_DETAILS, string PROJECT_NUMBER, string PROJECT_NAME, string SUBTASK_PROJECT, string PARENT_WO_NUMBER, string PARENT_WO_NAME, string PARENT_BILLABLE, string PARENT_CHARGEABLE, string TOP_TASK_ID, string SUBTASK_WO_NUMBER, string SUBTASK_WO_NAME, string SUBTASK_WO_BILLABLE, string SUBTASK_WO_CHARGEABLE, string CREW_LEADER, string RATE_SCHEDULE_DETAILS, string EXP_NAME, string RATE_SCHEDULE_NAME, string UNIT_NAME, string RATE, string UNIT_OF_MEASURE, string RATE_START_DATE, string RATE_END_DATE, string FBDILOADER, string EXPENDITUREDATE, string PERSONNAME, string PERSONNUMBER, string HUMANRESOURCEASSIGNMENT, string PROJECTNAME, string PROJECTNUMBER, string TASK_NAME, string TASK_NUMBER, string EXPENDITURETYPE, string EXPENDITUREORGANIZATION, string CONTRACTNUMBER, string FUNDINGSOURCENUMBER, string NONLABORRESOURCE, string NONLABORRESOURCEORGANIZATION, string QUANTITY, string WORKTYPE, string ADDITIONALINFO, string PROJID, string NLRID, string TASKID, string EXISTINGQTYINPPM, string REGION, string LEGALENTITY, string BUSINESSUNIT, string Compatible_Unit_Id)
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

        #region SQL_CU_Data_Update
        public static int SQL_CU_Data_Update(int CompatibleUnitId, string StatusDetails, string MessageDetails)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_CU_DATA_UPDATE";
            SqlParameter[] commandParameters =
            {
                new SqlParameter("@CompatibleUnitId",SqlDbType.Int),
                new SqlParameter("@Status_Details",SqlDbType.NVarChar,100),
                new SqlParameter("@Message_Details",SqlDbType.NVarChar,500)

            };
            commandParameters[0].Value = CompatibleUnitId;
            commandParameters[1].Value = StatusDetails;
            commandParameters[2].Value = MessageDetails;

            int res = SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
            return res;
        }
        #endregion

        #region SQL_Oracle_PPM_GetList
        public static DataSet SQL_Oracle_PPM_GetList()
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_ORACLE_PPM_GETLIST";

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #region SQL_Duke_NLR_DataUpdate
        public static int SQL_Duke_NLR_DataUpdate(string TransactionNumber, string OracleStatus, string OracleMessage, string UnprocessTransactionId, string CompatibleUnitId)
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_ORACLE_NLR_DATA_UPDATE";
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

        #region CPR_Data_Create
        public static int CPR_Data_Create(string Approval_Date, string Invoice_Num, string Invoice_Status, string Invoice_CheckNum, string Invoice_De_Psvoucheramt, string De_Rejectcode, string Rejection_Date, string CPR, string Invoice_Due_Date, string Invoice_Paid_Date, string Status, string HDescription, string Vendor_Invoice_Num, string Week_Ending, string Vendor_Project_ID, string Site, string CPR_Type, string Reactive_Time_Report_ID, string Payment_Type, string Foreman, string Derived_Contract, string First_Approver, string Second_Approver, string Line_of_Business, string Rejection_Remarks, string Requested_By_Name, string Requested_By_Email, string Requested_By_Phone_Num, string CPR_Submit_Date, string CPR_Created_Date, string Total_Cost, string Line, string WO_Task_Num, string Service_Item, string PrLine_Description, string PRLineQuantity, string Order_Unit, string Unit_Cost, string Line_Cost, string GL_Debit_Account, string WO_Num, string Description, string Point_Span, string CU_ID, string CU_Name, string CU_Description, string CU_Service_Item, string Service_Item_Description, string Estimated_Qty, string Asbuilt_Qty, string Work_Function, string Contract, string ShStatus, string ShChangeDate, string ShChangedBy, string ShRejectionCode)

        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_CPR_DATA_CREATE";

            SqlParameter[] commandParameters =
            {
        new SqlParameter("@Approval_Date",SqlDbType.NVarChar,100),
        new SqlParameter("@Invoice_Num ",SqlDbType.NVarChar,100),
        new SqlParameter("@Invoice_Status ",SqlDbType.NVarChar,100),
        new SqlParameter("@Invoice_CheckNum ",SqlDbType.NVarChar,100),
        new SqlParameter("@Invoice_De_Psvoucheramt ",SqlDbType.NVarChar,100),
        new SqlParameter("@De_Rejectcode ",SqlDbType.NVarChar,100),
        new SqlParameter("@Rejection_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@CPR ",SqlDbType.NVarChar,100),
        new SqlParameter("@Invoice_Due_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@Invoice_Paid_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@Status ",SqlDbType.NVarChar,100),
        new SqlParameter("@Description ",SqlDbType.NVarChar,100),
        new SqlParameter("@Vendor_Invoice_Num ",SqlDbType.NVarChar,100),
        new SqlParameter("@Week_Ending ",SqlDbType.NVarChar,100),
        new SqlParameter("@Vendor_Project_ID ",SqlDbType.NVarChar,100),
        new SqlParameter("@Site ",SqlDbType.NVarChar,100),
        new SqlParameter("@CPR_Type ",SqlDbType.NVarChar,100),
        new SqlParameter("@Reactive_Time_Report_ID ",SqlDbType.NVarChar,100),
        new SqlParameter("@Payment_Type ",SqlDbType.NVarChar,100),
        new SqlParameter("@Foreman ",SqlDbType.NVarChar,100),
        new SqlParameter("@Derived_Contract ",SqlDbType.NVarChar,100),
        new SqlParameter("@_1st_Approver ",SqlDbType.NVarChar,100),
        new SqlParameter("@_2nd_Approver ",SqlDbType.NVarChar,100),
        new SqlParameter("@Line_of_Business ",SqlDbType.NVarChar,100),
        new SqlParameter("@Rejection_Remarks ",SqlDbType.NVarChar,100),
        new SqlParameter("@Requested_By_Name ",SqlDbType.NVarChar,100),
        new SqlParameter("@Requested_By_Email ",SqlDbType.NVarChar,100),
        new SqlParameter("@Requested_By_Phone_Num ",SqlDbType.NVarChar,100),
        new SqlParameter("@CPR_Submit_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@CPR_Created_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@Total_Cost ",SqlDbType.NVarChar,100),
        new SqlParameter("@Line ",SqlDbType.NVarChar,100),
        new SqlParameter("@WO_Task_Num ",SqlDbType.NVarChar,100),
        new SqlParameter("@Service_Item ",SqlDbType.NVarChar,100),
        new SqlParameter("@PrLine_Description ",SqlDbType.NVarChar,100),
        new SqlParameter("@PRLineQuantity ",SqlDbType.NVarChar,100),
        new SqlParameter("@Order_Unit ",SqlDbType.NVarChar,100),
        new SqlParameter("@Unit_Cost ",SqlDbType.NVarChar,100),
        new SqlParameter("@Line_Cost ",SqlDbType.NVarChar,100),
        new SqlParameter("@GL_Debit_Account ",SqlDbType.NVarChar,100),
        new SqlParameter("@WO_Num ",SqlDbType.NVarChar,100),
        new SqlParameter("@Description2 ",SqlDbType.NVarChar,100),
        new SqlParameter("@Point_Span ",SqlDbType.NVarChar,100),
        new SqlParameter("@CU_ID ",SqlDbType.NVarChar,100),
        new SqlParameter("@CU_Name ",SqlDbType.NVarChar,100),
        new SqlParameter("@CU_Description ",SqlDbType.NVarChar,100),
        new SqlParameter("@Service_Item3 ",SqlDbType.NVarChar,100),
        new SqlParameter("@Service_Item_Description ",SqlDbType.NVarChar,100),
        new SqlParameter("@Estimated_Qty ",SqlDbType.NVarChar,100),
        new SqlParameter("@Asbuilt_Qty ",SqlDbType.NVarChar,100),
        new SqlParameter("@Work_Function ",SqlDbType.NVarChar,100),                
        //new SqlParameter("@Work_Start_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@Contract ",SqlDbType.NVarChar,100),
        new SqlParameter("@Status_History ",SqlDbType.NVarChar,100),
        new SqlParameter("@Status4 ",SqlDbType.NVarChar,100),
        new SqlParameter("@Change_Date ",SqlDbType.NVarChar,100),
        new SqlParameter("@Changed_By ",SqlDbType.NVarChar,100),
        new SqlParameter("@Rejection_Code ",SqlDbType.NVarChar,100)
    };
            commandParameters[0].Value = Approval_Date;
            commandParameters[1].Value = Invoice_Num;
            commandParameters[2].Value = Invoice_Status;
            commandParameters[3].Value = Invoice_CheckNum;
            commandParameters[4].Value = Invoice_De_Psvoucheramt;
            commandParameters[5].Value = De_Rejectcode;
            commandParameters[6].Value = Rejection_Date;
            commandParameters[7].Value = CPR;
            commandParameters[8].Value = Invoice_Due_Date;
            commandParameters[9].Value = Invoice_Paid_Date;
            commandParameters[10].Value = Status;
            commandParameters[11].Value = HDescription;
            commandParameters[12].Value = Vendor_Invoice_Num;
            commandParameters[13].Value = Week_Ending;
            commandParameters[14].Value = Vendor_Project_ID;
            commandParameters[15].Value = Site;
            commandParameters[16].Value = CPR_Type;
            commandParameters[17].Value = Reactive_Time_Report_ID;
            commandParameters[18].Value = Payment_Type;
            commandParameters[19].Value = Foreman;
            commandParameters[20].Value = Derived_Contract;
            commandParameters[21].Value = First_Approver;
            commandParameters[22].Value = Second_Approver;
            commandParameters[23].Value = Line_of_Business;
            commandParameters[24].Value = Rejection_Remarks;
            commandParameters[25].Value = Requested_By_Name;
            commandParameters[26].Value = Requested_By_Email;
            commandParameters[27].Value = Requested_By_Phone_Num;
            commandParameters[28].Value = CPR_Submit_Date;
            commandParameters[29].Value = CPR_Created_Date;
            commandParameters[30].Value = Total_Cost;
            commandParameters[31].Value = Line;
            commandParameters[32].Value = WO_Task_Num;
            commandParameters[33].Value = Service_Item;
            commandParameters[34].Value = PrLine_Description;
            commandParameters[35].Value = PRLineQuantity;
            commandParameters[36].Value = Order_Unit;
            commandParameters[37].Value = Unit_Cost;
            commandParameters[38].Value = Line_Cost;
            commandParameters[39].Value = GL_Debit_Account;
            commandParameters[40].Value = WO_Num;
            commandParameters[41].Value = Description;
            commandParameters[42].Value = Point_Span;
            commandParameters[43].Value = CU_ID;
            commandParameters[44].Value = CU_Name;
            commandParameters[45].Value = CU_Description;
            commandParameters[46].Value = CU_Service_Item;
            commandParameters[47].Value = Service_Item_Description;
            commandParameters[48].Value = Estimated_Qty;
            commandParameters[49].Value = Asbuilt_Qty;
            commandParameters[50].Value = Work_Function;
            commandParameters[51].Value = Contract;
            commandParameters[52].Value = ShStatus;
            commandParameters[53].Value = ShChangeDate;
            commandParameters[54].Value = ShChangedBy;
            commandParameters[55].Value = ShRejectionCode;


            return SqlHelper.ExecuteNonQuery(dsn, CommandType.StoredProcedure, cmd, commandParameters);
        }
        #endregion

        #region SQL_Oracle_PC_GetList
        public static DataSet SQL_Oracle_PC_GetList()
        {
            string dsn = System.Configuration.ConfigurationManager.AppSettings["DbConnectionString"]!.ToString();
            string cmd = "SP_DUKE_SQL_ORACLE_PC_GETLIST";

            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.StoredProcedure, cmd);
            DataTable dt = ds.Tables[0];
            return ds;
        }
        #endregion

        #endregion

    }
}
