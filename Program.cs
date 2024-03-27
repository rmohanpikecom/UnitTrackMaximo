using Newtonsoft.Json;
using System.Data;
using System.Text;
using System.Web;
using RestSharp;
using Dynamics_Oracle_UnitBilling;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using System.Dynamic;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Threading.Tasks;

namespace UnitTrackMaximo
{
    class Program
    {
        static string AppName = "UnitTrackMaximo";
        string? filePath = System.Configuration.ConfigurationManager.AppSettings["FilePath"]!.ToString();
        string? FolderName = System.Configuration.ConfigurationManager.AppSettings["FolderName"]!.ToString();

        //Application Static Details
        static string Oracle_Url = System.Configuration.ConfigurationManager.AppSettings["Oracle_Url"]!.ToString();
        static string UnitBilling_Sub_Url = System.Configuration.ConfigurationManager.AppSettings["UnitBilling_Sub_Url"]!.ToString();
        static string PPM_SubUrl = System.Configuration.ConfigurationManager.AppSettings["PPM_Sub_Url"]!.ToString();
        static string PPM_EssJob_Sub_Url = System.Configuration.ConfigurationManager.AppSettings["PPM_EssJob_Sub_Url"]!.ToString();

        static string UserName = System.Configuration.ConfigurationManager.AppSettings["ServiceUserId"]!.ToString();
        static string Password = clsTools.Decrypt(System.Configuration.ConfigurationManager.AppSettings["ServicePassword"]!.ToString(), true);
        static string ServiceTimeout = System.Configuration.ConfigurationManager.AppSettings["ServiceTimeout"]!.ToString();

        //Dynamics URL's
        static readonly string Dynamics_Url = System.Configuration.ConfigurationManager.AppSettings["Dynamics_Url"]!.ToString();


        #region Main
        static void Main(string[] args)
        {
            try
            {
                string? filePath = System.Configuration.ConfigurationManager.AppSettings["FilePath"]!.ToString();
                string? FolderName = System.Configuration.ConfigurationManager.AppSettings["FolderName"]!.ToString();

                FileStream? File_FStream = null;
                System.IO.Directory.CreateDirectory(filePath + FolderName);
                string? File_FName = filePath + FolderName + "\\UnitBilling_Log_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";

                // Create a FileStream with mode CreateNew  
                File_FStream = new FileStream(File_FName, FileMode.OpenOrCreate);
                // Create a StreamWriter from FileStream  
                StreamWriter Writer = new StreamWriter(File_FStream, Encoding.UTF8);
                Writer.AutoFlush = true;

                Writer.WriteLine("DynamicsPikeService - " + AppName + " - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - Started");

                //Get Workorder Details from Dynamics

                DateTime dt= DateTime.Now;
                //dt = Convert.ToDateTime("03/18/2024");

                GetWorkorders_DYN(Writer, dt);

                GetDukeMaximoUnits_List(Writer, dt);

                GetSubTaskNumber_Oracle(Writer, dt);

                GetNLRData_Oracle(Writer, dt);

                GetOracle_WorkORderList(Writer, dt);

                //Writer.WriteLine("DynamicsPikeService - UnitBilling CREATE Started :" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                Writer.Close();

                Console.WriteLine("DynamicsPikeService - " + AppName + " - Completed");
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.Message.ToString());
            }
            //return 0;

        }
        #endregion

        #region GetWorkorders_DYN
        public static void GetWorkorders_DYN(StreamWriter writer , DateTime dt)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetWorkorders_DYN - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetWorkorders_DYN - Started");

                DataSet dsProjectTask = clsDAL.Dynamics_WorkOrders_GetList();
                if (dsProjectTask.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < dsProjectTask.Tables[0].Rows.Count; i++)
                    {
                        try
                        {
                            Console.WriteLine("GetWorkorders_DYN - Processing - " +(i+1).ToString() +" out of " + dsProjectTask.Tables[0].Rows.Count.ToString());
                            writer.WriteLine("GetWorkorders_DYN - Processing - " + (i + 1).ToString() + " out of " + dsProjectTask.Tables[0].Rows.Count.ToString());

                            string Parent_Task_Number = "";
                            string Parent_Task_Id = "";
                            string Project_Number = "";
                            string Project_Id = "";
                            string BusinessUnitName = "";
                            string ProcessDate = "";
                            int Project_SubTask_Flag = 0;

                            Parent_Task_Number = dsProjectTask.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;
                            Parent_Task_Id = dsProjectTask.Tables[0].Rows[i]["Parent_Task_Id"].ToString()!;
                            Project_Number = dsProjectTask.Tables[0].Rows[i]["Project_Number"].ToString()!;
                            Project_Id = dsProjectTask.Tables[0].Rows[i]["Project_Id"].ToString()!;
                            BusinessUnitName = dsProjectTask.Tables[0].Rows[i]["BusinessUnitName"].ToString()!;
                            ProcessDate = dt.ToString("yyyy-MM-dd");
                            Project_SubTask_Flag = Convert.ToInt32(dsProjectTask.Tables[0].Rows[i]["Project_SubTask_Flag"].ToString()!);

                            int res = clsDAL.Dynamics_WorkOrder_Create(Parent_Task_Number, Parent_Task_Id, Project_Number, Project_Id, BusinessUnitName, ProcessDate, Project_SubTask_Flag);
                            
                            if (res > 0)
                            {
                                //update workorder status in dynamics
                                UpdateWorkorderStatus(Parent_Task_Number);
                            }
                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine(exp.Message.ToString());
                            writer.WriteLine(exp.Message.ToString());                            
                        }
                    }

                    Console.WriteLine("GetWorkorders_DYN - Processing Completed");
                    writer.WriteLine("GetWorkorders_DYN - Processing Completed");
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("GetWorkorders_DYN - Failure" + exp.Message.ToString());
                writer.WriteLine("GetWorkorders_DYN - Failure" + exp.Message.ToString());
            }

        }
        #endregion

        #region GetDukeMaximoUnits_List
        public static void GetDukeMaximoUnits_List(StreamWriter writer, DateTime dt)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeMaximoUnits - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeMaximoUnits - Started");

                DataSet ds = clsDAL.WorkOrder_GetList_Maximo(dt);

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetDukeMaximoUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeMaximoUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        int WorkOrder_Id = 0;
                        string WorkOrder_Number = ""; 

                        WorkOrder_Id = Convert.ToInt32(ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString());
                        WorkOrder_Number = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;

                        GetDukeMaximoUnits_Details(WorkOrder_Id, WorkOrder_Number, writer);

                    }
                    catch (Exception exp)
                    {
                        Console.WriteLine(exp.Message.ToString());
                        writer.WriteLine(exp.Message.ToString());
                    }
                }

                Console.WriteLine("GetDukeMaximoUnits - Processing Completed");
                writer.WriteLine("GetDukeMaximoUnits - Processing Completed");
            }
            catch (Exception exp)
            {
                Console.WriteLine("GetDukeMaximoUnits - Failure" + exp.Message.ToString());
                writer.WriteLine("GetDukeMaximoUnits - Failure" + exp.Message.ToString());
            }
        }
        #endregion

        #region GetDukeMaximoUnits_Details
        public static string GetDukeMaximoUnits_Details(int WorkOrder_Id, string WorkOrder_Number, StreamWriter writer)
        {
            string token = GetToken();
            string result = "";
            dynamic dyArray = "";


            string ServiceUrl = System.Configuration.ConfigurationManager.AppSettings["OauthUrl"]!.ToString();


            try
            {
                if (token != null)
                {
                    var options = new RestClientOptions(ServiceUrl)
                    {
                        MaxTimeout = -1,
                    };
                    var client = new RestClient(options);
                    var request = new RestRequest("/wmes-external/workorders/tasks/" + WorkOrder_Number + "/cprs", Method.Get);
                    request.AddHeader("Authorization", "Bearer " + token);
                    RestResponse response = client.Execute(request);
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                        dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!.ToString())!;
                    else
                    {
                        clsDAL.WorkOrder_StatusUpdate(WorkOrder_Id, 9);
                        return result;

                    }

                    result = dyArray.ToString();

                    string ReportDate = "";
                    string Site = "";
                    string Vendor = "";
                    string Vendor_Name = "";
                    string CPR = "";
                    string Payment_Type = "";
                    string Contract = "";
                    string Status = "";
                    string CurrentCprStatusDate = "";
                    string Total_Cost = "";
                    string Vendor_Invoice_Num = "";
                    string CPR_Submit_Date = "";
                    string CPR_Created_Date = "";

                    string CPR_Type = "";
                    string Derived_Contract = "";
                    string GL_Debit_Account = "";
                    string WorkStartDate = "";


                    //CPR Lines
                    string Line = "";
                    string WO_Task_Num = "";
                    string Service_Item = "";
                    string PrLine_Description = "";
                    string PRLineQuantity = "";
                    string Order_Unit = "";
                    string Unit_Cost = "";
                    string Line_Cost = "";

                    DateTime dt = DateTime.Now;
                    ReportDate = dt.ToString("yyyy-MM-dd");
                    CurrentCprStatusDate = dt.ToString("yyyy-MM-dd");


                    int Count = 1;
                    int RecordCount = dyArray.Count;

                    if (dyArray.Count != null)
                    {
                        clsDAL.DukeMaximoUnits_Delete(WorkOrder_Id);
                        for (int i = 0; i < dyArray.Count; i++)
                        {
                            for (int j = 0; j < dyArray[i].CPR_Lines.Count; j++)
                            {
                                try
                                {
                                    Console.WriteLine("GetCPR_Details - Details - Processing - " + (j + 1).ToString() + " out of " + dyArray[i].CPR_Lines.Count.ToString());
                                    writer.WriteLine("GetCPR_Details - Details - Processing - " + (j + 1).ToString() + " out of " + dyArray[i].CPR_Lines.Count.ToString());

                                    if (dyArray[i].Site != null)
                                        Site = dyArray[i].Site.ToString();

                                    if (dyArray[i].CPR != null)
                                        CPR = dyArray[i].CPR.ToString();

                                    if (dyArray[i].Payment_Type != null)
                                        Payment_Type = dyArray[i].Payment_Type.ToString();

                                    if (dyArray[i].Contract != null)
                                        WorkStartDate = dyArray[i].Work_Start_Date.ToString();

                                    if (dyArray[i].Status != null)
                                        Status = dyArray[i].Status.ToString();

                                    if (dyArray[i].Total_Cost != null)
                                        Total_Cost = dyArray[i].Total_Cost.ToString();

                                    if (dyArray[i].Vendor_Invoice_Num != null)
                                        Vendor_Invoice_Num = dyArray[i].Vendor_Invoice_Num.ToString();

                                    if (dyArray[i].CPR_Created_Date != null)
                                        CPR_Created_Date = dyArray[i].CPR_Created_Date.ToString();

                                    if (dyArray[i].CPR_Submit_Date != null)
                                        CPR_Submit_Date = dyArray[i].CPR_Submit_Date.ToString();

                                    if (dyArray[i].CPR_Type != null)
                                        CPR_Type = dyArray[i].CPR_Type.ToString();

                                    if (dyArray[i].Derived_Contract != null)
                                        Derived_Contract = dyArray[i].Derived_Contract.ToString();

                                    if (dyArray[i].Work_Start_Date != null)
                                        WorkStartDate = dyArray[i].Work_Start_Date.ToString();


                                    //CPR Lines
                                    if (dyArray[i].CPR_Lines[j].Line != null)
                                        Line = dyArray[i].CPR_Lines[j].Line.ToString();

                                    if (dyArray[i].CPR_Lines[j].WO_Task_Num != null)
                                        WO_Task_Num = dyArray[i].CPR_Lines[j].WO_Task_Num.ToString();

                                    if (dyArray[i].CPR_Lines[j].Service_Item != null)
                                        Service_Item = dyArray[i].CPR_Lines[j].Service_Item.ToString();

                                    if (dyArray[i].CPR_Lines[j].PrLine_Description != null)
                                        PrLine_Description = dyArray[i].CPR_Lines[j].PrLine_Description.ToString();

                                    if (dyArray[i].CPR_Lines[j].PRLineQuantity != null)
                                        PRLineQuantity = dyArray[i].CPR_Lines[j].PRLineQuantity.ToString();

                                    if (dyArray[i].CPR_Lines[j].Order_Unit != null)
                                        Order_Unit = dyArray[i].CPR_Lines[j].Order_Unit.ToString();

                                    if (dyArray[i].CPR_Lines[j].Unit_Cost != null)
                                        Unit_Cost = dyArray[i].CPR_Lines[j].Unit_Cost.ToString();

                                    if (dyArray[i].CPR_Lines[j].Line_Cost != null)
                                        Line_Cost = dyArray[i].CPR_Lines[j].Line_Cost.ToString();

                                    if (dyArray[i].CPR_Lines[j].GL_Debit_Account != null)
                                        GL_Debit_Account = dyArray[i].CPR_Lines[i].GL_Debit_Account.ToString();

                                        //CPR Details

                                    int res = clsDAL.DukeMaximoUnits_Create(ReportDate, Site, Vendor, Vendor_Name, CPR, Payment_Type, Contract, Status, CurrentCprStatusDate, Total_Cost, Vendor_Invoice_Num, CPR_Created_Date, CPR_Submit_Date, CPR_Type, Derived_Contract, WorkStartDate, Line, WO_Task_Num, Service_Item, PrLine_Description, PRLineQuantity, Order_Unit, Unit_Cost, Line_Cost, GL_Debit_Account, WorkOrder_Id);

                                    if (Count == dyArray[i].CPR_Lines.Count)
                                    {
                                        clsDAL.WorkOrder_StatusUpdate(WorkOrder_Id, 2);
                                    }

                                    Count++;
                                }
                                catch (Exception exp)
                                {
                                    writer.WriteLine("UnitTrack Maximo GetCPR_Details Failure");
                                    writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                                    Console.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                                    writer.WriteLine("=========================================================================");
                                }

                            }
                        }
                    }


                }

            }
            catch (Exception exp)
            {

                writer.WriteLine("UnitTrack Maximo GetCPR_Details Failure");
                writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                Console.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                writer.WriteLine("=========================================================================");
            }
            return result;
        }
        #endregion

        #region GetSubTaskNumber_Oracle
        public static void GetSubTaskNumber_Oracle(StreamWriter writer, DateTime dt)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetSubTaskNumber_Oracle - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetSubTaskNumber_Oracle - Started");

                DataSet ds = clsDAL.WorkOrder_GetList_SubTask(dt);
                int WorkOrder_Id = 0;

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetWorkorder SubTaskNumber - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetWorkorder SubTaskNumber - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        WorkOrder_Id = Convert.ToInt32(ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString())!;
                        string ProjectNumber = ds.Tables[0].Rows[i]["Project_Number"].ToString()!;
                        string TaskNumber = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;
                        int Project_SubTask_Flag = Convert.ToInt32(ds.Tables[0].Rows[i]["Project_SubTask_Flag"].ToString()!)+1;

                        string OracleWrapperUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperUrl"]!.ToString();
                        string OracleWrapperSubUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperSubUrl"]!.ToString();
                        string OracleWrapperEnvironment = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperEnvironment"]!.ToString();

                        var configuration = new ConfigurationBuilder()
                            .AddJsonFile("DBQueries.json", optional: false, reloadOnChange: true)
                            .AddEnvironmentVariables()
                            .Build();
                        string cmd = configuration.GetSection("qryOracleSubTaskData").Value!.Replace("@ProjectNumber", "'" + ProjectNumber + "'").Replace("@TaskNumber", "'" + TaskNumber + "'").Replace("@Project_SubTask_Flag", "'" + Project_SubTask_Flag + "'");

                        var options = new RestClientOptions(OracleWrapperUrl)
                        {
                            MaxTimeout = -1,
                        };
                        var client = new RestClient(options);
                        var request = new RestRequest(OracleWrapperSubUrl, Method.Post);
                        request.AddHeader("Content-Type", "application/json");
                        var body = "{ "
                                                        + "\"env\":\"" + OracleWrapperEnvironment + "\","
                                                        + "\"format\":\"xml\","
                                                        + "\"query\":\"" + cmd + "\"}";



                        request.AddStringBody(body, DataFormat.Json);
                        RestResponse response = client.Execute(request);

                        string data = response.Content!.ToString();
                        string replaceWith = "";

                        //string xml = "\"\r\n<?xml version=\\\"1.0\\\" encoding=\\\"UTF-8\\\"?>\\n\r\n<!--Generated by Oracle Analytics Publisher -Dataengine, datamodel:_Custom_EDW_Synapse_Live_Sync_Live_Data_Model_FSCM_xdm -->\\n\r\n<ROWSET>\\n\r\n    <ROW>\\n\r\n        <PARAM_PROJ_TO_BILL>23-01219-000</PARAM_PROJ_TO_BILL>\\n\r\n        <PARAM_WONO_TO_BILL>10012889</PARAM_WONO_TO_BILL>\\n\r\n        <PARAM_SUBTASK_TO_BILL>44107-10012889</PARAM_SUBTASK_TO_BILL>\\n\r\n        <PARAM_UNIT_TO_BILL>OAA01-MAN HOUR RATE PER MAN OH</PARAM_UNIT_TO_BILL>\\n\r\n        <PARAM_QUANTITY>10</PARAM_QUANTITY>\\n\r\n        <PARAM_WE_DATE>2023-01-22</PARAM_WE_DATE>\\n\r\n        <ERROR_CODES>ERROR CODES-&gt;</ERROR_CODES>\\n\r\n        <ORACLE_TASK_NO></ORACLE_TASK_NO>\\n\r\n        <ORACLE_SUBTASK>Project Requires Billable Subtask Name</ORACLE_SUBTASK>\\n\r\n        <PARENT_BILLABLE_FLAG>Parent Task Must be Billable=Y</PARENT_BILLABLE_FLAG>\\n\r\n        <SUBTASK_BILLABLE_CHARGEABLE_FLAG>Subtask Not Found</SUBTASK_BILLABLE_CHARGEABLE_FLAG>\\n\r\n        <RATE_SCHEDUL_ERROR></RATE_SCHEDUL_ERROR>\\n\r\n        <PROJ_TASK_DETAILS>PROJ_TASK DETAILS-&gt;</PROJ_TASK_DETAILS>\\n\r\n        <PROJECT_NUMBER>23-01219-000</PROJECT_NUMBER>\\n\r\n        <PROJECT_NAME>23-01219-000 OHD FL BOCA RATON</PROJECT_NAME>\\n\r\n        <SUBTASK_PROJECT>Y</SUBTASK_PROJECT>\\n\r\n        <PARENT_WO_NUMBER>10012889</PARENT_WO_NUMBER>\\n\r\n        <PARENT_WO_NAME>10012889</PARENT_WO_NAME>\\n\r\n        <PARENT_BILLABLE>N</PARENT_BILLABLE>\\n\r\n        <PARENT_CHARGEABLE>N</PARENT_CHARGEABLE>\\n\r\n        <TOP_TASK_ID></TOP_TASK_ID>\\n\r\n        <SUBTASK_WO_NUMBER></SUBTASK_WO_NUMBER>\\n\r\n        <SUBTASK_WO_NAME></SUBTASK_WO_NAME>\\n\r\n        <SUBTASK_WO_BILLABLE></SUBTASK_WO_BILLABLE>\\n\r\n        <SUBTASK_WO_CHARGEABLE></SUBTASK_WO_CHARGEABLE>\\n\r\n        <CREW_LEADER></CREW_LEADER>\\n\r\n        <RATE_SCHEDULE_DETAILS>RATE SCHEDULE DETAIL-&gt;</RATE_SCHEDULE_DETAILS>\\n\r\n        <EXP_NAME>Unit Production</EXP_NAME>\\n\r\n        <RATE_SCHEDULE_NAME>2023-FPL</RATE_SCHEDULE_NAME>\\n\r\n        <UNIT_NAME>50|OAA01-MAN HOUR RATE PER MAN OH|N|INSTALL|108|HOUR</UNIT_NAME>\\n\r\n        <RATE>135.02</RATE>\\n\r\n        <UNIT_OF_MEASURE>HOURS</UNIT_OF_MEASURE>\\n\r\n        <RATE_START_DATE>2023-01-02</RATE_START_DATE>\\n\r\n        <RATE_END_DATE></RATE_END_DATE>\\n\r\n        <FBDILOADER>FBDI</FBDILOADER>\\n\r\n        <EXPENDITUREDATE>2023-01-22</EXPENDITUREDATE>\\n\r\n        <PERSONNAME></PERSONNAME>\\n\r\n        <PERSONNUMBER></PERSONNUMBER>\\n\r\n        <HUMANRESOURCEASSIGNMENT></HUMANRESOURCEASSIGNMENT>\\n\r\n        <PROJECTNAME>23-01219-000 OHD FL BOCA RATON</PROJECTNAME>\\n\r\n        <PROJECTNUMBER>23-01219-000</PROJECTNUMBER>\\n\r\n        <TASK_NAME></TASK_NAME>\\n\r\n        <TASK_NUMBER></TASK_NUMBER>\\n\r\n        <EXPENDITURETYPE>Unit Production</EXPENDITURETYPE>\\n\r\n        <EXPENDITUREORGANIZATION>Florida</EXPENDITUREORGANIZATION>\\n\r\n        <CONTRACTNUMBER></CONTRACTNUMBER>\\n\r\n        <FUNDINGSOURCENUMBER></FUNDINGSOURCENUMBER>\\n\r\n        <NONLABORRESOURCE>50|OAA01-MAN HOUR RATE PER MAN OH|N|INSTALL|108|HOUR</NONLABORRESOURCE>\\n\r\n        <NONLABORRESOURCEORGANIZATION>PIKE Project Unit Org</NONLABORRESOURCEORGANIZATION>\\n\r\n        <QUANTITY>10</QUANTITY>\\n\r\n        <WORKTYPE></WORKTYPE>\\n\r\n        <ADDITIONALINFO>Additional Info</ADDITIONALINFO>\\n\r\n        <PROJID>300002425947378</PROJID>\\n\r\n        <NLRID>300000013044500</NLRID>\\n\r\n        <TASKID>100007676556772</TASKID>\\n\r\n        <EXISTINGQTYINPPM></EXISTINGQTYINPPM>\\n\r\n        <REGION>Florida</REGION>\\n\r\n        <LEGALENTITY>Pike Electric, LLC</LEGALENTITY>\\n\r\n        <BUSINESSUNIT>Pike Business Unit</BUSINESSUNIT>\\n\r\n    </ROW>\\n\r\n</ROWSET>\\n\"";

                        data = Regex.Replace(data, "^\"|\"$", "");
                        data = data.Replace(System.Environment.NewLine, replaceWith);
                        data = data.Replace("\r\n", replaceWith).Replace("\\n", replaceWith).Replace("\r", replaceWith).Replace("version=\\\"1.0\\\"", "version=\"1.0\"").Replace("encoding=\\\"UTF-8\\\"", "encoding=\"UTF-8\"");




                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(data);

                        var json = JsonConvert.SerializeXmlNode(doc, (Newtonsoft.Json.Formatting)System.Xml.Formatting.None, true);
                        json = json.Replace("\"?xml\":{\"@version\":\"1.0\",\"@encoding\":\"UTF-8\"}{\"ROW\":", "");
                        json = json.Trim().TrimEnd('}');
                        json = json + "}";
                        dynamic dyArray = JObject.Parse(json);
                        var dynamicObject = JsonConvert.DeserializeObject<dynamic>(json)!;


                        string Project_Number = dynamicObject.PROJECT_NUMBER.ToString();
                        string Sub_Task_Project_Flag = dynamicObject.SUBTASK_PROJECT_FLAG.ToString();
                        string Task_WBS_Level = dynamicObject.WBS_LEVEL.ToString();
                        string Parent_Task_Number = dynamicObject.PARENT_TASK_NO.ToString();
                        string Sub_Task_Number = dynamicObject.TASK_NUMBER.ToString();
                        string Sub_Task_Name = dynamicObject.TASK_NAME.ToString();
                        string Sub_Task_Id = dynamicObject.TASK_ID.ToString();
                        string Sub_Task_Billable_Flag = dynamicObject.CHARGEABLE_FLAG.ToString();
                        string Sub_Task_Crew_Leader = dynamicObject.CREW_LEADER.ToString();

                        clsDAL.WorkOrder_Update_SubTaskDetails(WorkOrder_Id, Sub_Task_Number, Sub_Task_Name, Sub_Task_Id, Sub_Task_Billable_Flag, Sub_Task_Crew_Leader, Sub_Task_Project_Flag);
                        clsDAL.WorkOrder_StatusUpdate(WorkOrder_Id, 3);

                    }
                    catch (Exception exp)
                    {
                        clsDAL.WorkOrder_StatusUpdate_Error(WorkOrder_Id,"SubTask Failure",exp.Message.ToString());
                        writer.WriteLine("UnitTrack Maximo GetSubTaskNumber_Oracle Failure");
                        Console.WriteLine("GetSubTaskNumber_Oracle - Failure" + exp.Message.ToString());
                        writer.WriteLine("GetSubTaskNumber_Oracle - Failure" + exp.Message.ToString());
                        writer.WriteLine("=========================================================================");
                    }
                }

            }
            catch (Exception exp)
            {
                writer.WriteLine("UnitTrack Maximo GetSubTaskNumber_Oracle Failure");
                Console.WriteLine("GetSubTaskNumber_Oracle - Failure" + exp.Message.ToString());
                writer.WriteLine("GetSubTaskNumber_Oracle - Failure" + exp.Message.ToString());
                writer.WriteLine("=========================================================================");
            }
        }
        #endregion

        #region GetNLRData_Oracle
        public static void GetNLRData_Oracle(StreamWriter writer, DateTime dt)
        {
            writer.WriteLine("DynamicsPikeService - " + AppName + " - GetNLRDatatoOracle - Started");
            Console.WriteLine("DynamicsPikeService - " + AppName + " - GetNLRDatatoOracle - Started");

            DataSet ds = clsDAL.WorkOrder_GetList(3, dt);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Console.WriteLine("Get NLR Work Orders - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                writer.WriteLine("Get NLR Work Orders - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                string WorkOrder_Id = ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString()!;

                GetNLRData_Oracle_Details(writer, WorkOrder_Id);

            }
        }
        #endregion

        #region GetNLRData_Oracle_Details
        public static void GetNLRData_Oracle_Details(StreamWriter writer, string WorkOrder_Id)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetNLRData_Oracle_Details - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetNLRData_Oracle_Details - Started");


                DataSet ds = clsDAL.NLR_GetList(WorkOrder_Id);
                string Compatible_Unit_Id = "";
                int Count = 1;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string oracle_status_details = "";
                    string oracle_message_details = "";
                    try
                    {
                        Console.WriteLine("GetNLRData_Oracle_Details - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetNLRData_Oracle_Details - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        Compatible_Unit_Id = ds.Tables[0].Rows[i]["Compatible_Unit_Id"].ToString()!;
                        string @Project = ds.Tables[0].Rows[i]["Project_Number"].ToString()!;
                        string @PARENT_TASKNO = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;
                        string @SUBTASK = ds.Tables[0].Rows[i]["Sub_Task_Number"].ToString()!;
                        string @SERVICE_ITEM = ds.Tables[0].Rows[i]["Service_Item"].ToString()!;
                        string @Quantity = ds.Tables[0].Rows[i]["Quanity"].ToString()!;
                        string @EffectiveDate = Convert.ToDateTime(ds.Tables[0].Rows[i]["EffectiveDate"]).ToString("yyyy-MM-dd")!;

                        var configuration = new ConfigurationBuilder()
                           .AddJsonFile("DBQueries.json", optional: false, reloadOnChange: true)
                           .AddEnvironmentVariables()
                           .Build();
                        string cmd = configuration.GetSection("qryOracleNLRData").Value!.Replace("@Project", "'" + @Project + "'").Replace("@PARENT_TASKNO", "'" + @PARENT_TASKNO + "'").Replace("@SUBTASK", "'" + @SUBTASK + "'").Replace("@SERVICE_ITEM", "'" + @SERVICE_ITEM + "'").Replace("@Quantity", "'" + @Quantity + "'").Replace("@EffectiveDate", "'" + @EffectiveDate + "'");


                        string OracleWrapperUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperUrl"]!.ToString();
                        string OracleWrapperSubUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperSubUrl"]!.ToString();
                        string OracleWrapperEnvironment = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperEnvironment"]!.ToString();

                        var options = new RestClientOptions(OracleWrapperUrl)
                        {
                            MaxTimeout = -1,
                        };
                        var client = new RestClient(options);
                        var request = new RestRequest(OracleWrapperSubUrl, Method.Post);
                        request.AddHeader("Content-Type", "application/json");
                        var body = "{ "
                                                      + "\"env\":\"" + OracleWrapperEnvironment + "\","
                                                      + "\"format\":\"xml\","
                                                      + "\"query\":\"" + cmd + "\"}";



                        request.AddStringBody(body, DataFormat.Json);
                        RestResponse response = client.Execute(request);

                        string data = response.Content!.ToString();
                        string replaceWith = "";

                        //string xml = "\"\r\n<?xml version=\\\"1.0\\\" encoding=\\\"UTF-8\\\"?>\\n\r\n<!--Generated by Oracle Analytics Publisher -Dataengine, datamodel:_Custom_EDW_Synapse_Live_Sync_Live_Data_Model_FSCM_xdm -->\\n\r\n<ROWSET>\\n\r\n    <ROW>\\n\r\n        <PARAM_PROJ_TO_BILL>23-01219-000</PARAM_PROJ_TO_BILL>\\n\r\n        <PARAM_WONO_TO_BILL>10012889</PARAM_WONO_TO_BILL>\\n\r\n        <PARAM_SUBTASK_TO_BILL>44107-10012889</PARAM_SUBTASK_TO_BILL>\\n\r\n        <PARAM_UNIT_TO_BILL>OAA01-MAN HOUR RATE PER MAN OH</PARAM_UNIT_TO_BILL>\\n\r\n        <PARAM_QUANTITY>10</PARAM_QUANTITY>\\n\r\n        <PARAM_WE_DATE>2023-01-22</PARAM_WE_DATE>\\n\r\n        <ERROR_CODES>ERROR CODES-&gt;</ERROR_CODES>\\n\r\n        <ORACLE_TASK_NO></ORACLE_TASK_NO>\\n\r\n        <ORACLE_SUBTASK>Project Requires Billable Subtask Name</ORACLE_SUBTASK>\\n\r\n        <PARENT_BILLABLE_FLAG>Parent Task Must be Billable=Y</PARENT_BILLABLE_FLAG>\\n\r\n        <SUBTASK_BILLABLE_CHARGEABLE_FLAG>Subtask Not Found</SUBTASK_BILLABLE_CHARGEABLE_FLAG>\\n\r\n        <RATE_SCHEDUL_ERROR></RATE_SCHEDUL_ERROR>\\n\r\n        <PROJ_TASK_DETAILS>PROJ_TASK DETAILS-&gt;</PROJ_TASK_DETAILS>\\n\r\n        <PROJECT_NUMBER>23-01219-000</PROJECT_NUMBER>\\n\r\n        <PROJECT_NAME>23-01219-000 OHD FL BOCA RATON</PROJECT_NAME>\\n\r\n        <SUBTASK_PROJECT>Y</SUBTASK_PROJECT>\\n\r\n        <PARENT_WO_NUMBER>10012889</PARENT_WO_NUMBER>\\n\r\n        <PARENT_WO_NAME>10012889</PARENT_WO_NAME>\\n\r\n        <PARENT_BILLABLE>N</PARENT_BILLABLE>\\n\r\n        <PARENT_CHARGEABLE>N</PARENT_CHARGEABLE>\\n\r\n        <TOP_TASK_ID></TOP_TASK_ID>\\n\r\n        <SUBTASK_WO_NUMBER></SUBTASK_WO_NUMBER>\\n\r\n        <SUBTASK_WO_NAME></SUBTASK_WO_NAME>\\n\r\n        <SUBTASK_WO_BILLABLE></SUBTASK_WO_BILLABLE>\\n\r\n        <SUBTASK_WO_CHARGEABLE></SUBTASK_WO_CHARGEABLE>\\n\r\n        <CREW_LEADER></CREW_LEADER>\\n\r\n        <RATE_SCHEDULE_DETAILS>RATE SCHEDULE DETAIL-&gt;</RATE_SCHEDULE_DETAILS>\\n\r\n        <EXP_NAME>Unit Production</EXP_NAME>\\n\r\n        <RATE_SCHEDULE_NAME>2023-FPL</RATE_SCHEDULE_NAME>\\n\r\n        <UNIT_NAME>50|OAA01-MAN HOUR RATE PER MAN OH|N|INSTALL|108|HOUR</UNIT_NAME>\\n\r\n        <RATE>135.02</RATE>\\n\r\n        <UNIT_OF_MEASURE>HOURS</UNIT_OF_MEASURE>\\n\r\n        <RATE_START_DATE>2023-01-02</RATE_START_DATE>\\n\r\n        <RATE_END_DATE></RATE_END_DATE>\\n\r\n        <FBDILOADER>FBDI</FBDILOADER>\\n\r\n        <EXPENDITUREDATE>2023-01-22</EXPENDITUREDATE>\\n\r\n        <PERSONNAME></PERSONNAME>\\n\r\n        <PERSONNUMBER></PERSONNUMBER>\\n\r\n        <HUMANRESOURCEASSIGNMENT></HUMANRESOURCEASSIGNMENT>\\n\r\n        <PROJECTNAME>23-01219-000 OHD FL BOCA RATON</PROJECTNAME>\\n\r\n        <PROJECTNUMBER>23-01219-000</PROJECTNUMBER>\\n\r\n        <TASK_NAME></TASK_NAME>\\n\r\n        <TASK_NUMBER></TASK_NUMBER>\\n\r\n        <EXPENDITURETYPE>Unit Production</EXPENDITURETYPE>\\n\r\n        <EXPENDITUREORGANIZATION>Florida</EXPENDITUREORGANIZATION>\\n\r\n        <CONTRACTNUMBER></CONTRACTNUMBER>\\n\r\n        <FUNDINGSOURCENUMBER></FUNDINGSOURCENUMBER>\\n\r\n        <NONLABORRESOURCE>50|OAA01-MAN HOUR RATE PER MAN OH|N|INSTALL|108|HOUR</NONLABORRESOURCE>\\n\r\n        <NONLABORRESOURCEORGANIZATION>PIKE Project Unit Org</NONLABORRESOURCEORGANIZATION>\\n\r\n        <QUANTITY>10</QUANTITY>\\n\r\n        <WORKTYPE></WORKTYPE>\\n\r\n        <ADDITIONALINFO>Additional Info</ADDITIONALINFO>\\n\r\n        <PROJID>300002425947378</PROJID>\\n\r\n        <NLRID>300000013044500</NLRID>\\n\r\n        <TASKID>100007676556772</TASKID>\\n\r\n        <EXISTINGQTYINPPM></EXISTINGQTYINPPM>\\n\r\n        <REGION>Florida</REGION>\\n\r\n        <LEGALENTITY>Pike Electric, LLC</LEGALENTITY>\\n\r\n        <BUSINESSUNIT>Pike Business Unit</BUSINESSUNIT>\\n\r\n    </ROW>\\n\r\n</ROWSET>\\n\"";

                        data = Regex.Replace(data, "^\"|\"$", "");
                        data = data.Replace(System.Environment.NewLine, replaceWith);
                        data = data.Replace("\r\n", replaceWith).Replace("\\n", replaceWith).Replace("\r", replaceWith).Replace("version=\\\"1.0\\\"", "version=\"1.0\"").Replace("encoding=\\\"UTF-8\\\"", "encoding=\"UTF-8\"");




                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(data);

                        var json = JsonConvert.SerializeXmlNode(doc, (Newtonsoft.Json.Formatting)System.Xml.Formatting.None, true);
                        json = json.Replace("\"?xml\":{\"@version\":\"1.0\",\"@encoding\":\"UTF-8\"}{\"ROW\":", "");
                        json = json.Trim().TrimEnd('}');
                        json = json + "}";
                        dynamic dyArray = JObject.Parse(json);
                        var dynamicObject = JsonConvert.DeserializeObject<dynamic>(json)!;


                        string PARAM_PROJ_TO_BILL = dynamicObject.PARAM_PROJ_TO_BILL.ToString();
                        string PARAM_WONO_TO_BILL = dynamicObject.PARAM_WONO_TO_BILL.ToString();
                        string PARAM_SUBTASK_TO_BILL = dynamicObject.PARAM_SUBTASK_TO_BILL.ToString();
                        string PARAM_UNIT_TO_BILL = dynamicObject.PARAM_UNIT_TO_BILL.ToString();
                        string PARAM_QUANTITY = dynamicObject.PARAM_QUANTITY.ToString();
                        string PARAM_WE_DATE = dynamicObject.PARAM_WE_DATE.ToString();
                        string ERROR_CODES = dynamicObject.ERROR_CODES.ToString();
                        string ORACLE_TASK_NO = dynamicObject.ORACLE_TASK_NO.ToString();
                        string ORACLE_SUBTASK = dynamicObject.ORACLE_SUBTASK.ToString();
                        string PARENT_BILLABLE_FLAG = dynamicObject.PARENT_BILLABLE_FLAG.ToString();
                        string SUBTASK_BILLABLE_CHARGEABLE_FLAG = dynamicObject.SUBTASK_BILLABLE_CHARGEABLE_FLAG.ToString();
                        string RATE_SCHEDUL_ERROR = dynamicObject.RATE_SCHEDUL_ERROR.ToString();
                        string PROJ_TASK_DETAILS = dynamicObject.PROJ_TASK_DETAILS.ToString();
                        string PROJECT_NUMBER = dynamicObject.PROJECT_NUMBER.ToString();
                        string PROJECT_NAME = dynamicObject.PROJECT_NAME.ToString();
                        string SUBTASK_PROJECT = dynamicObject.SUBTASK_PROJECT.ToString();
                        string PARENT_WO_NUMBER = dynamicObject.PARENT_WO_NUMBER.ToString();
                        string PARENT_WO_NAME = dynamicObject.PARENT_WO_NAME.ToString();
                        string PARENT_BILLABLE = dynamicObject.PARENT_BILLABLE.ToString();
                        string PARENT_CHARGEABLE = dynamicObject.PARENT_CHARGEABLE.ToString();
                        string TOP_TASK_ID = dynamicObject.TOP_TASK_ID.ToString();
                        string SUBTASK_WO_NUMBER = dynamicObject.SUBTASK_WO_NUMBER.ToString();
                        string SUBTASK_WO_NAME = dynamicObject.SUBTASK_WO_NAME.ToString();
                        string SUBTASK_WO_BILLABLE = dynamicObject.SUBTASK_WO_BILLABLE.ToString();
                        string SUBTASK_WO_CHARGEABLE = dynamicObject.SUBTASK_WO_CHARGEABLE.ToString();
                        string CREW_LEADER = dynamicObject.CREW_LEADER.ToString();
                        string RATE_SCHEDULE_DETAILS = dynamicObject.RATE_SCHEDULE_DETAILS.ToString();
                        string EXP_NAME = dynamicObject.EXP_NAME.ToString();
                        string RATE_SCHEDULE_NAME = dynamicObject.RATE_SCHEDULE_NAME.ToString();
                        string UNIT_NAME = dynamicObject.UNIT_NAME.ToString();
                        string RATE = dynamicObject.RATE.ToString();
                        string UNIT_OF_MEASURE = dynamicObject.UNIT_OF_MEASURE.ToString();
                        string RATE_START_DATE = dynamicObject.RATE_START_DATE.ToString();
                        string RATE_END_DATE = dynamicObject.RATE_END_DATE.ToString();
                        string FBDILOADER = dynamicObject.FBDI_LOADER.ToString();
                        string EXPENDITUREDATE = dynamicObject.FBDI_EXPENDITURE_DATE.ToString();
                        string PERSONNAME = dynamicObject.FBDI_PERSON_NAME.ToString();
                        string PERSONNUMBER = dynamicObject.FBDI_PERSON_NUMBER.ToString();
                        string HUMANRESOURCEASSIGNMENT = dynamicObject.FBDI_HUMAN_RESOURCE_ASSIGNMENT.ToString();
                        string PROJECTNAME = dynamicObject.FBDI_PROJECT_NAME.ToString();
                        string PROJECTNUMBER = dynamicObject.FBDI_PROJECT_NUMBER.ToString();
                        string TASK_NAME = dynamicObject.FBDI_TASK_NAME.ToString();
                        string TASK_NUMBER = dynamicObject.FBDI_TASK_NUMBER.ToString();
                        string EXPENDITURETYPE = dynamicObject.FBDI_EXPENDITURE_TYPE.ToString();
                        string EXPENDITUREORGANIZATION = dynamicObject.FBDI_EXPENDITURE_ORGANIZATION.ToString();
                        string CONTRACTNUMBER = dynamicObject.FBDI_CONTRACT_NUMBER.ToString();
                        string FUNDINGSOURCENUMBER = dynamicObject.FBDI_FUNDING_SOURCE_NUMBER.ToString();
                        string NONLABORRESOURCE = dynamicObject.FBDI_NONLABOR_RESOURCE.ToString();
                        string NONLABORRESOURCEORGANIZATION = dynamicObject.FBDI_NONLABOR_RESOURCE_ORGANIZATION.ToString();
                        string QUANTITY = dynamicObject.FBDI_QUANTITY.ToString();
                        string WORKTYPE = dynamicObject.FBDI_WORK_TYPE.ToString();
                        string ADDITIONALINFO = dynamicObject.ADDITIONAL_INFO.ToString();
                        string PROJID = dynamicObject.PROJID.ToString();
                        string NLRID = dynamicObject.NLRID.ToString();
                        string TASKID = dynamicObject.TASKID.ToString();
                        string EXISTINGQTYINPPM = dynamicObject.EXISTINGQTYINPPM.ToString();
                        string REGION = dynamicObject.REGION.ToString();
                        string LEGALENTITY = dynamicObject.LEGAL_ENTITY.ToString();
                        string BUSINESSUNIT = dynamicObject.BUSINESS_UNIT.ToString();

                        //Delete NLR Data if already Exist
                        clsDAL.NLR_Data_Delete(Compatible_Unit_Id);

                        int res = clsDAL.NLR_Create(PARAM_PROJ_TO_BILL, PARAM_WONO_TO_BILL, PARAM_SUBTASK_TO_BILL, PARAM_UNIT_TO_BILL, PARAM_QUANTITY, PARAM_WE_DATE, ERROR_CODES, ORACLE_TASK_NO, ORACLE_SUBTASK, PARENT_BILLABLE_FLAG, SUBTASK_BILLABLE_CHARGEABLE_FLAG, RATE_SCHEDUL_ERROR, PROJ_TASK_DETAILS, PROJECT_NUMBER, PROJECT_NAME, SUBTASK_PROJECT, PARENT_WO_NUMBER, PARENT_WO_NAME, PARENT_BILLABLE, PARENT_CHARGEABLE, TOP_TASK_ID, SUBTASK_WO_NUMBER, SUBTASK_WO_NAME, SUBTASK_WO_BILLABLE, SUBTASK_WO_CHARGEABLE, CREW_LEADER, RATE_SCHEDULE_DETAILS, EXP_NAME, RATE_SCHEDULE_NAME, UNIT_NAME, RATE, UNIT_OF_MEASURE, RATE_START_DATE, RATE_END_DATE, FBDILOADER, EXPENDITUREDATE, PERSONNAME, PERSONNUMBER, HUMANRESOURCEASSIGNMENT, PROJECTNAME, PROJECTNUMBER, TASK_NAME, TASK_NUMBER, EXPENDITURETYPE, EXPENDITUREORGANIZATION, CONTRACTNUMBER, FUNDINGSOURCENUMBER, NONLABORRESOURCE, NONLABORRESOURCEORGANIZATION, QUANTITY, WORKTYPE, ADDITIONALINFO, PROJID, NLRID, TASKID, EXISTINGQTYINPPM, REGION, LEGALENTITY, BUSINESSUNIT, Compatible_Unit_Id);

                        if (ds.Tables[0].Rows.Count == Count)
                        {
                            clsDAL.WorkOrder_StatusUpdate(Convert.ToInt32(WorkOrder_Id), 4);
                        }

                        Count++;

                    }
                    catch (Exception exp)
                    {
                        oracle_status_details = "Error";
                        oracle_message_details = exp.Message.ToString();

                        Console.WriteLine("GetDukeProjectTaskData - Failure" + exp.Message.ToString());
                        writer.WriteLine("GetDukeProjectTaskData - Failure" + exp.Message.ToString());
                       // clsDAL.SQL_CU_Data_Update(Convert.ToInt32(Compatible_Unit_Id), oracle_status_details, oracle_message_details);
                    }
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("GetDukeProjectTaskData - Failure" + exp.Message.ToString());
                writer.WriteLine("GetDukeProjectTaskData - Failure" + exp.Message.ToString());
            }

        }
        #endregion

        #region GetOracle_WorkORderList
        public static void GetOracle_WorkORderList(StreamWriter writer, DateTime dt)
        {
            writer.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeProjectTaskData - Started");
            Console.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeProjectTaskData - Started");

            DataSet ds = clsDAL.WorkOrder_GetList_Oracle(dt);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Console.WriteLine("GetOracle_WorkORderList - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                writer.WriteLine("GetOracle_WorkORderList - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                
                int WorkOrder_Id = Convert.ToInt32(ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString())!;

                PPM_Push_Oracle(writer, WorkOrder_Id);


            }
        }
        #endregion

        #region PPM_Push_Oracle
        public static void PPM_Push_Oracle(StreamWriter writer, int WorkOrder_Id)
        {
            writer.WriteLine("DynamicsPikeService (Create) - " + AppName + " - Started");
            Console.WriteLine("DynamicsPikeService - " + AppName + " - PPM_Push_Oracle - Started");

            string TransactionNumber = "";
            string UnprocessTransactionId = "";

            string DetailRecordId = "";
            string IntegrationBatchHeaderId = "";
            string ExpenditureBatchName = "";

            DataSet dsDetails = clsDAL.Oracle_PPM_GetList(WorkOrder_Id);

            try
            {
                if (dsDetails.Tables[0].Rows.Count > 0)
                {
                    int RecordCount = 0;
                    for (int j = 0; j < dsDetails.Tables[0].Rows.Count; j++)
                    {
                        string oracle_status_details = "";
                        string oracle_message_details = "";
                        try
                        {
                            ExpenditureBatchName = dsDetails.Tables[0].Rows[j]["BatchName"].ToString()!;
                            //PPM_EXEC_ESSJOB(ExpenditureBatchName);
                            IntegrationBatchHeaderId = dsDetails.Tables[0].Rows[j]["IntegrationBatchHeaderId"].ToString()!;
                            UnprocessTransactionId = dsDetails.Tables[0].Rows[j]["UnprocessTransactionId"].ToString()!;

                            Console.WriteLine("PushDataTo_Oracle - Processing - " + (j + 1).ToString() + " out of " + dsDetails.Tables[0].Rows.Count.ToString());
                            writer.WriteLine("PushDataTo_Oracle - Processing - " + (j + 1).ToString() + " out of " + dsDetails.Tables[0].Rows.Count.ToString());

                            int RecordFlag = 1;

                            byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(UserName + ":" + Password);
                            string credentials = System.Convert.ToBase64String(toEncodeAsBytes);

                            DetailRecordId = dsDetails.Tables[0].Rows[j]["DetailRecordId"].ToString()!.ToLower();

                            string? OriginalTransactionReference = "Replicon_" + DetailRecordId;


                            #region Check if the record got created
                            writer.WriteLine("Checking if the Transaction is already pushed to PPM for Detail Record Id " + DetailRecordId);
                            string? PPM_QueryParam = "?q=OriginalTransactionReference=" + OriginalTransactionReference + ";NetZeroItemFlag is null&expand=ProjectStandardCostCollectionFlexFields";
                            writer.WriteLine("Payload for Getting the Transaction Number " + PPM_QueryParam);

                            var PPMoptions = new RestClientOptions(Oracle_Url)
                            {
                                MaxTimeout = -1,
                            };
                            var PPMclient = new RestClient(PPMoptions);
                            var PPMrequest = new RestRequest(PPM_SubUrl + PPM_QueryParam, Method.Get);
                            PPMrequest.AddHeader("Content-Type", "application/json");
                            PPMrequest.AddHeader("Authorization", "Basic " + credentials);

                            RestResponse PPMresponse = PPMclient.Execute(PPMrequest);
                            writer.WriteLine("PPM Response to get TransactionNumber : " + PPMresponse.Content);


                            #endregion
                            if (RecordFlag == 1)
                            {

                                Console.WriteLine("Processing Unit Billing Detail Events Records - " + (j + 1).ToString() + " - out of " + dsDetails.Tables[0].Rows.Count.ToString());
                                writer.WriteLine("Processing Unit Billing Detail Events Records - " + (j + 1).ToString() + " - out of " + dsDetails.Tables[0].Rows.Count.ToString());


                                string? BusinessUnit = dsDetails.Tables[0].Rows[j]["BusinessUnit"].ToString();
                                string? TransactionSource = dsDetails.Tables[0].Rows[j]["TransactionSource"].ToString();

                                string? Document = dsDetails.Tables[0].Rows[j]["Document"].ToString();
                                string? DocumentEntry = dsDetails.Tables[0].Rows[j]["DocumentEntry"].ToString();


                                string? NonlaborResourceID = dsDetails.Tables[0].Rows[j]["NonlaborResourceID"].ToString();
                                string? NonlaborResource = dsDetails.Tables[0].Rows[j]["NonlaborResource"].ToString();
                                string? NonlaborResourceOrganization = dsDetails.Tables[0].Rows[j]["NonlaborResourceOrganization"].ToString();
                                string? TransactionCurrencyCode = dsDetails.Tables[0].Rows[j]["TransactionCurrencyCode"].ToString();

                                string? Quantity = dsDetails.Tables[0].Rows[j]["Quantity"].ToString();
                                string? UnitOfMeasure = dsDetails.Tables[0].Rows[j]["UnitOfMeasure"].ToString();
                                //string? UnitOfMeasure = "Hours";

                                DateTime dtExpenditureDate = Convert.ToDateTime(dsDetails.Tables[0].Rows[j]["EXPENDITURE_ITEM_DATE"].ToString());
                                string? ExpenditureDate = dtExpenditureDate.ToString("yyyy-MM-dd");

                                string? ProjectId = dsDetails.Tables[0].Rows[j]["PROJECT_ID"].ToString();
                                string? TaskId = dsDetails.Tables[0].Rows[j]["SUB_TASK_ID"].ToString();
                                string? ExpenditureTypeId = dsDetails.Tables[0].Rows[j]["EXPENDITURE_TYPE_ID_Display"].ToString();
                                string? OrganizationId = dsDetails.Tables[0].Rows[j]["ORGANIZATION_ID_Display"].ToString();


                                writer.WriteLine("Updating the Detail Record Status to In Progress for " + DetailRecordId);

                                //PPM UPC CREATE
                                var options = new RestClientOptions(Oracle_Url)
                                {
                                    MaxTimeout = -1,
                                };
                                var client = new RestClient(options);
                                var request = new RestRequest(UnitBilling_Sub_Url, Method.Post);
                                request.AddHeader("Content-Type", "application/json");
                                request.AddHeader("Authorization", "Basic " + credentials);

                                var body = "{"

                                                + "\"ExpenditureBatch\":\"" + ExpenditureBatchName + "\","
                                                + "\"BusinessUnit\":\"" + BusinessUnit + "\","
                                                + "\"TransactionSource\":\"" + TransactionSource + "\","
                                                + "\"Document\":\"" + Document + "\","
                                                + "\"DocumentEntry\":\"" + DocumentEntry + "\","
                                                + "\"NonlaborResourceId\":\"" + NonlaborResourceID + "\","
                                                + "\"NonlaborResource\":\"" + NonlaborResource + "\","
                                                + "\"NonlaborResourceOrganization\":\"" + NonlaborResourceOrganization + "\","
                                                + "\"TransactionCurrencyCode\":\"" + TransactionCurrencyCode + "\","
                                                + "\"OriginalTransactionReference\":\"" + OriginalTransactionReference + "\","
                                                + "\"Quantity\":\"" + Quantity + "\","
                                                + "\"UnitOfMeasure\":\"" + UnitOfMeasure + "\","
                                                + "\"ProjectStandardCostCollectionFlexfields\": ["
                                                        + "{"
                                                        + "\"_EXPENDITURE_ITEM_DATE\":\"" + ExpenditureDate + "\","
                                                    + "\"_PROJECT_ID\":\"" + ProjectId + "\","
                                                    + "\"_TASK_ID\":\"" + TaskId + "\","
                                                    + "\"_EXPENDITURE_TYPE_ID_Display\":\"" + ExpenditureTypeId + "\","
                                                    + "\"_ORGANIZATION_ID_Display\":\"" + OrganizationId + "\""
                                                    + "}]"
                                                    + "}";



                                request.AddStringBody(body, DataFormat.Json);
                                RestResponse response = client.Execute(request);

                                writer.WriteLine("Payload for Detail Transaction Record :  " + body);

                                if (response.Content != "" && response.StatusCode.ToString() == "Created")
                                {
                                    

                                    dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;
                                    writer.WriteLine("Oracle Response for Detail Transaction Record :  " + response.Content!);

                                    writer.WriteLine("Updating the Detail Record Status in Dynamics along with other values for " + DetailRecordId);

                                    oracle_status_details = "Success";
                                    oracle_message_details = "Sent to UPC";
                                    UnprocessTransactionId = dyArray.UnprocessedTransactionReferenceId.ToString();


                                    clsDAL.NLR_DataUpdate("", oracle_status_details, oracle_message_details, UnprocessTransactionId, DetailRecordId);
                                }
                                else
                                {
                                    writer.WriteLine("Oracle Response for Detail Transaction Record :  " + response.Content!);
                                    writer.WriteLine("Updating the Detail Record Status in Dynamics along with other values for " + DetailRecordId);
                                    oracle_status_details = "Error";
                                    oracle_message_details = response.Content!.ToString();
                                    //Update the PPM Status      

                                    clsDAL.NLR_DataUpdate("", oracle_status_details, oracle_message_details, "", DetailRecordId);
                                }
                            }

                            RecordCount++;

                            if (dsDetails.Tables[0].Rows.Count == RecordCount)
                            {
                                clsDAL.WorkOrder_StatusUpdate(Convert.ToInt32(WorkOrder_Id), 5);
                            }
                            
                        }
                        catch (Exception exp)
                        {
                            writer.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);
                            Console.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);

                            oracle_status_details = "Error";
                            oracle_message_details = exp.Message.ToString();

                            clsDAL.NLR_DataUpdate("", oracle_status_details, oracle_message_details, "", DetailRecordId);
                        }
                    }


                }
            }
            catch (Exception exp)
            {
                writer.WriteLine("Application Exception " + exp.Message);
                Console.WriteLine("Application Exception " + exp.Message);
            }


            //PPM_EXEC_ESSJOB(ExpenditureBatchName);


            writer.WriteLine("DynamicsPikeService (Create) - " + AppName + " - Completed");
        }
        #endregion

        #region GenericFunction

        #region PPM_EXEC_ESSJOB
        public static void PPM_EXEC_ESSJOB(string ExpenditureBatch)
        {

            string? PPM_EssJob_Url = System.Configuration.ConfigurationManager.AppSettings["Oracle_Url"]!.ToString();
            string? PPM_EssJob_Sub_Url = System.Configuration.ConfigurationManager.AppSettings["PPM_EssJob_Sub_Url"]!.ToString();

            Console.WriteLine("DynamicsPikeService - PPM_ESSJOB Started :" + DateTime.Now.ToString("yyyy-MM-ddTHH\\_mm"));

            string? strBU_Name = System.Configuration.ConfigurationManager.AppSettings["Business_Unit"]!.ToString();
            string? strEss_Units = System.Configuration.ConfigurationManager.AppSettings["ESS_Job_Units"]!.ToString();
            string? strCost_Units = System.Configuration.ConfigurationManager.AppSettings["ESS_Cost_Units"]!.ToString();
            string? ESSParameters_Batch = "Pike Business Unit," + strBU_Name + ",IMPORT_AND_PROCESS,PREV_NOT_IMPORTED,," + strEss_Units + "," + strCost_Units + "," + ExpenditureBatch + ",,,,,ORA_PJC_DETAIL";
            Console.WriteLine("ESSParameters_Units : " + ESSParameters_Batch);
            string strESS_Job_Units = PPM_ESSJOB(PPM_EssJob_Url, PPM_EssJob_Sub_Url, UserName, Password, ESSParameters_Batch);
            Console.WriteLine("PPM ESS JOB Created Successful!!");
            Console.WriteLine("");

        }
        #endregion

        #region PPM_ESSJOB
        public static string PPM_ESSJOB(string MainUrl, string SubUrl, string username, string pwd, string ESSParameters)
        {
            string? UserJSONString = string.Empty;

            string? TransactionType = "";
            int? Transaction_Id = 0;

            string? OperationName = "submitESSJobRequest";
            string? JobPackageName = "oracle/apps/ess/projects/costing/transactions/onestop";
            string? JobDefName = "ImportProcessParallelEssJob";
            //string ESSParameters = "Pike Engineering Business Unit,300001588285232,IMPORT_AND_PROCESS,PREV_NOT_IMPORTED,,300000007509139,," + Expenditure_Batch + ",,,,,ORA_PJC_DETAIL";
            string? ReqstId = "null";
            string? strEss_Job_Result = "";
            UserJSONString = "{\"OperationName\":\"" + OperationName + "\", \"JobPackageName\":\"" + JobPackageName + "\", \"JobDefName\": \"" + JobDefName + "\",\"ESSParameters\":\"" + ESSParameters + "\",\"ReqstId\":\"" + ReqstId + "\" }";
            //Console.WriteLine("Oracle Request : " + UserJSONString);

            var options = new RestClientOptions(MainUrl)
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest(SubUrl, Method.Post);

            byte[] toEncodeAsBytes1 = System.Text.ASCIIEncoding.ASCII.GetBytes(username + ":" + pwd);
            string? credentials1 = System.Convert.ToBase64String(toEncodeAsBytes1);
            request.AddHeader("Authorization", "Basic " + credentials1);

            // request.AddHeader("Content-Type", "application/json");
            //request.AddParameter("application/json", UserJSONString, ParameterType.RequestBody);
            request.AddHeader("Content-Type", "application/vnd.oracle.adf.resourceitem+json");
            request.AddStringBody(UserJSONString, DataFormat.Json);
            RestResponse response = client.Execute(request);

            if (response.Content != "" && response.StatusCode.ToString() == "Created" || response.StatusCode.ToString() == "NoContent")
            {
                dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;

                TransactionType = dyArray.JobDefName!.ToString();
                Transaction_Id = dyArray.ReqstId;
                strEss_Job_Result = Transaction_Id.ToString();
                Console.WriteLine("Oracle Schedule Transaction ID :  " + strEss_Job_Result);
                Console.WriteLine("---------------------------------------");
            }

            return strEss_Job_Result!;
        }
        #endregion       

        #region UpdateWorkorderStatus
        public static string UpdateWorkorderStatus(string WorkorderNumber)
        {
            try
            {
                string Token = Dyn_GetToken();
                string msdyn_workorderid = "";
                string WorkorderName = "";
                string WorkorderStatus = "Completed";
                string WorkorderStatusFlag = "false";

                string SuccessFalg = "Failure";

                //Get the WorkorderList

                var options = new RestClientOptions(Dynamics_Url)
                {
                    MaxTimeout = -1,
                };
                var client = new RestClient(options);                
                var request = new RestRequest("/api/data/v9.2/msdyn_workorders?$select=hsl_readyforbillingoraclestatus,msdyn_workorderid,hsl_workorderid&$filter=hsl_workorderid eq " + "'"+ WorkorderNumber + "'", Method.Get);
                request.AddHeader("OData-MaxVersion", "4.0");
                request.AddHeader("OData-Version", "4.0");
                request.AddHeader("Content-Type", "application/json; charset=utf-8");
                request.AddHeader("Accept", "application/json");
                request.AddHeader("Prefer", "odata.include-annotations=*");
                request.AddHeader("Authorization", "Bearer " + Token);               
                RestResponse response = client.Execute(request);
                if (response.Content != "" && response.StatusCode.ToString() == "OK")
                {
                    dynamic HArray = JObject.Parse(response.Content!);
                    msdyn_workorderid = HArray.value[0].msdyn_workorderid.ToString();
                    WorkorderName = HArray.value[0].hsl_workorderid.ToString();
                }


                //Update the Workorder status
                if (msdyn_workorderid != "")
                {
                    var msdyn_options = new RestClientOptions(Dynamics_Url)
                    {
                        MaxTimeout = -1,
                    };
                    var msdyn_client = new RestClient(msdyn_options);
                    var msdyn_request = new RestRequest("/api/data/v9.2/msdyn_workorders("+ msdyn_workorderid +")", Method.Patch);
                    msdyn_request.AddHeader("OData-MaxVersion", "4.0");
                    msdyn_request.AddHeader("OData-Version", "4.0");
                    msdyn_request.AddHeader("Content-Type", "application/json; charset=utf-8");
                    msdyn_request.AddHeader("Accept", "application/json");
                    msdyn_request.AddHeader("Prefer", "odata.include-annotations=*");
                    msdyn_request.AddHeader("Authorization", "Bearer " + Token);

                    var body = "{"
                    + "\"hsl_readyforbillingoraclestatus\":\"" + WorkorderStatus + "\" ,"
                    + "\"hsl_readyforbilling\":\"" + WorkorderStatusFlag + "\" }";
                    msdyn_request.AddStringBody(body, DataFormat.Json);
                    RestResponse msdyn_response = msdyn_client.Execute(msdyn_request);
                    if (msdyn_response.Content != "" && msdyn_response.StatusCode.ToString() == "Created" || msdyn_response.StatusCode.ToString() == "NoContent")
                    {
                        SuccessFalg = "Success";
                    }

                }
                return SuccessFalg;
            }
            catch (Exception exp)
            {

                throw new Exception(exp.Message);
            }
        }
        #endregion

        #region GetToken
        public static string GetToken()
        {
            string OauthUrl = System.Configuration.ConfigurationManager.AppSettings["OauthUrl"]!.ToString();
            string OauthTokenUrl = System.Configuration.ConfigurationManager.AppSettings["OauthTokenUrl"]!.ToString();
            string OauthTokenCred = System.Configuration.ConfigurationManager.AppSettings["OauthTokenCred"]!.ToString();
            string AccessToken = "";
            var options = new RestClientOptions(OauthUrl)
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest(OauthTokenUrl, Method.Post);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddHeader("Authorization", "Basic " + OauthTokenCred);
            request.AddParameter("grant_type", "client_credentials");
            RestResponse response = client.Execute(request);
            if (response.Content != "" && response.StatusCode.ToString() == "OK")
            {
                var dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;
                //var dyArray = JObject.Parse(response.Content!);
                AccessToken = dyArray.access_token;
                return AccessToken;

            }

            return null!;

        }
        #endregion

        #region Generate Token
        public static string Dyn_GetToken()
        {
            string Client_Id = System.Configuration.ConfigurationManager.AppSettings["Client_Id"]!.ToString();
            string Client_Secret = clsTools.Decrypt(System.Configuration.ConfigurationManager.AppSettings["Client_Secret"]!.ToString(), true);
            string Grant_Type = System.Configuration.ConfigurationManager.AppSettings["Grant_Type"]!.ToString();
            string Realm = System.Configuration.ConfigurationManager.AppSettings["Realm"]!.ToString();
            string resourceurl = HttpUtility.UrlEncode(Dynamics_Url);
            string AccessToken = "";

            var options = new RestClientOptions("https://login.microsoftonline.com")
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest("/" + Realm + "/oauth2/token", Method.Post);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("grant_type", "client_credentials");
            request.AddParameter("client_id", "" + Client_Id + "@" + Realm + "");
            request.AddParameter("client_secret", Client_Secret);
            request.AddParameter("resource", Dynamics_Url);
            RestResponse response = client.Execute(request);



            if (response.Content != "" && response.StatusCode.ToString() == "OK")
            {
                var dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;
                //var dyArray = JObject.Parse(response.Content!);
                AccessToken = dyArray.access_token;
                return AccessToken;

            }

            return null!;
        }

        #endregion

        #endregion

    }


}