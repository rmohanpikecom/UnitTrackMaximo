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

                //GetWorkorders_DYN(Writer);

                GetDukeMaximoUnits(Writer);

                //UpdateServiceItemData(Writer);

                //GetSubTaskNumber_Oracle(Writer);

                //GetNLRData_Oracle(Writer);

                //PushDataTo_Oracle(Writer);

                //UpdateDataTo_Oracle(Writer);

                //Upload Timesheets from Dynamics to Oracle
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
        public static void GetWorkorders_DYN(StreamWriter writer)
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
                            ProcessDate = DateTime.Now.ToString("yyyy-MM-dd");
                            Project_SubTask_Flag = Convert.ToInt32(dsProjectTask.Tables[0].Rows[i]["Project_SubTask_Flag"].ToString()!);

                            int res = clsDAL.Dynamics_WorkOrder_Create(Parent_Task_Number, Parent_Task_Id, Project_Number, Project_Id, BusinessUnitName, ProcessDate, Project_SubTask_Flag);
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

        #region GetDukeMaximoUnits
        public static void GetDukeMaximoUnits(StreamWriter writer)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeMaximoUnits - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeMaximoUnits - Started");

                DataSet ds = clsDAL.SQL_WorkOrder_GetList(1);

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetDukeMaximoUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeMaximoUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        int WorkOrder_Id = 0;
                        string WorkOrder_Number = "";

                        WorkOrder_Id= Convert.ToInt32(ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString());
                        WorkOrder_Number = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;

                        // GetDukeMaximoUnits_Details(WorkOrder_Id, WorkOrder_Number, writer);
                        GetCPR_Details(WorkOrder_Id, WorkOrder_Number, writer);

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


            string ServiceUrl = System.Configuration.ConfigurationManager.AppSettings["ServiceUrl"]!.ToString();
            string ServiceSubUrl = System.Configuration.ConfigurationManager.AppSettings["ServiceSubUrl"]!.ToString();

            try
            {
                if (token != null)
                {
                    var options = new RestClientOptions(ServiceUrl)
                    {
                        MaxTimeout = -1,
                    };
                    var client = new RestClient(options);
                    var request = new RestRequest(ServiceSubUrl + WorkOrder_Number, Method.Get);
                    request.AddHeader("Authorization", "Bearer " + token);
                    RestResponse response = client.Execute(request);
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                        dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!.ToString())!;
                    else
                        return result;

                    result = dyArray.ToString();

                    string Station_Detail_ID = "";
                    string Payment_Status = "";
                    string Work_Function = "";
                    string Estimate_Quantity = "";
                    string Compatible_Unit_Quantity = "";
                    string Completed_Date = "";
                    string Compatible_Unit_Changed_By = "";
                    string Compatible_Unit_Changed_Date = "";
                    string Compatible_Unit_Station = "";
                    string Parent_Compatible_Unit_Name = "";
                    string Parent_Compatible_Unit_Description = "";
                    string Payment_Type = "";
                    string Field_To_From = "";
                    string Field_ID_To = "";
                    string Serial_Number = "";
                    string Manufacturer = "";
                    string Vendor_Id = "";

                    int Count = 1;
                    int RecordCount = dyArray.responseList.Count;

                    if (dyArray.responseList.Count != null)
                    {
                        for (int i = 0; i < dyArray.responseList.Count; i++)
                        {
                            try
                            {
                                Console.WriteLine("GetDukeMaximoUnits_Details - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray.responseList.Count.ToString());
                                writer.WriteLine("GetDukeMaximoUnits_Details - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray.responseList.Count.ToString());

                                if (dyArray.responseList[i].Station_Detail_ID != null)
                                    Station_Detail_ID = dyArray.responseList[i].Station_Detail_ID;

                                if (dyArray.responseList[i].Payment_Status != null)
                                    Payment_Status = dyArray.responseList[i].Payment_Status;

                                if (dyArray.responseList[i].Work_Function != null)
                                    Work_Function = dyArray.responseList[i].Work_Function;

                                if (dyArray.responseList[i].Estimate_Quantity != null)
                                    Estimate_Quantity = dyArray.responseList[i].Estimate_Quantity;

                                if (dyArray.responseList[i].Compatible_Unit_Quantity != null)
                                    Compatible_Unit_Quantity = dyArray.responseList[i].Compatible_Unit_Quantity;

                                if (dyArray.responseList[i].Completed_Date != null)
                                    Completed_Date = dyArray.responseList[i].Completed_Date;

                                if (dyArray.responseList[i].Compatible_Unit_Changed_By != null)
                                    Compatible_Unit_Changed_By = dyArray.responseList[i].Compatible_Unit_Changed_By;

                                if (dyArray.responseList[i].Compatible_Unit_Changed_Date != null)
                                    Compatible_Unit_Changed_Date = dyArray.responseList[i].Compatible_Unit_Changed_Date;

                                if (dyArray.responseList[i].Compatible_Unit_Station != null)
                                    Compatible_Unit_Station = dyArray.responseList[i].Compatible_Unit_Station;

                                if (dyArray.responseList[i].Parent_Compatible_Unit_Name != null)
                                    Parent_Compatible_Unit_Name = dyArray.responseList[i].Parent_Compatible_Unit_Name;

                                if (dyArray.responseList[i].Parent_Compatible_Unit_Description != null)
                                    Parent_Compatible_Unit_Description = dyArray.responseList[i].Parent_Compatible_Unit_Description;

                                if (dyArray.responseList[i].Payment_Type != null)
                                    Payment_Type = dyArray.responseList[i].Payment_Type;

                                if (dyArray.responseList[i].Field_To_From != null)
                                    Field_To_From = dyArray.responseList[i].Field_To_From;

                                if (dyArray.responseList[i].Field_ID_To != null)
                                    Field_ID_To = dyArray.responseList[i].Field_ID_To;

                                if (dyArray.responseList[i].Serial_Number != null)
                                    Serial_Number = dyArray.responseList[i].Serial_Number;

                                if (dyArray.responseList[i].Manufacturer != null)
                                    Manufacturer = dyArray.responseList[i].Manufacturer;

                                if (dyArray.responseList[i].Vendor_Id != null)
                                    Vendor_Id = dyArray.responseList[i].Vendor_Id;

                                int res = clsDAL.Maximo_MaximoUnits_Create(Station_Detail_ID, Payment_Status, Work_Function, Estimate_Quantity, Compatible_Unit_Quantity, Completed_Date, Compatible_Unit_Changed_By, Compatible_Unit_Changed_Date, Compatible_Unit_Station, Parent_Compatible_Unit_Name, Parent_Compatible_Unit_Description, Payment_Type, Field_To_From, Field_ID_To, Serial_Number, Manufacturer, Vendor_Id, WorkOrder_Id);

                                if (Count == RecordCount)
                                {
                                    clsDAL.SQL_WorkOrder_StatusUpdate(WorkOrder_Id);
                                }

                                Count++;

                            }
                            catch (Exception exp)
                            {
                                writer.WriteLine("UnitTrack Maximo GetDukeMaximoUnits_Details Failure");
                                writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                                Console.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                                writer.WriteLine("=========================================================================");
                            }

                        }
                    }

                }
            }
            catch (Exception exp)
            {

                writer.WriteLine("UnitTrack Maximo GetDukeMaximoUnits_Details Failure");
                writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                Console.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                writer.WriteLine("=========================================================================");
            }
            return result;
        }
        #endregion

        #region UpdateServiceItemData
        public static void UpdateServiceItemData(StreamWriter writer)
        {
            try
            {
                clsDAL.SQL_Update_ServiceItem();
            }
            catch (Exception exp)
            {
                Console.WriteLine("Update Service ItemData - Failure" + exp.Message.ToString());
                writer.WriteLine("Update Service ItemData - Failure" + exp.Message.ToString());
            }
        }
        #endregion

        #region GetSubTaskNumber_Oracle
        public static void GetSubTaskNumber_Oracle(StreamWriter writer)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetSubTaskNumber_Oracle - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetSubTaskNumber_Oracle - Started");

                DataSet ds = clsDAL.SQL_WorkOrder_GetList(2);
                string WorkOrder_Id = "";

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetDukeCompatibleUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeCompatibleUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        WorkOrder_Id = ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString()!;
                        string ProjectNumber = ds.Tables[0].Rows[i]["Project_Number"].ToString()!;
                        string TaskNumber = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;

                        string OracleWrapperUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperUrl"]!.ToString();
                        string OracleWrapperSubUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperSubUrl"]!.ToString();
                        string OracleWrapperEnvironment = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperEnvironment"]!.ToString();

                        var configuration = new ConfigurationBuilder()
                           .AddJsonFile("DBQueries.json", optional: false, reloadOnChange: true)
                           .AddEnvironmentVariables()
                           .Build();
                        string cmd = configuration.GetSection("qryOracleSubTaskData").Value!.Replace("@ProjectNumber", "'" + ProjectNumber + "'").Replace("@TaskNumber", "'" + TaskNumber + "'");

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

                        clsDAL.SQL_Update_SubTaskDetails(WorkOrder_Id, Sub_Task_Number, Sub_Task_Name, Sub_Task_Id, Sub_Task_Billable_Flag, Sub_Task_Crew_Leader, Sub_Task_Project_Flag);
                    }
                    catch (Exception exp)
                    {
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
        public static void GetNLRData_Oracle(StreamWriter writer)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeProjectTaskData - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeProjectTaskData - Started");


                DataSet ds = clsDAL.SQL_NLR_GetList();
                string Compatible_Unit_Id = "";
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string oracle_status_details = "";
                    string oracle_message_details = "";
                    try
                    {
                        Console.WriteLine("GetDukeProjectTaskData - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeProjectTaskData - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

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

                        int res = clsDAL.Oracle_NLR_Create(PARAM_PROJ_TO_BILL, PARAM_WONO_TO_BILL, PARAM_SUBTASK_TO_BILL, PARAM_UNIT_TO_BILL, PARAM_QUANTITY, PARAM_WE_DATE, ERROR_CODES, ORACLE_TASK_NO, ORACLE_SUBTASK, PARENT_BILLABLE_FLAG, SUBTASK_BILLABLE_CHARGEABLE_FLAG, RATE_SCHEDUL_ERROR, PROJ_TASK_DETAILS, PROJECT_NUMBER, PROJECT_NAME, SUBTASK_PROJECT, PARENT_WO_NUMBER, PARENT_WO_NAME, PARENT_BILLABLE, PARENT_CHARGEABLE, TOP_TASK_ID, SUBTASK_WO_NUMBER, SUBTASK_WO_NAME, SUBTASK_WO_BILLABLE, SUBTASK_WO_CHARGEABLE, CREW_LEADER, RATE_SCHEDULE_DETAILS, EXP_NAME, RATE_SCHEDULE_NAME, UNIT_NAME, RATE, UNIT_OF_MEASURE, RATE_START_DATE, RATE_END_DATE, FBDILOADER, EXPENDITUREDATE, PERSONNAME, PERSONNUMBER, HUMANRESOURCEASSIGNMENT, PROJECTNAME, PROJECTNUMBER, TASK_NAME, TASK_NUMBER, EXPENDITURETYPE, EXPENDITUREORGANIZATION, CONTRACTNUMBER, FUNDINGSOURCENUMBER, NONLABORRESOURCE, NONLABORRESOURCEORGANIZATION, QUANTITY, WORKTYPE, ADDITIONALINFO, PROJID, NLRID, TASKID, EXISTINGQTYINPPM, REGION, LEGALENTITY, BUSINESSUNIT, Compatible_Unit_Id);
                    }
                    catch (Exception exp)
                    {
                        oracle_status_details = "Error";
                        oracle_message_details = exp.Message.ToString();

                        Console.WriteLine("GetDukeProjectTaskData - Failure" + exp.Message.ToString());
                        writer.WriteLine("GetDukeProjectTaskData - Failure" + exp.Message.ToString());
                        clsDAL.SQL_CU_Data_Update(Convert.ToInt32(Compatible_Unit_Id), oracle_status_details, oracle_message_details);
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

        #region GetCPR_Details
        public static string GetCPR_Details(int WorkOrder_Id, string WorkOrder_Number, StreamWriter writer)
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
                        return result;

                    result = dyArray.ToString();

                    string Approval_Date = "";
                    string Invoice_Num = "";
                    string Invoice_Status = "";
                    string Invoice_CheckNum = "";
                    string Invoice_De_Psvoucheramt = "";
                    string De_Rejectcode = "";
                    string Rejection_Date = "";
                    string CPR = "";
                    string Invoice_Due_Date = "";
                    string Invoice_Paid_Date = "";
                    string Status = "";
                    string HDescription = "";
                    string Vendor_Invoice_Num = "";
                    string Week_Ending = "";
                    string Vendor_Project_ID = "";
                    string Site = "";
                    string CPR_Type = "";
                    string Reactive_Time_Report_ID = "";
                    string Payment_Type = "";
                    string Foreman = "";
                    string Derived_Contract = "";
                    string First_Approver = "";
                    string Second_Approver = "";
                    string Line_of_Business = "";
                    string Rejection_Remarks = "";
                    string Requested_By_Name = "";
                    string Requested_By_Email = "";
                    string Requested_By_Phone_Num = "";
                    string CPR_Submit_Date = "";
                    string CPR_Created_Date = "";
                    string Total_Cost = "";

                    //CPR Lines
                    string Line = "";
                    string WO_Task_Num = "";
                    string Service_Item = "";
                    string PrLine_Description = "";
                    string PRLineQuantity = "";
                    string Order_Unit = "";
                    string Unit_Cost = "";
                    string Line_Cost = "";
                    string GL_Debit_Account = "";

                    //CU Details
                    string WO_Num = "";
                    string Description = "";
                    string Point_Span = "";
                    string CU_ID = "";
                    string CU_Name = "";
                    string CU_Description = "";
                    string CU_Service_Item = "";
                    string Service_Item_Description = "";
                    string Estimated_Qty = "";
                    string Asbuilt_Qty = "";
                    string Work_Function = "";

                    string Contract = "";

                    //Status History
                    string SH_Status = "";
                    string SH_change_Date = "";
                    string SH_Changed_By = "";
                    string SH_Rejection_Code = "";


                    int Count = 1;
                    int RecordCount = dyArray.Count;

                    if (dyArray.Count != null)
                    {
                        for (int i = 0; i < dyArray[0].CPR_Lines.Count; i++)
                        {
                            try
                            {
                                Console.WriteLine("GetCPR_Details - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray[0].CPR_Lines.Count.ToString());
                                writer.WriteLine("GetCPR_Details - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray[0].CPR_Lines.Count.ToString());

                                if (dyArray[0].Approval_Date != null)
                                    Approval_Date = dyArray[0].Approval_Date.ToString();

                                if (dyArray[0].Invoice_Num != null)
                                    Invoice_Num = dyArray[0].Invoice_Num.ToString();

                                if (dyArray[0].Invoice_Status != null)
                                    Invoice_Status = dyArray[0].Invoice_Status.ToString();

                                if (dyArray[0].Invoice_CheckNum != null)
                                    Invoice_CheckNum = dyArray[0].Invoice_CheckNum.ToString();

                                if (dyArray[0].Invoice_De_Psvoucheramt != null)
                                    Invoice_De_Psvoucheramt = dyArray[0].Invoice_De_Psvoucheramt.ToString();

                                if (dyArray[0].De_Rejectcode != null)
                                    De_Rejectcode = dyArray[0].De_Rejectcode.ToString();

                                if (dyArray[0].Rejection_Date != null)
                                    Rejection_Date = dyArray[0].Rejection_Date.ToString();

                                if (dyArray[0].CPR != null)
                                    CPR = dyArray[0].CPR.ToString();

                                if (dyArray[0].Invoice_Due_Date != null)
                                    Invoice_Due_Date = dyArray[0].Invoice_Due_Date.ToString();

                                if (dyArray[0].Invoice_Paid_Date != null)
                                    Invoice_Paid_Date = dyArray[0].Invoice_Paid_Date.ToString();

                                if (dyArray[0].Status != null)
                                    Status = dyArray[0].Status.ToString();

                                if (dyArray[0].Description != null)
                                    HDescription = dyArray[0].Description.ToString();

                                if (dyArray[0].Vendor_Invoice_Num != null)
                                    Vendor_Invoice_Num = dyArray[0].Vendor_Invoice_Num.ToString();

                                if (dyArray[0].Week_Ending != null)
                                    Week_Ending = dyArray[0].Week_Ending.ToString();

                                if (dyArray[0].Vendor_Project_ID != null)
                                    Vendor_Project_ID = dyArray[0].Vendor_Project_ID.ToString();

                                if (dyArray[0].Site != null)
                                    Site = dyArray[0].Site.ToString();

                                if (dyArray[0].CPR_Type != null)
                                    CPR_Type = dyArray[0].CPR_Type.ToString();

                                if (dyArray[0].Reactive_Time_Report_ID != null)
                                    Reactive_Time_Report_ID = dyArray[0].Reactive_Time_Report_ID.ToString();

                                if (dyArray[0].Payment_Type != null)
                                    Payment_Type = dyArray[0].Payment_Type.ToString();

                                if (dyArray[0].Foreman != null)
                                    Foreman = dyArray[0].Foreman.ToString();

                                if (dyArray[0].Derived_Contract != null)
                                    Derived_Contract = dyArray[0].Derived_Contract.ToString();

                                if (dyArray[0]["1st_Approver"] != null)
                                    First_Approver = dyArray[0]["1st_Approver"].ToString();

                                if (dyArray[0]["2nd_Approver"] != null)
                                    Second_Approver = dyArray[0]["2nd_Approver"].ToString();

                                if (dyArray[0].Line_of_Business != null)
                                    Line_of_Business = dyArray[0].Line_of_Business.ToString();

                                if (dyArray[0].Rejection_Remarks != null)
                                    Rejection_Remarks = dyArray[0].Rejection_Remarks.ToString();

                                if (dyArray[0].Requested_By_Name != null)
                                    Requested_By_Name = dyArray[0].Requested_By_Name.ToString();

                                if (dyArray[0].Requested_By_Email != null)
                                    Requested_By_Email = dyArray[0].Requested_By_Email.ToString();

                                if (dyArray[0].Requested_By_Phone_Num != null)
                                    Requested_By_Phone_Num = dyArray[0].Requested_By_Phone_Num.ToString();

                                if (dyArray[0].CPR_Submit_Date != null)
                                    CPR_Submit_Date = dyArray[0].CPR_Submit_Date.ToString();

                                if (dyArray[0].CPR_Created_Date != null)
                                    CPR_Created_Date = dyArray[0].CPR_Created_Date.ToString();

                                if (dyArray[0].Total_Cost != null)
                                    Total_Cost = dyArray[0].Total_Cost.ToString();

                                if (dyArray[0].Contract != null)
                                    Contract = dyArray[0].Contract.ToString();

                                //CPR Lines

                                if (dyArray[0].CPR_Lines[i].Line != null)
                                    Line = dyArray[0].CPR_Lines[i].Line.ToString();

                                if (dyArray[0].CPR_Lines[i].WO_Task_Num != null)
                                    WO_Task_Num = dyArray[0].CPR_Lines[i].WO_Task_Num.ToString();

                                if (dyArray[0].CPR_Lines[i].Service_Item != null)
                                    Service_Item = dyArray[0].CPR_Lines[i].Service_Item.ToString();

                                if (dyArray[0].CPR_Lines[i].PrLine_Description != null)
                                    PrLine_Description = dyArray[0].CPR_Lines[i].PrLine_Description.ToString();

                                if (dyArray[0].CPR_Lines[i].PRLineQuantity != null)
                                    PRLineQuantity = dyArray[0].CPR_Lines[i].PRLineQuantity.ToString();

                                if (dyArray[0].CPR_Lines[i].Order_Unit != null)
                                    Order_Unit = dyArray[0].CPR_Lines[i].Order_Unit.ToString();

                                if (dyArray[0].CPR_Lines[i].Unit_Cost != null)
                                    Unit_Cost = dyArray[0].CPR_Lines[i].Unit_Cost.ToString();

                                if (dyArray[0].CPR_Lines[i].Line_Cost != null)
                                    Line_Cost = dyArray[0].CPR_Lines[i].Line_Cost.ToString();

                                if (dyArray[0].CPR_Lines[i].GL_Debit_Account != null)
                                    GL_Debit_Account = dyArray[0].CPR_Lines[i].GL_Debit_Account.ToString();


                                //CU Details

                                int res = clsDAL.CPR_Data_Create(Approval_Date, Invoice_Num, Invoice_Status, Invoice_CheckNum, Invoice_De_Psvoucheramt, De_Rejectcode, Rejection_Date, CPR, Invoice_Due_Date, Invoice_Paid_Date, Status, HDescription, Vendor_Invoice_Num, Week_Ending, Vendor_Project_ID, Site, CPR_Type, Reactive_Time_Report_ID, Payment_Type, Foreman, Derived_Contract, First_Approver, Second_Approver, Line_of_Business, Rejection_Remarks, Requested_By_Name, Requested_By_Email, Requested_By_Phone_Num, CPR_Submit_Date, CPR_Created_Date, Total_Cost, Line, WO_Task_Num, Service_Item, PrLine_Description, PRLineQuantity, Order_Unit, Unit_Cost, Line_Cost, GL_Debit_Account, WO_Num, Description, Point_Span, CU_ID, CU_Name, CU_Description, CU_Service_Item, Service_Item_Description, Estimated_Qty, Asbuilt_Qty, Work_Function, Contract, SH_Status, SH_change_Date, SH_Changed_By, SH_Rejection_Code);



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

                    //    for (int i = 0; i < dyArray[0].CU_Details.Count; i++)
                    //    {
                    //        try
                    //        {
                    //            Console.WriteLine("GetCPR_CU_Details - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray[0].CU_Details.Count.ToString());
                    //            writer.WriteLine("GetCPR_CU_Details - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray[0].CU_Details.Count.ToString());

                    //            if (dyArray[0].Approval_Date != null)
                    //                Approval_Date = dyArray[0].Approval_Date.ToString();

                    //            if (dyArray[0].Invoice_Num != null)
                    //                Invoice_Num = dyArray[0].Invoice_Num.ToString();

                    //            if (dyArray[0].Invoice_Status != null)
                    //                Invoice_Status = dyArray[0].Invoice_Status.ToString();

                    //            if (dyArray[0].Invoice_CheckNum != null)
                    //                Invoice_CheckNum = dyArray[0].Invoice_CheckNum.ToString();

                    //            if (dyArray[0].Invoice_De_Psvoucheramt != null)
                    //                Invoice_De_Psvoucheramt = dyArray[0].Invoice_De_Psvoucheramt.ToString();

                    //            if (dyArray[0].De_Rejectcode != null)
                    //                De_Rejectcode = dyArray[0].De_Rejectcode.ToString();

                    //            if (dyArray[0].Rejection_Date != null)
                    //                Rejection_Date = dyArray[0].Rejection_Date.ToString();

                    //            if (dyArray[0].CPR != null)
                    //                CPR = dyArray[0].CPR.ToString();

                    //            if (dyArray[0].Invoice_Due_Date != null)
                    //                Invoice_Due_Date = dyArray[0].Invoice_Due_Date.ToString();

                    //            if (dyArray[0].Invoice_Paid_Date != null)
                    //                Invoice_Paid_Date = dyArray[0].Invoice_Paid_Date.ToString();

                    //            if (dyArray[0].Status != null)
                    //                Status = dyArray[0].Status.ToString();

                    //            if (dyArray[0].Description != null)
                    //                HDescription = dyArray[0].Description.ToString();

                    //            if (dyArray[0].Vendor_Invoice_Num != null)
                    //                Vendor_Invoice_Num = dyArray[0].Vendor_Invoice_Num.ToString();

                    //            if (dyArray[0].Week_Ending != null)
                    //                Week_Ending = dyArray[0].Week_Ending.ToString();

                    //            if (dyArray[0].Vendor_Project_ID != null)
                    //                Vendor_Project_ID = dyArray[0].Vendor_Project_ID.ToString();

                    //            if (dyArray[0].Site != null)
                    //                Site = dyArray[0].Site.ToString();

                    //            if (dyArray[0].CPR_Type != null)
                    //                CPR_Type = dyArray[0].CPR_Type.ToString();

                    //            if (dyArray[0].Reactive_Time_Report_ID != null)
                    //                Reactive_Time_Report_ID = dyArray[0].Reactive_Time_Report_ID.ToString();

                    //            if (dyArray[0].Payment_Type != null)
                    //                Payment_Type = dyArray[0].Payment_Type.ToString();

                    //            if (dyArray[0].Foreman != null)
                    //                Foreman = dyArray[0].Foreman.ToString();

                    //            if (dyArray[0].Derived_Contract != null)
                    //                Derived_Contract = dyArray[0].Derived_Contract.ToString();

                    //            if (dyArray[0]["1st_Approver"] != null)
                    //                First_Approver = dyArray[0]["1st_Approver"].ToString();

                    //            if (dyArray[0]["2nd_Approver"] != null)
                    //                Second_Approver = dyArray[0]["2nd_Approver"].ToString();

                    //            if (dyArray[0].Line_of_Business != null)
                    //                Line_of_Business = dyArray[0].Line_of_Business.ToString();

                    //            if (dyArray[0].Rejection_Remarks != null)
                    //                Rejection_Remarks = dyArray[0].Rejection_Remarks.ToString();

                    //            if (dyArray[0].Requested_By_Name != null)
                    //                Requested_By_Name = dyArray[0].Requested_By_Name.ToString();

                    //            if (dyArray[0].Requested_By_Email != null)
                    //                Requested_By_Email = dyArray[0].Requested_By_Email.ToString();

                    //            if (dyArray[0].Requested_By_Phone_Num != null)
                    //                Requested_By_Phone_Num = dyArray[0].Requested_By_Phone_Num.ToString();

                    //            if (dyArray[0].CPR_Submit_Date != null)
                    //                CPR_Submit_Date = dyArray[0].CPR_Submit_Date.ToString();

                    //            if (dyArray[0].CPR_Created_Date != null)
                    //                CPR_Created_Date = dyArray[0].CPR_Created_Date.ToString();

                    //            if (dyArray[0].Total_Cost != null)
                    //                Total_Cost = dyArray[0].Total_Cost.ToString();

                    //            //CU Details
                    //            if (dyArray[0].CU_Details[i].WO_Num != null)
                    //                WO_Num = dyArray[0].CU_Details[i].WO_Num.ToString();

                    //            if (dyArray[0].CU_Details[i].Description != null)
                    //                Description = dyArray[0].CU_Details[i].Description.ToString();

                    //            if (dyArray[0].CU_Details[i]["Point / Span"] != null)
                    //                Point_Span = dyArray[0].CU_Details[i]["Point / Span"].ToString();

                    //            if (dyArray[0].CU_Details[i].CU_ID != null)
                    //                CU_ID = dyArray[0].CU_Details[i].CU_ID.ToString();

                    //            if (dyArray[0].CU_Details[i].WO_CU_NameNum != null)
                    //                CU_Name = dyArray[0].CU_Details[i].CU_Name.ToString();

                    //            if (dyArray[0].CU_Details[i].CU_Description != null)
                    //                CU_Description = dyArray[0].CU_Details[i].CU_Description.ToString();

                    //            if (dyArray[0].CU_Details[i].Service_Item != null)
                    //                CU_Service_Item = dyArray[0].CU_Details[i].Service_Item.ToString();

                    //            if (dyArray[0].CU_Details[i].Service_Item_Description != null)
                    //                Service_Item_Description = dyArray[0].CU_Details[i].Service_Item_Description.ToString();

                    //            if (dyArray[0].CU_Details[i].Estimated_Qty != null)
                    //                Estimated_Qty = dyArray[0].CU_Details[i].Estimated_Qty.ToString();

                    //            if (dyArray[0].CU_Details[i].Asbuilt_Qty != null)
                    //                Asbuilt_Qty = dyArray[0].CU_Details[i].Asbuilt_Qty.ToString();

                    //            if (dyArray[0].CU_Details[i].Work_Function != null)
                    //                Work_Function = dyArray[0].CU_Details[i].Work_Function.ToString();


                    //            int res = clsDAL.CPR_Data_Create(Approval_Date, Invoice_Num, Invoice_Status, Invoice_CheckNum, Invoice_De_Psvoucheramt, De_Rejectcode, Rejection_Date, CPR, Invoice_Due_Date, Invoice_Paid_Date, Status, HDescription, Vendor_Invoice_Num, Week_Ending, Vendor_Project_ID, Site, CPR_Type, Reactive_Time_Report_ID, Payment_Type, Foreman, Derived_Contract, First_Approver, Second_Approver, Line_of_Business, Rejection_Remarks, Requested_By_Name, Requested_By_Email, Requested_By_Phone_Num, CPR_Submit_Date, CPR_Created_Date, Total_Cost, Line, WO_Task_Num, Service_Item, PrLine_Description, PRLineQuantity, Order_Unit, Unit_Cost, Line_Cost, GL_Debit_Account, WO_Num, Description, Point_Span, CU_ID, CU_Name, CU_Description, CU_Service_Item, Service_Item_Description, Estimated_Qty, Asbuilt_Qty, Work_Function, Contract, SH_Status, SH_change_Date, SH_Changed_By, SH_Rejection_Code);



                    //        }
                    //        catch (Exception exp)
                    //        {
                    //            writer.WriteLine("UnitTrack Maximo GetCPR_CU_Details Failure");
                    //            writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                    //            Console.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                    //            writer.WriteLine("=========================================================================");
                    //        }

                    //    }


                    //    for (int i = 0; i < dyArray[0].Status_History.Count; i++)
                    //    {
                    //        try
                    //        {
                    //            Console.WriteLine("GetCPR_Staus - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray[0].CU_Details.Count.ToString());
                    //            writer.WriteLine("GetCPR_Staus - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray[0].CU_Details.Count.ToString());

                    //            if (dyArray[0].Approval_Date != null)
                    //                Approval_Date = dyArray[0].Approval_Date.ToString();

                    //            if (dyArray[0].Invoice_Num != null)
                    //                Invoice_Num = dyArray[0].Invoice_Num.ToString();

                    //            if (dyArray[0].Invoice_Status != null)
                    //                Invoice_Status = dyArray[0].Invoice_Status.ToString();

                    //            if (dyArray[0].Invoice_CheckNum != null)
                    //                Invoice_CheckNum = dyArray[0].Invoice_CheckNum.ToString();

                    //            if (dyArray[0].Invoice_De_Psvoucheramt != null)
                    //                Invoice_De_Psvoucheramt = dyArray[0].Invoice_De_Psvoucheramt.ToString();

                    //            if (dyArray[0].De_Rejectcode != null)
                    //                De_Rejectcode = dyArray[0].De_Rejectcode.ToString();

                    //            if (dyArray[0].Rejection_Date != null)
                    //                Rejection_Date = dyArray[0].Rejection_Date.ToString();

                    //            if (dyArray[0].CPR != null)
                    //                CPR = dyArray[0].CPR.ToString();

                    //            if (dyArray[0].Invoice_Due_Date != null)
                    //                Invoice_Due_Date = dyArray[0].Invoice_Due_Date.ToString();

                    //            if (dyArray[0].Invoice_Paid_Date != null)
                    //                Invoice_Paid_Date = dyArray[0].Invoice_Paid_Date.ToString();

                    //            if (dyArray[0].Status != null)
                    //                Status = dyArray[0].Status.ToString();

                    //            if (dyArray[0].Description != null)
                    //                HDescription = dyArray[0].Description.ToString();

                    //            if (dyArray[0].Vendor_Invoice_Num != null)
                    //                Vendor_Invoice_Num = dyArray[0].Vendor_Invoice_Num.ToString();

                    //            if (dyArray[0].Week_Ending != null)
                    //                Week_Ending = dyArray[0].Week_Ending.ToString();

                    //            if (dyArray[0].Vendor_Project_ID != null)
                    //                Vendor_Project_ID = dyArray[0].Vendor_Project_ID.ToString();

                    //            if (dyArray[0].Site != null)
                    //                Site = dyArray[0].Site.ToString();

                    //            if (dyArray[0].CPR_Type != null)
                    //                CPR_Type = dyArray[0].CPR_Type.ToString();

                    //            if (dyArray[0].Reactive_Time_Report_ID != null)
                    //                Reactive_Time_Report_ID = dyArray[0].Reactive_Time_Report_ID.ToString();

                    //            if (dyArray[0].Payment_Type != null)
                    //                Payment_Type = dyArray[0].Payment_Type.ToString();

                    //            if (dyArray[0].Foreman != null)
                    //                Foreman = dyArray[0].Foreman.ToString();

                    //            if (dyArray[0].Derived_Contract != null)
                    //                Derived_Contract = dyArray[0].Derived_Contract.ToString();

                    //            if (dyArray[0]["1st_Approver"] != null)
                    //                First_Approver = dyArray[0]["1st_Approver"].ToString();

                    //            if (dyArray[0]["2nd_Approver"] != null)
                    //                Second_Approver = dyArray[0]["2nd_Approver"].ToString();

                    //            if (dyArray[0].Line_of_Business != null)
                    //                Line_of_Business = dyArray[0].Line_of_Business.ToString();

                    //            if (dyArray[0].Rejection_Remarks != null)
                    //                Rejection_Remarks = dyArray[0].Rejection_Remarks.ToString();

                    //            if (dyArray[0].Requested_By_Name != null)
                    //                Requested_By_Name = dyArray[0].Requested_By_Name.ToString();

                    //            if (dyArray[0].Requested_By_Email != null)
                    //                Requested_By_Email = dyArray[0].Requested_By_Email.ToString();

                    //            if (dyArray[0].Requested_By_Phone_Num != null)
                    //                Requested_By_Phone_Num = dyArray[0].Requested_By_Phone_Num.ToString();

                    //            if (dyArray[0].CPR_Submit_Date != null)
                    //                CPR_Submit_Date = dyArray[0].CPR_Submit_Date.ToString();

                    //            if (dyArray[0].CPR_Created_Date != null)
                    //                CPR_Created_Date = dyArray[0].CPR_Created_Date.ToString();

                    //            if (dyArray[0].Total_Cost != null)
                    //                Total_Cost = dyArray[0].Total_Cost.ToString();

                    //            //Status History
                    //            if (dyArray[0].Status_History[i].Status != null)
                    //                SH_Status = dyArray[0].Status_History[i].Status.ToString();

                    //            if (dyArray[0].Status_History[i].Change_Date != null)
                    //                SH_change_Date = dyArray[0].Status_History[i].Change_Date.ToString();

                    //            if (dyArray[0].Status_History[i].Changed_By != null)
                    //                SH_Changed_By = dyArray[0].Status_History[i].Changed_By.ToString();

                    //            if (dyArray[0].Status_History[i].Rejection_Code != null)
                    //                SH_Rejection_Code = dyArray[0].Status_History[i].Rejection_Code.ToString();

                    //            int res = clsDAL.CPR_Data_Create(Approval_Date, Invoice_Num, Invoice_Status, Invoice_CheckNum, Invoice_De_Psvoucheramt, De_Rejectcode, Rejection_Date, CPR, Invoice_Due_Date, Invoice_Paid_Date, Status, HDescription, Vendor_Invoice_Num, Week_Ending, Vendor_Project_ID, Site, CPR_Type, Reactive_Time_Report_ID, Payment_Type, Foreman, Derived_Contract, First_Approver, Second_Approver, Line_of_Business, Rejection_Remarks, Requested_By_Name, Requested_By_Email, Requested_By_Phone_Num, CPR_Submit_Date, CPR_Created_Date, Total_Cost, Line, WO_Task_Num, Service_Item, PrLine_Description, PRLineQuantity, Order_Unit, Unit_Cost, Line_Cost, GL_Debit_Account, WO_Num, Description, Point_Span, CU_ID, CU_Name, CU_Description, CU_Service_Item, Service_Item_Description, Estimated_Qty, Asbuilt_Qty, Work_Function, Contract, SH_Status, SH_change_Date, SH_Changed_By, SH_Rejection_Code);
                    //        }
                    //        catch (Exception exp)
                    //        {
                    //            writer.WriteLine("UnitTrack Maximo GetCPR_Staus_Details Failure");
                    //            writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                    //            Console.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                    //            writer.WriteLine("=========================================================================");
                    //        }

                    //    }
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

        #region PPM_Push_Oracle
        public static void PPM_Push_Oracle(StreamWriter writer)
        {
            writer.WriteLine("DynamicsPikeService (Create) - " + AppName + " - Started");
            Console.WriteLine("DynamicsPikeService - " + AppName + " - PPM_Push_Oracle - Started");

            string TransactionNumber = "";
            string UnprocessTransactionId = "";

            string DetailRecordId = "";
            string IntegrationBatchHeaderId = "";
            string ExpenditureBatchName = "";

            DataSet dsDetails = clsDAL.SQL_Oracle_PPM_GetList();

            try
            {
                if (dsDetails.Tables[0].Rows.Count > 0)
                {
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

                            if (PPMresponse.StatusCode == System.Net.HttpStatusCode.OK)
                            {
                                //Update the PPM Status
                                string Result = PPMresponse.Content!.ToString();
                                dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(PPMresponse.Content!)!;

                                if (dyArray.count != 0)
                                    TransactionNumber = dyArray.items[0].TransactionNumber.ToString();

                                writer.WriteLine("TransactionNumber : " + TransactionNumber);
                                Console.WriteLine("TransactionNumber : " + TransactionNumber);

                                if (TransactionNumber != null && TransactionNumber != "" && TransactionNumber != "0")
                                {
                                    RecordFlag = 0;

                                    writer.WriteLine("Updating the Details Record with the Transaction Number into Dynamics for Detail ID= " + DetailRecordId);

                                    oracle_status_details = "Success";
                                    oracle_message_details = "Transaction Number Updated";

                                    clsDAL.SQL_Duke_NLR_DataUpdate(TransactionNumber, oracle_status_details, oracle_message_details, UnprocessTransactionId, DetailRecordId);
                                }
                            }

                            #endregion

                            if (RecordFlag == 1)
                            {

                                Console.WriteLine("Processing Unit Billing Detail Events Records - " + j.ToString() + " - out of " + dsDetails.Tables[0].Rows.Count.ToString());
                                writer.WriteLine("Processing Unit Billing Detail Events Records - " + j.ToString() + " - out of " + dsDetails.Tables[0].Rows.Count.ToString());


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


                                    clsDAL.SQL_Duke_NLR_DataUpdate("", oracle_status_details, oracle_message_details, UnprocessTransactionId, DetailRecordId);
                                }
                                else
                                {
                                    writer.WriteLine("Oracle Response for Detail Transaction Record :  " + response.Content!);
                                    writer.WriteLine("Updating the Detail Record Status in Dynamics along with other values for " + DetailRecordId);
                                    oracle_status_details = "Error";
                                    oracle_message_details = response.Content!.ToString();
                                    //Update the PPM Status      

                                    clsDAL.SQL_Duke_NLR_DataUpdate("", oracle_status_details, oracle_message_details, "", DetailRecordId);
                                }
                            }
                        }
                        catch (Exception exp)
                        {
                            writer.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);
                            Console.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);

                            oracle_status_details = "Error";
                            oracle_message_details = exp.Message.ToString();

                            clsDAL.SQL_Duke_NLR_DataUpdate("", oracle_status_details, oracle_message_details, "", DetailRecordId);
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

        #region PC_Update_Oracle
        public static void PC_Update_Oracle(StreamWriter writer)
        {
            writer.WriteLine("DynamicsPikeService (Update) - " + AppName + " - Started");           
            Console.WriteLine("DynamicsPikeService - " + AppName + " - UpdateDataTo_Oracle - Started");

            string TransactionNumber = "";
            string UnprocessTransactionId = "";

            string DetailRecordId = "";
            string IntegrationBatchHeaderId = "";
            string ExpenditureBatchName = "";

            DataSet dsDetails = clsDAL.SQL_Oracle_PC_GetList();

            try
            {
                if (dsDetails.Tables[0].Rows.Count > 0)
                {
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

                            if (PPMresponse.StatusCode == System.Net.HttpStatusCode.OK)
                            {
                                //Update the PPM Status
                                string Result = PPMresponse.Content!.ToString();
                                dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(PPMresponse.Content!)!;
                              
                                if (dyArray.count != 0)
                                    TransactionNumber = dyArray.items[0].TransactionNumber.ToString();

                                writer.WriteLine("TransactionNumber : " + TransactionNumber);
                                Console.WriteLine("TransactionNumber : " + TransactionNumber);

                                if (TransactionNumber != null && TransactionNumber != "" && TransactionNumber != "0")
                                {
                                    writer.WriteLine("Updating the Details Record with the Transaction Number into Dynamics for Detail ID= " + DetailRecordId);

                                    oracle_status_details = "Success";
                                    oracle_message_details = "Transaction Number Updated";                                   

                                    clsDAL.SQL_Duke_NLR_DataUpdate(TransactionNumber, oracle_status_details, oracle_message_details, UnprocessTransactionId, DetailRecordId);
                                }
                            }

                            #endregion
                           
                        }
                        catch (Exception exp)
                        {
                            writer.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);
                            Console.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);

                            oracle_status_details = "Error";
                            oracle_message_details = exp.Message.ToString();
                        }
                    }
                }

            }
            catch (Exception exp)
            {
                writer.WriteLine("Application Exception " + exp.Message);
                Console.WriteLine("Application Exception " + exp.Message);
            }

            writer.WriteLine("DynamicsPikeService (Update) - " + AppName + " - Completed");

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