using Newtonsoft.Json;
using System.Configuration;
using System.Data;
using System.Text;
using System.Web;
using RestSharp;
using Dynamics_Oracle_UnitBilling;
using static System.Net.Mime.MediaTypeNames;
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
        static readonly int ExceptionOnScreen = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ExceptionOnScreen"]!.ToString());
        public static int RunGenerateTimecardProcess = 0;
        static int ESSJob = 0;

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

                GetWorkorders_DYN(Writer);

                GetDukeCompatibleUnits(Writer);

                UpdateServiceItemData(Writer);

                GetSubTaskNumber_Oracle(Writer);

                GetDukeProjectTaskData(Writer);

                //GetNLRData_Oracle();

                //PushDataTo_Oracle();



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

                DataSet dsProjectTask = clsDAL.Dynamics_ProjectTask_GetList();
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

                            Parent_Task_Number = dsProjectTask.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;
                            Parent_Task_Id = dsProjectTask.Tables[0].Rows[i]["Parent_Task_Id"].ToString()!;
                            Project_Number = dsProjectTask.Tables[0].Rows[i]["Project_Number"].ToString()!;
                            Project_Id = dsProjectTask.Tables[0].Rows[i]["Project_Id"].ToString()!;
                            BusinessUnitName = dsProjectTask.Tables[0].Rows[i]["BusinessUnitName"].ToString()!;
                            ProcessDate = DateTime.Now.ToString("yyyy-MM-dd");

                            int res = clsDAL.SQL_ProjectTask_Create(Parent_Task_Number, Parent_Task_Id, Project_Number, Project_Id, BusinessUnitName, ProcessDate);
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
                Console.WriteLine("GetWorkorders_DYN - Processing Completed");
                writer.WriteLine("GetWorkorders_DYN - Processing Completed");
            }

        }
        #endregion

        #region GetDukeCompatibleUnits
        public static void GetDukeCompatibleUnits(StreamWriter writer)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeCompatibleUnits - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeCompatibleUnits - Started");

                DataSet ds = clsDAL.SQL_ProjectTask_GetList(1);

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetDukeCompatibleUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeCompatibleUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        int WorkOrder_Id = 0;
                        string WorkOrder_Number = "";

                        WorkOrder_Id= Convert.ToInt32(ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString());
                        WorkOrder_Number = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;

                        GetDukeCompatibleUnits_Details(WorkOrder_Id, WorkOrder_Number, writer);
                      
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
            catch (Exception exp)
            {
                Console.WriteLine(exp.Message.ToString());
            }
        }
        #endregion

        #region GetDukeCompatibleUnits_Details
        public static string GetDukeCompatibleUnits_Details(int WorkOrder_Id, string WorkOrder_Number, StreamWriter writer)
        {           
            string token = GetToken();
            string result = "";
            dynamic dyArray = "";


            string ServiceUrl = System.Configuration.ConfigurationManager.AppSettings["ServiceUrl"]!.ToString();
            string ServiceSubUrl = System.Configuration.ConfigurationManager.AppSettings["ServiceSubUrl"]!.ToString();

            try
            {


                #region Getting Compatible Units with token
                //if (token != null)
                //{
                //    var options = new RestClientOptions(OauthUrl)
                //    {
                //        MaxTimeout = -1,
                //    };
                //    var client = new RestClient(options);

                //    var request = new RestRequest(OauthSubUrl + "workorders/tasks/" + WorkOrder_Number + "/compatible-units", Method.Get);

                //    request.AddHeader("Authorization", "Bearer " + token);
                //    RestResponse response = client.Execute(request);
                //    //Console.WriteLine(response.Content);

                //    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                //        dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!.ToString())!;
                //    else
                //        return result;

                //    result = dyArray.ToString();

                //    string Station_Detail_ID = "";
                //    string Payment_Status = "";
                //    string Work_Function = "";
                //    string Estimate_Quantity = "";
                //    string Compatible_Unit_Quantity = "";
                //    string Completed_Date = "";
                //    string Compatible_Unit_Changed_By = "";
                //    string Compatible_Unit_Changed_Date = "";
                //    string Compatible_Unit_Station = "";
                //    string Parent_Compatible_Unit_Name = "";
                //    string Parent_Compatible_Unit_Description = "";
                //    string Payment_Type = "";
                //    string Field_To_From = "";
                //    string Field_ID_To = "";
                //    string Serial_Number = "";
                //    string Manufacturer = "";
                //    string Vendor_Id = "";

                //    int Count = 1;
                //    int RecordCount = dyArray.responseList.Count;

                //    if (dyArray.responseList.Count != null)
                //    {
                //        for (int i = 0; i < dyArray.responseList.Count; i++)
                //        {
                //            try
                //            {
                //                Console.WriteLine("GetDukeCompatibleUnits - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray.responseList.Count.ToString());
                //                writer.WriteLine("GetDukeCompatibleUnits - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray.responseList.Count.ToString());

                //                if (dyArray.responseList[i].Station_Detail_ID != null)
                //                    Station_Detail_ID = dyArray.responseList[i].Station_Detail_ID;

                //                if (dyArray.responseList[i].Payment_Status != null)
                //                    Payment_Status = dyArray.responseList[i].Payment_Status;

                //                if (dyArray.responseList[i].Work_Function != null)
                //                    Work_Function = dyArray.responseList[i].Work_Function;

                //                if (dyArray.responseList[i].Estimate_Quantity != null)
                //                    Estimate_Quantity = dyArray.responseList[i].Estimate_Quantity;

                //                if (dyArray.responseList[i].Compatible_Unit_Quantity != null)
                //                    Compatible_Unit_Quantity = dyArray.responseList[i].Compatible_Unit_Quantity;

                //                if (dyArray.responseList[i].Completed_Date != null)
                //                    Completed_Date = dyArray.responseList[i].Completed_Date;

                //                if (dyArray.responseList[i].Compatible_Unit_Changed_By != null)
                //                    Compatible_Unit_Changed_By = dyArray.responseList[i].Compatible_Unit_Changed_By;

                //                if (dyArray.responseList[i].Compatible_Unit_Changed_Date != null)
                //                    Compatible_Unit_Changed_Date = dyArray.responseList[i].Compatible_Unit_Changed_Date;

                //                if (dyArray.responseList[i].Compatible_Unit_Station != null)
                //                    Compatible_Unit_Station = dyArray.responseList[i].Compatible_Unit_Station;

                //                if (dyArray.responseList[i].Parent_Compatible_Unit_Name != null)
                //                    Parent_Compatible_Unit_Name = dyArray.responseList[i].Parent_Compatible_Unit_Name;

                //                if (dyArray.responseList[i].Parent_Compatible_Unit_Description != null)
                //                    Parent_Compatible_Unit_Description = dyArray.responseList[i].Parent_Compatible_Unit_Description;

                //                if (dyArray.responseList[i].Payment_Type != null)
                //                    Payment_Type = dyArray.responseList[i].Payment_Type;

                //                if (dyArray.responseList[i].Field_To_From != null)
                //                    Field_To_From = dyArray.responseList[i].Field_To_From;

                //                if (dyArray.responseList[i].Field_ID_To != null)
                //                    Field_ID_To = dyArray.responseList[i].Field_ID_To;

                //                if (dyArray.responseList[i].Serial_Number != null)
                //                    Serial_Number = dyArray.responseList[i].Serial_Number;

                //                if (dyArray.responseList[i].Manufacturer != null)
                //                    Manufacturer = dyArray.responseList[i].Manufacturer;

                //                if (dyArray.responseList[i].Vendor_Id != null)
                //                    Vendor_Id = dyArray.responseList[i].Vendor_Id;

                //                int res = clsDAL.SQL_MaximoUnits_Create(Station_Detail_ID, Payment_Status, Work_Function, Estimate_Quantity, Compatible_Unit_Quantity, Completed_Date, Compatible_Unit_Changed_By, Compatible_Unit_Changed_Date, Compatible_Unit_Station, Parent_Compatible_Unit_Name, Parent_Compatible_Unit_Description, Payment_Type, Field_To_From, Field_ID_To, Serial_Number, Manufacturer, Vendor_Id, WorkOrder_Id);

                //                if (Count == RecordCount)
                //                {
                //                    clsDAL.SQL_ProjectTask_StatusUpdate(WorkOrder_Id);
                //                }

                //                Count++;

                //            }
                //            catch (Exception exp)
                //            {
                //                writer.WriteLine("UnitTrack Maximo GetDukeCompatibleUnits Failure");
                //                writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                //                writer.WriteLine("=========================================================================");
                //            }

                //        }
                //    }


                //} 
                #endregion


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
                                Console.WriteLine("GetDukeCompatibleUnits - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray.responseList.Count.ToString());
                                writer.WriteLine("GetDukeCompatibleUnits - Details - Processing - " + (i + 1).ToString() + " out of " + dyArray.responseList.Count.ToString());

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

                                int res = clsDAL.SQL_MaximoUnits_Create(Station_Detail_ID, Payment_Status, Work_Function, Estimate_Quantity, Compatible_Unit_Quantity, Completed_Date, Compatible_Unit_Changed_By, Compatible_Unit_Changed_Date, Compatible_Unit_Station, Parent_Compatible_Unit_Name, Parent_Compatible_Unit_Description, Payment_Type, Field_To_From, Field_ID_To, Serial_Number, Manufacturer, Vendor_Id, WorkOrder_Id);

                                if (Count == RecordCount)
                                {
                                    clsDAL.SQL_ProjectTask_StatusUpdate(WorkOrder_Id);
                                }

                                Count++;

                            }
                            catch (Exception exp)
                            {
                                writer.WriteLine("UnitTrack Maximo GetDukeCompatibleUnits Failure");
                                writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
                                writer.WriteLine("=========================================================================");
                            }

                        }
                    }

                }
            }
            catch (Exception exp)
            {

                writer.WriteLine("UnitTrack Maximo DukeCompatibleUnits Failure");
                writer.WriteLine("UnitTrack Maximo Message :  " + exp.Message.ToString());
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
            catch (Exception exp) { }
        }
        #endregion

        #region GetSubTaskNumber_Oracle
        public static void GetSubTaskNumber_Oracle(StreamWriter writer)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetSubTaskNumber_Oracle - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetSubTaskNumber_Oracle - Started");

                DataSet ds = clsDAL.SQL_ProjectTask_GetList(2);

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetDukeCompatibleUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeCompatibleUnits - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        string WorkOrder_Id = ds.Tables[0].Rows[i]["WorkOrder_Id"].ToString()!;
                        string ProjectNumber = ds.Tables[0].Rows[i]["Project_Number"].ToString()!;
                        string TaskNumber = ds.Tables[0].Rows[i]["Parent_Task_Number"].ToString()!;

                        string OracleWrapperUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperUrl"]!.ToString();
                        string OracleWrapperSubUrl = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperSubUrl"]!.ToString();
                        string OracleWrapperEnvironment = System.Configuration.ConfigurationManager.AppSettings["OracleWrapperEnvironment"]!.ToString();

                        var configuration = new ConfigurationBuilder()
                           .AddJsonFile("DBQueries.json", optional: false, reloadOnChange: true)
                           .AddEnvironmentVariables()
                           .Build();
                        string cmd = configuration.GetSection("qrySubTaskData").Value!.Replace("@ProjectNumber", "'" + ProjectNumber + "'").Replace("@TaskNumber", "'" + TaskNumber + "'");

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

                        clsDAL.Update_SubTaskDetails(WorkOrder_Id, Sub_Task_Number, Sub_Task_Name, Sub_Task_Id, Sub_Task_Billable_Flag, Sub_Task_Crew_Leader, Sub_Task_Project_Flag);
                    }
                    catch (Exception ex)
                    {
                        int j = 0;
                    }
                }


                    
                }
                catch (Exception ex2) {
                int k = 0;
            }
        }
        #endregion

        #region GetDukeProjectTaskData
        public static void GetDukeProjectTaskData(StreamWriter writer)
        {
            try
            {
                writer.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeProjectTaskData - Started");
                Console.WriteLine("DynamicsPikeService - " + AppName + " - GetDukeProjectTaskData - Started");


                DataSet ds = clsDAL.SQL_NLR_GetList();

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine("GetDukeProjectTaskData - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());
                        writer.WriteLine("GetDukeProjectTaskData - Processing - " + (i + 1).ToString() + " out of " + ds.Tables[0].Rows.Count.ToString());

                        string Compatible_Unit_Id = ds.Tables[0].Rows[i]["Compatible_Unit_Id"].ToString()!;
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
                        string cmd = configuration.GetSection("qryProjectTaskData").Value!.Replace("@Project", "'" + @Project + "'").Replace("@PARENT_TASKNO", "'" + @PARENT_TASKNO + "'").Replace("@SUBTASK", "'" + @SUBTASK + "'").Replace("@SERVICE_ITEM", "'" + @SERVICE_ITEM + "'").Replace("@Quantity", "'" + @Quantity + "'").Replace("@EffectiveDate", "'" + @EffectiveDate + "'");


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

                        int res = clsDAL.Oracle_ProjectTask_NLR_Create(PARAM_PROJ_TO_BILL, PARAM_WONO_TO_BILL, PARAM_SUBTASK_TO_BILL, PARAM_UNIT_TO_BILL, PARAM_QUANTITY, PARAM_WE_DATE, ERROR_CODES, ORACLE_TASK_NO, ORACLE_SUBTASK, PARENT_BILLABLE_FLAG, SUBTASK_BILLABLE_CHARGEABLE_FLAG, RATE_SCHEDUL_ERROR, PROJ_TASK_DETAILS, PROJECT_NUMBER, PROJECT_NAME, SUBTASK_PROJECT, PARENT_WO_NUMBER, PARENT_WO_NAME, PARENT_BILLABLE, PARENT_CHARGEABLE, TOP_TASK_ID, SUBTASK_WO_NUMBER, SUBTASK_WO_NAME, SUBTASK_WO_BILLABLE, SUBTASK_WO_CHARGEABLE, CREW_LEADER, RATE_SCHEDULE_DETAILS, EXP_NAME, RATE_SCHEDULE_NAME, UNIT_NAME, RATE, UNIT_OF_MEASURE, RATE_START_DATE, RATE_END_DATE, FBDILOADER, EXPENDITUREDATE, PERSONNAME, PERSONNUMBER, HUMANRESOURCEASSIGNMENT, PROJECTNAME, PROJECTNUMBER, TASK_NAME, TASK_NUMBER, EXPENDITURETYPE, EXPENDITUREORGANIZATION, CONTRACTNUMBER, FUNDINGSOURCENUMBER, NONLABORRESOURCE, NONLABORRESOURCEORGANIZATION, QUANTITY, WORKTYPE, ADDITIONALINFO, PROJID, NLRID, TASKID, EXISTINGQTYINPPM, REGION, LEGALENTITY, BUSINESSUNIT, Compatible_Unit_Id);
                    }
                    catch (Exception e)
                    {

                        int j = 0;
                    }
                }
            }
            catch(Exception exp)
            {
                int k = 0;
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

    }


}