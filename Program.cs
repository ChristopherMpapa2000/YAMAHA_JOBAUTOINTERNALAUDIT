using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using System.Reflection.Emit;
using WolfApprove.Model;
using System.Data.Linq;
using WolfApprove.Model.CustomClass;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Web.Script.Serialization;
using System.Data;
using WolfApprove.Model.Extension;
using WolfApprove.Model.Migrations;
using static WolfApprove.Model.CustomClass.CustomJsonAdvanceForm;
using System.Globalization;
using static System.Net.Mime.MediaTypeNames;
using JobOrder;

namespace JobAutoInternalAudit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("---------- Start JobAutoInternalAudit at :: " + DateTime.Now + " ----------");
            WriteLogFile("---------- Start JobAutoInternalAudit at :: " + DateTime.Now + " ----------");
            Console.WriteLine("---------- Start GetStartForm at :: " + DateTime.Now + " ----------");
            WriteLogFile("---------- Start GetStartForm at :: " + DateTime.Now + " ----------");
            GetStartForm();
            WriteLogFile("---------- Finish GetStartForm at :: " + DateTime.Now + " ----------");
            Console.WriteLine("---------- Finish GetStartForm at :: " + DateTime.Now + " ----------");

            Console.WriteLine("---------- Start GetStartFormCAR at :: " + DateTime.Now + " ----------");
            WriteLogFile("---------- Start GetStartFormCAR at :: " + DateTime.Now + " ----------");
            GetStartFormCAR();
            WriteLogFile("---------- Finish GetStartFormCAR at :: " + DateTime.Now + " ----------");
            Console.WriteLine("---------- Finish GetStartFormCAR at :: " + DateTime.Now + " ----------");
            WriteLogFile("---------- Finish JobAutoInternalAudit at :: " + DateTime.Now + " ----------");
            Console.WriteLine("---------- Finish JobAutoInternalAudit at :: " + DateTime.Now + " ----------");
        }

        private static string dbConnectionStringWolf
        {
            get
            {
                var dbConnectionString = System.Configuration.ConfigurationSettings.AppSettings["dbConnectionStringWolf"];
                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return "Integrated Security=SSPI;Initial Catalog=WolfApproveCore.tym-ypmt;Data Source=DESKTOP-SHLJS2M\\SQLEXPRESS2019;User Id=sa;Password=pass@word1;";
            }
        }
        private static string _LogFile
        {
            get
            {
                var LogFile = System.Configuration.ConfigurationSettings.AppSettings["LogFile"];
                if (!string.IsNullOrEmpty(LogFile))
                {
                    return (LogFile);
                }
                return string.Empty;
            }
        }
        private static int iIntervalTime
        {
            //ตั้งค่าเวลา
            get
            {
                var IntervalTime = System.Configuration.ConfigurationSettings.AppSettings["IntervalTimeMinute"];
                if (!string.IsNullOrEmpty(IntervalTime))
                {
                    return Convert.ToInt32(IntervalTime);
                }
                return -10;
            }
        }
        public static void GetStartForm()
        {
            DataClasses1DataContext dbWolf = new DataClasses1DataContext(dbConnectionStringWolf);
            DateTime lastRunTime = DateTime.Now.AddMinutes(-5);
            TRNMemo chtchmemo = new TRNMemo();

            var Getmemo = dbWolf.TRNMemos.Where(x => x.TemplateId == 99 && x.StatusName == "Completed" && x.CompletedDate >= DateTime.Now.AddMonths(iIntervalTime)).OrderBy(t => t.CompletedDate).ToList();
            var trnmemoform = dbWolf.TRNMemoForms.Where(z => z.TemplateId == 99 && z.obj_label == "Area").ToList();
            var joinedData = (from memo in Getmemo
                              join form in trnmemoform
                              on memo.MemoId equals form.MemoId
                              group new { memo, form.obj_value } by form.obj_value into grouped
                              select new
                              {
                                  Memo = grouped.OrderByDescending(m => m.memo.ModifiedDate).FirstOrDefault().memo,
                                  ObjValue = grouped.Key
                              }).ToList();
            try
            {
                if (_runcatch == "T")
                {
                    var memo = Int32.Parse(_memoid);
                    var GetmemoCom = dbWolf.TRNMemos
                       .Where(x => x.MemoId == memo).ToList();

                    if (GetmemoCom.Count > 0)
                    {
                        var Gettemplate = dbWolf.MSTTemplates.Where(x => x.TemplateId == 8).FirstOrDefault();
                        if (Gettemplate != null)
                        {
                            foreach (var template in GetmemoCom)
                            {
                                chtchmemo.MemoId = template.MemoId;
                                string addvaluetoadvanceform = template.MAdvancveForm;

                                var jsonObject = JObject.Parse(addvaluetoadvanceform);

                                var nonTableData = new List<KeyValuePair<string, string>>();
                                var tableData = new List<List<KeyValuePair<string, string>>>();

                                foreach (var item in jsonObject["items"])
                                {
                                    foreach (var layout in item["layout"])
                                    {
                                        var templateType = layout["template"]["type"]?.ToString();
                                        if (templateType == "tb") // Checking if it's a table type
                                        {
                                            string tableLabel = (string)layout["template"]?["label"];
                                            JObject attribute = (JObject)layout["template"]?["attribute"];
                                            JArray columns = (JArray)attribute?["column"];

                                            JArray rows = (JArray)layout["data"]?["row"];
                                            if (rows != null)
                                            {
                                                foreach (JArray row in rows)
                                                {
                                                    // Create a list to hold the label-value pairs for this row.
                                                    var labelValuePairs = new List<KeyValuePair<string, string>>();

                                                    for (int i = 0; i < row.Count; i++)
                                                    {
                                                        JObject cell = (JObject)row[i];
                                                        string value = (string)cell["value"];
                                                        if (columns != null && i < columns.Count)
                                                        {
                                                            JObject column = (JObject)columns[i];
                                                            string columnLabel = (string)column["label"];
                                                            if (!string.IsNullOrEmpty(columnLabel) && !string.IsNullOrEmpty(value))
                                                            {
                                                                labelValuePairs.Add(new KeyValuePair<string, string>(columnLabel, value));
                                                            }
                                                        }
                                                    }

                                                    // Add the list of label-value pairs for this row to the table data.
                                                    tableData.Add(labelValuePairs);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var label = layout["template"]["label"]?.ToString();
                                            var value = layout["data"]["value"]?.ToString();

                                            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(value))
                                            {
                                                nonTableData.Add(new KeyValuePair<string, string>(label, value));

                                            }
                                        }
                                    }
                                }

                                string Getadvancetem = Gettemplate.AdvanceForm;
                                var jsonObject2 = JObject.Parse(Getadvancetem);

                                var labels = new List<string>();

                                // เข้าถึง array 'items'
                                if (jsonObject2["items"] is JArray itemsArray)
                                {
                                    foreach (var item in itemsArray)
                                    {
                                        // เข้าถึง array 'layout' ภายในแต่ละ 'item'
                                        if (item["layout"] is JArray layoutArray)
                                        {
                                            foreach (var layout in layoutArray)
                                            {
                                                // ตรวจสอบและเพิ่ม label ลงในรายการ
                                                var label = layout["template"]?["label"]?.ToString();
                                                if (!string.IsNullOrEmpty(label))
                                                {
                                                    labels.Add(label);
                                                }
                                            }
                                        }
                                    }
                                }
                                foreach (var loopdata in tableData)
                                {
                                    var Arae = loopdata[1].Value;
                                    var SeleteMemo = joinedData.Where(a => a.ObjValue == Arae).FirstOrDefault();
                                    List<object> listObjectDCCArea = new List<object>();
                                    List<object> listObjectworkingteam = new List<object>();
                                    JObject jsonAdvanceForm = JsonUtils.createJsonObject(SeleteMemo.Memo.MAdvancveForm);
                                    JArray itemssArray = (JArray)jsonAdvanceForm["items"];
                                    foreach (JObject jItems in itemssArray)
                                    {
                                        JArray jLayoutArray = (JArray)jItems["layout"];
                                        if (jLayoutArray.Count >= 1)
                                        {
                                            JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                            JObject jData = (JObject)jLayoutArray[0]["data"];
                                            if ((String)jTemplateL["label"] == "Responsibilities : DCC Area")
                                            {

                                                foreach (JArray row in jData["row"])
                                                {
                                                    List<object> rowObject = new List<object>();
                                                    foreach (JObject item in row)
                                                    {
                                                        rowObject.Add(item["value"].ToString());
                                                    }
                                                    listObjectDCCArea.Add(rowObject);
                                                }
                                            }
                                            if ((String)jTemplateL["label"] == "Responsibilities : Working team")
                                            {

                                                foreach (JArray row in jData["row"])
                                                {
                                                    List<object> rowObject = new List<object>();
                                                    foreach (JObject item in row)
                                                    {
                                                        rowObject.Add(item["value"].ToString());
                                                    }
                                                    listObjectworkingteam.Add(rowObject);
                                                }
                                            }
                                        }
                                    }

                                    var getempcodere = dbWolf.MSTEmployees.Where(x => x.EmployeeId == template.RequesterId).FirstOrDefault();

                                    Getadvancetem = Gettemplate.AdvanceForm;
                                    WriteLogFile("getempcodere.EmployeeCode" + getempcodere.EmployeeCode);
                                    WriteLogFile("dbConnectionStringWolf" + dbConnectionStringWolf);
                                    var EmpCurrent = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == getempcodere.EmployeeCode).FirstOrDefault();
                                    var CurrentCom = dbWolf.MSTCompanies.Where(a => a.CompanyCode == EmpCurrent.CompanyCode).FirstOrDefault();
                                    TRNMemo objMemo = new TRNMemo();
                                    objMemo.StatusName = "Draft";
                                    objMemo.CreatedDate = DateTime.Now;
                                    objMemo.CreatedBy = EmpCurrent.NameEn;
                                    objMemo.CreatorId = EmpCurrent.EmployeeId;
                                    objMemo.RequesterId = EmpCurrent.EmployeeId;
                                    objMemo.CNameTh = EmpCurrent.NameTh;
                                    objMemo.CNameEn = EmpCurrent.NameEn;
                                    objMemo.CPositionId = EmpCurrent.PositionId;
                                    objMemo.CPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.CPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.CDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.CDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.CDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.RNameTh = EmpCurrent.NameTh;
                                    objMemo.RNameEn = EmpCurrent.NameEn;
                                    objMemo.RPositionId = EmpCurrent.PositionId;
                                    objMemo.RPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.RPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.RDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.RDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.RDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.ModifiedDate = DateTime.Now;
                                    objMemo.ModifiedBy = objMemo.ModifiedBy;
                                    objMemo.TemplateId = Gettemplate.TemplateId;
                                    objMemo.TemplateName = Gettemplate.TemplateName;
                                    objMemo.GroupTemplateName = Gettemplate.GroupTemplateName;
                                    objMemo.RequestDate = DateTime.Now;
                                    objMemo.PersonWaitingId = EmpCurrent.EmployeeId;
                                    objMemo.PersonWaiting = EmpCurrent.NameEn;

                                    objMemo.CompanyId = CurrentCom.CompanyId;
                                    objMemo.CompanyName = CurrentCom.NameTh;


                                    objMemo.TAdvanceForm = Gettemplate.AdvanceForm;
                                    objMemo.TemplateSubject = Gettemplate.TemplateSubject;
                                    objMemo.TemplateDetail = Guid.NewGuid().ToString().Replace("-", "");
                                    objMemo.ToPerson = Gettemplate.ToId;
                                    objMemo.CcPerson = Gettemplate.CcId;
                                    objMemo.CurrentApprovalLevel = null;
                                    objMemo.ProjectID = 0;
                                    objMemo.Amount = 0;
                                    objMemo.DocumentCode = GenControlRunning(EmpCurrent, Gettemplate.DocumentCode, objMemo, dbWolf);

                                    objMemo.MemoSubject = "รายการตรวจติดตามภายในองค์กร (Internal Audit Checklist) : " + loopdata[1].Value + " " + nonTableData[5].Value;
                                    string TempCode = Gettemplate.DocumentCode;
                                    String sPrefixDocNo = $"{TempCode}-{DateTime.Now.Year.ToString()}-";
                                    CustomControlRunning paaa = new CustomControlRunning
                                    {
                                        TemplateId = Gettemplate.TemplateId,
                                        PreFix = sPrefixDocNo,
                                        Digit = 3,
                                    };

                                    Program program = new Program();

                                    string docnoo = program.GetControlRunningAutonum(paaa);
                                    objMemo.DocumentNo = objMemo.DocumentCode;
                                    foreach (var label in labels)
                                    {
                                        if (label == "รหัสเอกสาร")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, docnoo, label);
                                        }
                                        if (label == "วันที่จัดทำแผน")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[1].Value, label);
                                        }
                                        if (label == "หัวหน้าผู้ตรวจ")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[3].Value, label);
                                        }
                                        if (label == "ผู้รับผิดชอบ")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[2].Value, label);
                                        }
                                        if (label == "รหัสพื้นที่ ISO ที่ถูกตรวจ")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[1].Value, label);
                                        }
                                        if (label == "วันที่เริ่มต้น")
                                        {
                                            if (loopdata.Count > 4)
                                            {
                                                Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[4].Value, label);
                                            }

                                        }
                                        if (label == "ผู้จัดทำแผน")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[2].Value, label);
                                        }
                                        if (label == "ฝ่าย / กลุ่มงาน / แผนก ผู้จัดทำแผน")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[4].Value, label);
                                        }
                                        if (label == "ปี")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[5].Value, label);
                                        }
                                        if (label == "วันที่สิ้นสุด")
                                        {
                                            if (loopdata.Count > 4)
                                            {
                                                Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[5].Value, label);
                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 45001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 45001 : 2018")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }

                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 50001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 50001 : 2018")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }

                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 14001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 14001 : 2015")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }
                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 9001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 9001 : 2015")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }

                                            }
                                        }
                                        if (label == "DCC พื้นที่")
                                        {
                                            string value = "";
                                            int i = 0;
                                            foreach (var Eitem in listObjectDCCArea)
                                            {
                                                dynamic item = Eitem;
                                                string name = item[2];
                                                if (i > 0) { value += ","; }
                                                value += $"[{{\"value\": \"{item[2]}\"}},{{\"value\": \"{Arae}\"}},{{\"value\": null}}]";
                                                i++;
                                            }
                                            value = $"[{value}]";
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, value, label);
                                        }
                                        if (label == "สมาชิกคณะทำงานฯ ที่ถูกตรวจ")
                                        {
                                            string value = "";
                                            int i = 0;
                                            foreach (var Eitem in listObjectworkingteam)
                                            {
                                                dynamic item = Eitem;
                                                string name = item[2];
                                                if (i > 0) { value += ","; }
                                                value += $"[{{\"value\": \"{item[2]}\"}},{{\"value\": \"{item[0]}\"}},{{\"value\": null}}]";
                                                i++;
                                            }
                                            value = $"[{value}]";
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, value, label);
                                        }
                                    }

                                    objMemo.MAdvancveForm = Getadvancetem;
                                    dbWolf.TRNMemos.InsertOnSubmit(objMemo);
                                    dbWolf.SubmitChanges();

                                    List<ApprovalDetail> lineapprove = new List<ApprovalDetail>();

                                    var getlineapprove2 = dbWolf.ViewEmployees.Where(x => x.NameEn == loopdata[3].Value).FirstOrDefault();

                                    var getlogic = dbWolf.MSTTemplateLogics.Where(x => x.TemplateId == Gettemplate.TemplateId && x.logictype == "datalineapprove").ToList();

                                    foreach (var loadlocgic in getlogic)
                                    {
                                        var logicType = JObject.Parse(loadlocgic.jsonvalue);
                                        List<string> jvalue = new List<string>();
                                        string jlabel = logicType["label"].ToString();
                                        JToken conditionsToken = logicType["Conditions"];
                                        List<string> labelnew = new List<string>();
                                        if (conditionsToken is JArray jsonArray)
                                        {
                                            foreach (var item in jsonArray)
                                            {
                                                string label = item["label"].ToString();
                                                labelnew.Add(label);
                                            }
                                        }
                                        jvalue.Add(loopdata[1].Value);

                                        var getallview = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == EmpCurrent.EmployeeCode).ToList();
                                        var getlineapprove = GetLineapprove(getallview, Gettemplate.TemplateId, jvalue, jlabel);

                                        lineapprove.AddRange(getlineapprove);

                                    }
                                    foreach (var getempcode in lineapprove)
                                    {
                                        var getempid = dbWolf.ViewEmployees.Where(x => x.EmployeeId == getempcode.emp_id).FirstOrDefault();
                                        TRNLineApprove trnLine3 = new TRNLineApprove();
                                        var SignatureIdmst = dbWolf.MSTMasterDatas.Where(x => x.MasterId == getempcode.signature_id).FirstOrDefault();

                                        trnLine3.MemoId = objMemo.MemoId;
                                        trnLine3.Seq = getempcode.sequence;
                                        trnLine3.EmployeeId = getempcode.emp_id;
                                        trnLine3.EmployeeCode = getempid.EmployeeCode;
                                        trnLine3.NameTh = getempid.NameTh;
                                        trnLine3.NameEn = getempid.NameEn;
                                        trnLine3.PositionTH = getempid.PositionNameTh;
                                        trnLine3.PositionEN = getempid.PositionNameEn;
                                        trnLine3.SignatureId = SignatureIdmst.MasterId;
                                        trnLine3.IsParallel = getempcode.IsParallel;
                                        trnLine3.IsApproveAll = getempcode.IsApproveAll;
                                        trnLine3.ApproveSlot = getempcode.ApproveSlot;
                                        trnLine3.SignatureTh = SignatureIdmst.Value1;
                                        trnLine3.SignatureEn = SignatureIdmst.Value2;
                                        trnLine3.IsActive = getempid.IsActive;
                                        WriteLogFile("trnLine3.IsParallel :" + getempcode.IsParallel.ToString());
                                        WriteLogFile("trnLine3.IsApproveAll :" + getempcode.IsApproveAll.ToString());
                                        WriteLogFile("trnLine3.ApproveSlot :" + getempcode.ApproveSlot.ToString());
                                        dbWolf.TRNLineApproves.InsertOnSubmit(trnLine3);
                                        dbWolf.SubmitChanges();


                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    var GetmemoCom = dbWolf.TRNMemos.Where(x => x.TemplateId == 7
             && x.StatusName == "Completed"
             && x.ModifiedDate >= lastRunTime
             && x.ModifiedDate < DateTime.Now)
                .ToList();

                    if (GetmemoCom.Count > 0)
                    {
                        var Gettemplate = dbWolf.MSTTemplates.Where(x => x.TemplateId == 8).FirstOrDefault();
                        if (Gettemplate != null)
                        {
                            foreach (var template in GetmemoCom)
                            {
                                chtchmemo.MemoId = template.MemoId;
                                string addvaluetoadvanceform = template.MAdvancveForm;

                                var jsonObject = JObject.Parse(addvaluetoadvanceform);

                                var nonTableData = new List<KeyValuePair<string, string>>();
                                var tableData = new List<List<KeyValuePair<string, string>>>();

                                foreach (var item in jsonObject["items"])
                                {
                                    foreach (var layout in item["layout"])
                                    {
                                        var templateType = layout["template"]["type"]?.ToString();
                                        if (templateType == "tb") // Checking if it's a table type
                                        {
                                            string tableLabel = (string)layout["template"]?["label"];
                                            JObject attribute = (JObject)layout["template"]?["attribute"];
                                            JArray columns = (JArray)attribute?["column"];

                                            JArray rows = (JArray)layout["data"]?["row"];
                                            if (rows != null)
                                            {
                                                foreach (JArray row in rows)
                                                {
                                                    // Create a list to hold the label-value pairs for this row.
                                                    var labelValuePairs = new List<KeyValuePair<string, string>>();

                                                    for (int i = 0; i < row.Count; i++)
                                                    {
                                                        JObject cell = (JObject)row[i];
                                                        string value = (string)cell["value"];
                                                        if (columns != null && i < columns.Count)
                                                        {
                                                            JObject column = (JObject)columns[i];
                                                            string columnLabel = (string)column["label"];
                                                            if (!string.IsNullOrEmpty(columnLabel) && !string.IsNullOrEmpty(value))
                                                            {
                                                                labelValuePairs.Add(new KeyValuePair<string, string>(columnLabel, value));
                                                            }
                                                        }
                                                    }

                                                    // Add the list of label-value pairs for this row to the table data.
                                                    tableData.Add(labelValuePairs);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var label = layout["template"]["label"]?.ToString();
                                            var value = layout["data"]["value"]?.ToString();

                                            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(value))
                                            {
                                                nonTableData.Add(new KeyValuePair<string, string>(label, value));

                                            }
                                        }
                                    }
                                }

                                string Getadvancetem = Gettemplate.AdvanceForm;
                                var jsonObject2 = JObject.Parse(Getadvancetem);

                                var labels = new List<string>();

                                // เข้าถึง array 'items'
                                if (jsonObject2["items"] is JArray itemsArray)
                                {
                                    foreach (var item in itemsArray)
                                    {
                                        // เข้าถึง array 'layout' ภายในแต่ละ 'item'
                                        if (item["layout"] is JArray layoutArray)
                                        {
                                            foreach (var layout in layoutArray)
                                            {
                                                // ตรวจสอบและเพิ่ม label ลงในรายการ
                                                var label = layout["template"]?["label"]?.ToString();
                                                if (!string.IsNullOrEmpty(label))
                                                {
                                                    labels.Add(label);
                                                }
                                            }
                                        }
                                    }
                                }
                                foreach (var loopdata in tableData)
                                {
                                    var Arae = loopdata[1].Value;
                                    var SeleteMemo = joinedData.Where(a => a.ObjValue == Arae).FirstOrDefault();
                                    List<object> listObjectDCCArea = new List<object>();
                                    List<object> listObjectworkingteam = new List<object>();
                                    JObject jsonAdvanceForm = JsonUtils.createJsonObject(SeleteMemo.Memo.MAdvancveForm);
                                    JArray itemssArray = (JArray)jsonAdvanceForm["items"];
                                    foreach (JObject jItems in itemssArray)
                                    {
                                        JArray jLayoutArray = (JArray)jItems["layout"];
                                        if (jLayoutArray.Count >= 1)
                                        {
                                            JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                            JObject jData = (JObject)jLayoutArray[0]["data"];
                                            if ((String)jTemplateL["label"] == "Responsibilities : DCC Area")
                                            {

                                                foreach (JArray row in jData["row"])
                                                {
                                                    List<object> rowObject = new List<object>();
                                                    foreach (JObject item in row)
                                                    {
                                                        rowObject.Add(item["value"].ToString());
                                                    }
                                                    listObjectDCCArea.Add(rowObject);
                                                }
                                            }
                                            if ((String)jTemplateL["label"] == "Responsibilities : Working team")
                                            {

                                                foreach (JArray row in jData["row"])
                                                {
                                                    List<object> rowObject = new List<object>();
                                                    foreach (JObject item in row)
                                                    {
                                                        rowObject.Add(item["value"].ToString());
                                                    }
                                                    listObjectworkingteam.Add(rowObject);
                                                }
                                            }
                                        }
                                    }

                                    var getempcodere = dbWolf.MSTEmployees.Where(x => x.EmployeeId == template.RequesterId).FirstOrDefault();

                                    Getadvancetem = Gettemplate.AdvanceForm;
                                    var EmpCurrent = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == getempcodere.EmployeeCode).FirstOrDefault();
                                    var CurrentCom = dbWolf.MSTCompanies.Where(a => a.CompanyCode == EmpCurrent.CompanyCode).FirstOrDefault();
                                    TRNMemo objMemo = new TRNMemo();
                                    objMemo.StatusName = "Draft";
                                    objMemo.CreatedDate = DateTime.Now;
                                    objMemo.CreatedBy = EmpCurrent.NameEn;
                                    objMemo.CreatorId = EmpCurrent.EmployeeId;
                                    objMemo.RequesterId = EmpCurrent.EmployeeId;
                                    objMemo.CNameTh = EmpCurrent.NameTh;
                                    objMemo.CNameEn = EmpCurrent.NameEn;
                                    objMemo.CPositionId = EmpCurrent.PositionId;
                                    objMemo.CPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.CPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.CDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.CDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.CDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.RNameTh = EmpCurrent.NameTh;
                                    objMemo.RNameEn = EmpCurrent.NameEn;
                                    objMemo.RPositionId = EmpCurrent.PositionId;
                                    objMemo.RPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.RPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.RDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.RDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.RDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.ModifiedDate = DateTime.Now;
                                    objMemo.ModifiedBy = objMemo.ModifiedBy;
                                    objMemo.TemplateId = Gettemplate.TemplateId;
                                    objMemo.TemplateName = Gettemplate.TemplateName;
                                    objMemo.GroupTemplateName = Gettemplate.GroupTemplateName;
                                    objMemo.RequestDate = DateTime.Now;
                                    objMemo.PersonWaitingId = EmpCurrent.EmployeeId;
                                    objMemo.PersonWaiting = EmpCurrent.NameEn;

                                    objMemo.CompanyId = CurrentCom.CompanyId;
                                    objMemo.CompanyName = CurrentCom.NameTh;


                                    objMemo.TAdvanceForm = Gettemplate.AdvanceForm;
                                    objMemo.TemplateSubject = Gettemplate.TemplateSubject;
                                    objMemo.TemplateDetail = Guid.NewGuid().ToString().Replace("-", "");
                                    objMemo.ToPerson = Gettemplate.ToId;
                                    objMemo.CcPerson = Gettemplate.CcId;
                                    objMemo.CurrentApprovalLevel = null;
                                    objMemo.ProjectID = 0;
                                    objMemo.Amount = 0;
                                    objMemo.DocumentCode = GenControlRunning(EmpCurrent, Gettemplate.DocumentCode, objMemo, dbWolf);

                                    objMemo.MemoSubject = "รายการตรวจติดตามภายในองค์กร (Internal Audit Checklist) : " + loopdata[1].Value + " " + nonTableData[5].Value;
                                    string TempCode = Gettemplate.DocumentCode;
                                    String sPrefixDocNo = $"{TempCode}-{DateTime.Now.Year.ToString()}-";
                                    CustomControlRunning paaa = new CustomControlRunning
                                    {
                                        TemplateId = Gettemplate.TemplateId,
                                        PreFix = sPrefixDocNo,
                                        Digit = 3,
                                    };

                                    Program program = new Program();

                                    string docnoo = program.GetControlRunningAutonum(paaa);
                                    objMemo.DocumentNo = objMemo.DocumentCode;
                                    foreach (var label in labels)
                                    {
                                        if (label == "รหัสเอกสาร")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, docnoo, label);
                                        }
                                        if (label == "วันที่จัดทำแผน")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[1].Value, label);
                                        }
                                        if (label == "หัวหน้าผู้ตรวจ")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[3].Value, label);
                                        }
                                        if (label == "ผู้รับผิดชอบ")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[2].Value, label);
                                        }
                                        if (label == "รหัสพื้นที่ ISO ที่ถูกตรวจ")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[1].Value, label);
                                        }
                                        if (label == "วันที่เริ่มต้น")
                                        {
                                            if (loopdata.Count > 4)
                                            {
                                                Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[4].Value, label);
                                            }

                                        }
                                        if (label == "ผู้จัดทำแผน")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[2].Value, label);
                                        }
                                        if (label == "ฝ่าย / กลุ่มงาน / แผนก ผู้จัดทำแผน")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[4].Value, label);
                                        }
                                        if (label == "ปี")
                                        {
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, nonTableData[5].Value, label);
                                        }
                                        if (label == "วันที่สิ้นสุด")
                                        {
                                            if (loopdata.Count > 4)
                                            {
                                                Getadvancetem = ReplaceDataProcessBC(Getadvancetem, loopdata[5].Value, label);
                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 45001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 45001 : 2018")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }

                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 50001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 50001 : 2018")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }

                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 14001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 14001 : 2015")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }
                                            }
                                        }
                                        if (label == "ตรวจติดตาม : ISO 9001")
                                        {
                                            var getlistva = loopdata[0].Value.Split(',');
                                            foreach (var item in getlistva)
                                            {
                                                if (item == "ISO 9001 : 2015")
                                                {
                                                    Getadvancetem = ReplaceDataProcessBC(Getadvancetem, item, label);
                                                }

                                            }
                                        }
                                        if (label == "DCC พื้นที่")
                                        {
                                            string value = "";
                                            int i = 0;
                                            foreach (var Eitem in listObjectDCCArea)
                                            {
                                                dynamic item = Eitem;
                                                string name = item[2];
                                                if (i > 0) { value += ","; }
                                                value += $"[{{\"value\": \"{item[2]}\"}},{{\"value\": \"{Arae}\"}},{{\"value\": null}}]";
                                                i++;
                                            }
                                            value = $"[{value}]";
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, value, label);
                                        }
                                        if (label == "สมาชิกคณะทำงานฯ ที่ถูกตรวจ")
                                        {
                                            string value = "";
                                            int i = 0;
                                            foreach (var Eitem in listObjectworkingteam)
                                            {
                                                dynamic item = Eitem;
                                                string name = item[2];
                                                if (i > 0) { value += ","; }
                                                value += $"[{{\"value\": \"{item[2]}\"}},{{\"value\": \"{item[0]}\"}},{{\"value\": null}}]";
                                                i++;
                                            }
                                            value = $"[{value}]";
                                            Getadvancetem = ReplaceDataProcessBC(Getadvancetem, value, label);
                                        }
                                    }

                                    objMemo.MAdvancveForm = Getadvancetem;
                                    dbWolf.TRNMemos.InsertOnSubmit(objMemo);
                                    dbWolf.SubmitChanges();

                                    List<ApprovalDetail> lineapprove = new List<ApprovalDetail>();

                                    var getlineapprove2 = dbWolf.ViewEmployees.Where(x => x.NameEn == loopdata[3].Value).FirstOrDefault();

                                    var getlogic = dbWolf.MSTTemplateLogics.Where(x => x.TemplateId == Gettemplate.TemplateId && x.logictype == "datalineapprove").ToList();

                                    foreach (var loadlocgic in getlogic)
                                    {
                                        var logicType = JObject.Parse(loadlocgic.jsonvalue);
                                        List<string> jvalue = new List<string>();
                                        string jlabel = logicType["label"].ToString();
                                        JToken conditionsToken = logicType["Conditions"];
                                        List<string> labelnew = new List<string>();
                                        if (conditionsToken is JArray jsonArray)
                                        {
                                            foreach (var item in jsonArray)
                                            {
                                                string label = item["label"].ToString();
                                                labelnew.Add(label);
                                            }
                                        }
                                        jvalue.Add(loopdata[1].Value);

                                        var getallview = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == EmpCurrent.EmployeeCode).ToList();
                                        var getlineapprove = GetLineapprove(getallview, Gettemplate.TemplateId, jvalue, jlabel);

                                        lineapprove.AddRange(getlineapprove);

                                    }
                                    foreach (var getempcode in lineapprove)
                                    {
                                        var getempid = dbWolf.ViewEmployees.Where(x => x.EmployeeId == getempcode.emp_id).FirstOrDefault();
                                        TRNLineApprove trnLine3 = new TRNLineApprove();
                                        var SignatureIdmst = dbWolf.MSTMasterDatas.Where(x => x.MasterId == getempcode.signature_id).FirstOrDefault();

                                        trnLine3.MemoId = objMemo.MemoId;
                                        trnLine3.Seq = getempcode.sequence;
                                        trnLine3.EmployeeId = getempcode.emp_id;
                                        trnLine3.EmployeeCode = getempid.EmployeeCode;
                                        trnLine3.NameTh = getempid.NameTh;
                                        trnLine3.NameEn = getempid.NameEn;
                                        trnLine3.PositionTH = getempid.PositionNameTh;
                                        trnLine3.PositionEN = getempid.PositionNameEn;
                                        trnLine3.SignatureId = SignatureIdmst.MasterId;
                                        trnLine3.IsParallel = getempcode.IsParallel;
                                        trnLine3.IsApproveAll = getempcode.IsApproveAll;
                                        trnLine3.ApproveSlot = getempcode.ApproveSlot;
                                        trnLine3.SignatureTh = SignatureIdmst.Value1;
                                        trnLine3.SignatureEn = SignatureIdmst.Value2;
                                        trnLine3.IsActive = getempid.IsActive;
                                        WriteLogFile("trnLine3.IsParallel :" + getempcode.IsParallel.ToString());
                                        WriteLogFile("trnLine3.IsApproveAll :" + getempcode.IsApproveAll.ToString());
                                        WriteLogFile("trnLine3.ApproveSlot :" + getempcode.ApproveSlot.ToString());
                                        dbWolf.TRNLineApproves.InsertOnSubmit(trnLine3);
                                        dbWolf.SubmitChanges();


                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                WriteLogFile("memo catch :" + chtchmemo.MemoId.ToString());
                WriteLogFile("catch :" + ex.ToString());
                WriteLogFile("catch :" + ex.Message);
                WriteLogFile("catch :" + ex.InnerException);
            }
        }
        public static void GetStartFormCAR()
        {
            DataClasses1DataContext dbWolf = new DataClasses1DataContext(dbConnectionStringWolf);
            DateTime lastRunTime = DateTime.Now.AddMinutes(-5);
            TRNMemo getmemoid = new TRNMemo();
            try
            {
                if (_runcatch == "T")
                {
                    var memo = Int32.Parse(_memoid);
                    var gettem = dbWolf.TRNMemos.Where(x => x.MemoId == memo).ToList();
                    foreach (var getnc in gettem)
                    {
                        getmemoid.MemoId = getnc.MemoId;
                        string addvaluetoadvanceform = getnc.MAdvancveForm;

                        var jsonObject = JObject.Parse(addvaluetoadvanceform);

                        var nonTableData = new List<KeyValuePair<string, string>>();
                        var tableData = new List<List<KeyValuePair<string, string>>>();

                        foreach (var item in jsonObject["items"])
                        {
                            foreach (var layout in item["layout"])
                            {
                                var templateType = (string)layout["template"]["type"];
                                if (templateType == "tb") // Checking if it's a table type
                                {
                                    string tableLabel = (string)layout["template"]["label"];
                                    JObject attribute = (JObject)layout["template"]["attribute"];
                                    JArray columns = attribute["column"] as JArray;

                                    var rowData = layout["data"]["row"];
                                    if (rowData is JArray rows) // This will ensure that rows is JArray or null
                                    {
                                        foreach (JArray row in rows)
                                        {
                                            // Create a list to hold the label-value pairs for this row.
                                            var labelValuePairs = new List<KeyValuePair<string, string>>();

                                            for (int i = 0; i < row.Count; i++)
                                            {
                                                JObject cell = (JObject)row[i];
                                                string value = (string)cell["value"];
                                                if (columns != null && i < columns.Count)
                                                {
                                                    JObject column = (JObject)columns[i];
                                                    string columnLabel = (string)column["label"];
                                                    if (!string.IsNullOrEmpty(columnLabel) && !string.IsNullOrEmpty(value))
                                                    {
                                                        labelValuePairs.Add(new KeyValuePair<string, string>(columnLabel, value));
                                                    }
                                                }
                                            }

                                            // Add the list of label-value pairs for this row to the table data.
                                            tableData.Add(labelValuePairs);
                                        }
                                    }
                                }
                                else
                                {
                                    string label = (string)layout["template"]["label"];
                                    string value = (string)layout["data"]["value"];

                                    if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(value))
                                    {
                                        nonTableData.Add(new KeyValuePair<string, string>(label, value));
                                    }
                                }
                            }
                        }

                        var gettemcar = dbWolf.MSTTemplates.Where(x => x.TemplateId == 3).FirstOrDefault();
                        string Getadvancetemcar = gettemcar.AdvanceForm;
                        var jsonObject2 = JObject.Parse(Getadvancetemcar);

                        var labels = new List<string>();

                        // เข้าถึง array 'items'
                        if (jsonObject2["items"] is JArray itemsArray)
                        {
                            foreach (var item in itemsArray)
                            {
                                // เข้าถึง array 'layout' ภายในแต่ละ 'item'
                                if (item["layout"] is JArray layoutArray)
                                {
                                    foreach (var layout in layoutArray)
                                    {
                                        // ตรวจสอบและเพิ่ม label ลงในรายการ
                                        var label = layout["template"]?["label"]?.ToString();
                                        if (!string.IsNullOrEmpty(label))
                                        {
                                            labels.Add(label);
                                        }
                                    }
                                }
                            }
                        }
                        string empcode = string.Empty;
                        foreach (var valenc in tableData)
                        {
                            TRNMemo objMemo = new TRNMemo();
                            foreach (var loppdata in valenc)
                            {
                                if (loppdata.Value == "NC")
                                {
                                    var getauditor = dbWolf.ViewEmployees.Where(x => x.NameEn == valenc[0].Value).FirstOrDefault();
                                    var EmpCurrent = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == getauditor.EmployeeCode).FirstOrDefault();
                                    var CurrentCom = dbWolf.MSTCompanies.Where(a => a.CompanyCode == EmpCurrent.CompanyCode).FirstOrDefault();
                                    empcode = EmpCurrent.EmployeeCode;
                                    objMemo.StatusName = "Draft";
                                    objMemo.CreatedDate = DateTime.Now;
                                    objMemo.CreatedBy = EmpCurrent.NameEn;
                                    objMemo.CreatorId = EmpCurrent.EmployeeId;
                                    objMemo.RequesterId = EmpCurrent.EmployeeId;
                                    objMemo.CNameTh = EmpCurrent.NameTh;
                                    objMemo.CNameEn = EmpCurrent.NameEn;
                                    objMemo.CPositionId = EmpCurrent.PositionId;
                                    objMemo.CPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.CPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.CDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.CDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.CDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.RNameTh = EmpCurrent.NameTh;
                                    objMemo.RNameEn = EmpCurrent.NameEn;
                                    objMemo.RPositionId = EmpCurrent.PositionId;
                                    objMemo.RPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.RPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.RDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.RDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.RDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.ModifiedDate = DateTime.Now;
                                    objMemo.ModifiedBy = objMemo.ModifiedBy;
                                    objMemo.TemplateId = gettemcar.TemplateId;
                                    objMemo.TemplateName = gettemcar.TemplateName;
                                    objMemo.GroupTemplateName = gettemcar.GroupTemplateName;
                                    objMemo.RequestDate = DateTime.Now;
                                    objMemo.PersonWaitingId = EmpCurrent.EmployeeId;
                                    objMemo.PersonWaiting = EmpCurrent.NameEn;
                                    objMemo.CompanyId = CurrentCom.CompanyId;
                                    objMemo.CompanyName = CurrentCom.NameTh;
                                    objMemo.MemoSubject = "การร้องขอให้องค์กรแก้ไขข้อบกพร่อง (Corrective Action Request) :" + nonTableData[6].Value + " " + nonTableData[4].Value;
                                    objMemo.TAdvanceForm = gettemcar.AdvanceForm;
                                    objMemo.TemplateSubject = gettemcar.TemplateSubject;
                                    objMemo.TemplateDetail = Guid.NewGuid().ToString().Replace("-", "");
                                    objMemo.ToPerson = gettemcar.ToId;
                                    objMemo.CcPerson = gettemcar.CcId;
                                    objMemo.CurrentApprovalLevel = null;
                                    objMemo.ProjectID = 0;
                                    objMemo.Amount = 0;
                                    objMemo.DocumentCode = GenControlRunning(EmpCurrent, gettemcar.DocumentCode, objMemo, dbWolf);
                                    objMemo.DocumentNo = objMemo.DocumentCode;
                                    string date = DateTime.Now.ToString("dd-MM-yyyy");
                                    string sould = "Internal audit (INT)";
                                    foreach (var getlabel in labels)
                                    {
                                        if (getlabel == "เลขที่คำร้อง")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, objMemo.DocumentCode, getlabel);
                                        }
                                        if (getlabel == "เลขที่อ้างอิง")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[0].Value, getlabel);
                                        }
                                        if (getlabel == "วันที่พบ")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, date, getlabel);
                                        }
                                        if (getlabel == "ผู้เปิด")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, EmpCurrent.NameEn, getlabel);
                                        }
                                        if (getlabel == "ฝ่าย / กลุ่มงาน / แผนก ผู้เปิด")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, EmpCurrent.DepartmentNameEn, getlabel);
                                        }
                                        if (getlabel == "หัวหน้าผู้ตรวจ")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[5].Value, getlabel);
                                        }
                                        if (getlabel == "ผู้รับผิดชอบ")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[7].Value, getlabel);
                                        }
                                        if (getlabel == "รหัสพื้นที ISO ที่เกี่ยวข้อง")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[6].Value, getlabel);
                                        }
                                        if (getlabel == "มาตรฐาน")
                                        {
                                            foreach (var getv in valenc)
                                            {
                                                if (getv.Key == "มาตรฐาน")
                                                {
                                                    Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, getv.Value, getlabel);
                                                }
                                            }


                                        }
                                        if (getlabel == "ข้อกำหนด")
                                        {
                                            foreach (var getv in valenc)
                                            {
                                                if (getv.Key == "ข้อกำหนด")
                                                {
                                                    Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, getv.Value, getlabel);
                                                }
                                            }
                                        }
                                        if (getlabel == "รายละเอียดข้อบกพร่อง")
                                        {
                                            foreach (var getv in valenc)
                                            {
                                                if (getv.Key == "สรุปผลการตรวจ")
                                                {
                                                    Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, getv.Value, getlabel);
                                                }
                                            }
                                        }
                                        if (getlabel == "แหล่งที่มา")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, sould, getlabel);
                                        }

                                    }

                                    objMemo.MAdvancveForm = Getadvancetemcar;
                                    dbWolf.TRNMemos.InsertOnSubmit(objMemo);
                                    dbWolf.SubmitChanges();

                                    List<ApprovalDetail> lineapprove = new List<ApprovalDetail>();
                                    var getlogic = dbWolf.MSTTemplateLogics.Where(x => x.TemplateId == gettemcar.TemplateId && x.logictype == "datalineapprove").ToList();

                                    foreach (var loadlocgic in getlogic)
                                    {
                                        var logicType = JObject.Parse(loadlocgic.jsonvalue);
                                        List<string> jvalue = new List<string>();
                                        string jlabel = logicType["label"].ToString();
                                        JToken conditionsToken = logicType["Conditions"];
                                        List<string> labelnew = new List<string>();
                                        if (conditionsToken is JArray jsonArray)
                                        {
                                            foreach (var item in jsonArray)
                                            {
                                                string label = item["label"].ToString();
                                                labelnew.Add(label);
                                            }
                                        }
                                        string getval1 = nonTableData[6].Key + "|" + nonTableData[6].Value;
                                        string getval2 = valenc[2].Key + "|" + valenc[2].Value;
                                        jvalue.Add(getval1);
                                        jvalue.Add(getval2);

                                        var getallview = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == empcode).ToList();
                                        var getlineapprove = GetLineapprove(getallview, gettemcar.TemplateId, jvalue, jlabel);

                                        lineapprove.AddRange(getlineapprove);
                                    }

                                    foreach (var getempcode in lineapprove)
                                    {

                                        TRNLineApprove trnLine3 = new TRNLineApprove();
                                        var getempid = dbWolf.ViewEmployees.Where(x => x.EmployeeId == getempcode.emp_id).FirstOrDefault();
                                        var SignatureIdmst = dbWolf.MSTMasterDatas.Where(x => x.MasterId == getempcode.signature_id).FirstOrDefault();
                                        trnLine3.MemoId = objMemo.MemoId;
                                        trnLine3.Seq = getempcode.sequence;
                                        trnLine3.EmployeeId = getempcode.emp_id;
                                        trnLine3.EmployeeCode = getempid.EmployeeCode;
                                        trnLine3.NameTh = getempid.NameTh;
                                        trnLine3.NameEn = getempid.NameEn;
                                        trnLine3.PositionTH = getempid.PositionNameTh;
                                        trnLine3.PositionEN = getempid.PositionNameEn;
                                        trnLine3.SignatureId = SignatureIdmst.MasterId;
                                        trnLine3.IsParallel = getempcode.IsParallel;
                                        trnLine3.IsApproveAll = getempcode.IsApproveAll;
                                        trnLine3.ApproveSlot = getempcode.ApproveSlot;
                                        trnLine3.SignatureTh = SignatureIdmst.Value1;
                                        trnLine3.SignatureEn = SignatureIdmst.Value2;
                                        trnLine3.IsActive = getempid.IsActive;

                                        dbWolf.TRNLineApproves.InsertOnSubmit(trnLine3);
                                        dbWolf.SubmitChanges();


                                    }
                                }
                            }


                        }


                    }
                }
                else
                {
                    var gettem = dbWolf.TRNMemos.Where(x => x.TemplateId == 8 && x.StatusName == "Completed" && x.ModifiedDate >= lastRunTime
                 && x.ModifiedDate < DateTime.Now).ToList();
                    foreach (var getnc in gettem)
                    {
                        getmemoid.MemoId = getnc.MemoId;
                        string addvaluetoadvanceform = getnc.MAdvancveForm;

                        var jsonObject = JObject.Parse(addvaluetoadvanceform);

                        var nonTableData = new List<KeyValuePair<string, string>>();
                        var tableData = new List<List<KeyValuePair<string, string>>>();

                        foreach (var item in jsonObject["items"])
                        {
                            foreach (var layout in item["layout"])
                            {
                                var templateType = (string)layout["template"]["type"];
                                if (templateType == "tb") // Checking if it's a table type
                                {
                                    string tableLabel = (string)layout["template"]["label"];
                                    JObject attribute = (JObject)layout["template"]["attribute"];
                                    JArray columns = attribute["column"] as JArray;

                                    var rowData = layout["data"]["row"];
                                    if (rowData is JArray rows) // This will ensure that rows is JArray or null
                                    {
                                        foreach (JArray row in rows)
                                        {
                                            // Create a list to hold the label-value pairs for this row.
                                            var labelValuePairs = new List<KeyValuePair<string, string>>();

                                            for (int i = 0; i < row.Count; i++)
                                            {
                                                JObject cell = (JObject)row[i];
                                                string value = (string)cell["value"];
                                                if (columns != null && i < columns.Count)
                                                {
                                                    JObject column = (JObject)columns[i];
                                                    string columnLabel = (string)column["label"];
                                                    if (!string.IsNullOrEmpty(columnLabel) && !string.IsNullOrEmpty(value))
                                                    {
                                                        labelValuePairs.Add(new KeyValuePair<string, string>(columnLabel, value));
                                                    }
                                                }
                                            }

                                            // Add the list of label-value pairs for this row to the table data.
                                            tableData.Add(labelValuePairs);
                                        }
                                    }
                                }
                                else
                                {
                                    string label = (string)layout["template"]["label"];
                                    string value = (string)layout["data"]["value"];

                                    if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(value))
                                    {
                                        nonTableData.Add(new KeyValuePair<string, string>(label, value));
                                    }
                                }
                            }
                        }

                        var gettemcar = dbWolf.MSTTemplates.Where(x => x.TemplateId == 3).FirstOrDefault();
                        string Getadvancetemcar = gettemcar.AdvanceForm;
                        var jsonObject2 = JObject.Parse(Getadvancetemcar);

                        var labels = new List<string>();

                        // เข้าถึง array 'items'
                        if (jsonObject2["items"] is JArray itemsArray)
                        {
                            foreach (var item in itemsArray)
                            {
                                // เข้าถึง array 'layout' ภายในแต่ละ 'item'
                                if (item["layout"] is JArray layoutArray)
                                {
                                    foreach (var layout in layoutArray)
                                    {
                                        // ตรวจสอบและเพิ่ม label ลงในรายการ
                                        var label = layout["template"]?["label"]?.ToString();
                                        if (!string.IsNullOrEmpty(label))
                                        {
                                            labels.Add(label);
                                        }
                                    }
                                }
                            }
                        }
                        string empcode = string.Empty;
                        foreach (var valenc in tableData)
                        {
                            TRNMemo objMemo = new TRNMemo();
                            foreach (var loppdata in valenc)
                            {
                                if (loppdata.Value == "NC")
                                {
                                    var getauditor = dbWolf.ViewEmployees.Where(x => x.NameEn == valenc[0].Value).FirstOrDefault();
                                  var EmpCurrent = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == getauditor.EmployeeCode).FirstOrDefault();
                                    var CurrentCom = dbWolf.MSTCompanies.Where(a => a.CompanyCode == EmpCurrent.CompanyCode).FirstOrDefault();
                                    empcode = EmpCurrent.EmployeeCode;
                                    objMemo.StatusName = "Draft";
                                    objMemo.CreatedDate = DateTime.Now;
                                    objMemo.CreatedBy = EmpCurrent.NameEn;
                                    objMemo.CreatorId = EmpCurrent.EmployeeId;
                                    objMemo.RequesterId = EmpCurrent.EmployeeId;
                                    objMemo.CNameTh = EmpCurrent.NameTh;
                                    objMemo.CNameEn = EmpCurrent.NameEn;
                                    objMemo.CPositionId = EmpCurrent.PositionId;
                                    objMemo.CPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.CPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.CDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.CDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.CDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.RNameTh = EmpCurrent.NameTh;
                                    objMemo.RNameEn = EmpCurrent.NameEn;
                                    objMemo.RPositionId = EmpCurrent.PositionId;
                                    objMemo.RPositionTh = EmpCurrent.PositionNameTh;
                                    objMemo.RPositionEn = EmpCurrent.PositionNameEn;
                                    objMemo.RDepartmentId = EmpCurrent.DepartmentId;
                                    objMemo.RDepartmentTh = EmpCurrent.DepartmentNameTh;
                                    objMemo.RDepartmentEn = EmpCurrent.DepartmentNameEn;
                                    objMemo.ModifiedDate = DateTime.Now;
                                    objMemo.ModifiedBy = objMemo.ModifiedBy;
                                    objMemo.TemplateId = gettemcar.TemplateId;
                                    objMemo.TemplateName = gettemcar.TemplateName;
                                    objMemo.GroupTemplateName = gettemcar.GroupTemplateName;
                                    objMemo.RequestDate = DateTime.Now;
                                    objMemo.PersonWaitingId = EmpCurrent.EmployeeId;
                                    objMemo.PersonWaiting = EmpCurrent.NameEn;
                                    objMemo.CompanyId = CurrentCom.CompanyId;
                                    objMemo.CompanyName = CurrentCom.NameTh;
                                    objMemo.MemoSubject = "การร้องขอให้องค์กรแก้ไขข้อบกพร่อง (Corrective Action Request) :" + nonTableData[6].Value + " " + nonTableData[4].Value;
                                    objMemo.TAdvanceForm = gettemcar.AdvanceForm;
                                    objMemo.TemplateSubject = gettemcar.TemplateSubject;
                                    objMemo.TemplateDetail = Guid.NewGuid().ToString().Replace("-", "");
                                    objMemo.ToPerson = gettemcar.ToId;
                                    objMemo.CcPerson = gettemcar.CcId;
                                    objMemo.CurrentApprovalLevel = null;
                                    objMemo.ProjectID = 0;
                                    objMemo.Amount = 0;
                                    objMemo.DocumentCode = GenControlRunning(EmpCurrent, gettemcar.DocumentCode, objMemo, dbWolf);
                                    objMemo.DocumentNo = objMemo.DocumentCode;
                                    string date = DateTime.Now.ToString("dd-MM-yyyy");
                                    string sould = "Internal audit (INT)";
                                    foreach (var getlabel in labels)
                                    {
                                        if (getlabel == "เลขที่คำร้อง")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, objMemo.DocumentCode, getlabel);
                                        }
                                        if (getlabel == "เลขที่อ้างอิง")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[0].Value, getlabel);
                                        }
                                        if (getlabel == "วันที่พบ")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, date, getlabel);
                                        }
                                        if (getlabel == "ผู้เปิด")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, EmpCurrent.NameEn, getlabel);
                                        }
                                        if (getlabel == "ฝ่าย / กลุ่มงาน / แผนก ผู้เปิด")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, EmpCurrent.DepartmentNameEn, getlabel);
                                        }
                                        if (getlabel == "หัวหน้าผู้ตรวจ")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[5].Value, getlabel);
                                        }
                                        if (getlabel == "ผู้รับผิดชอบ")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[7].Value, getlabel);
                                        }
                                        if (getlabel == "รหัสพื้นที ISO ที่เกี่ยวข้อง")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, nonTableData[6].Value, getlabel);
                                        }
                                        if (getlabel == "มาตรฐาน")
                                        {
                                            foreach (var getv in valenc)
                                            {
                                                if (getv.Key == "มาตรฐาน")
                                                {
                                                    Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, getv.Value, getlabel);
                                                }
                                            }


                                        }
                                        if (getlabel == "ข้อกำหนด")
                                        {
                                            foreach (var getv in valenc)
                                            {
                                                if (getv.Key == "ข้อกำหนด")
                                                {
                                                    Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, getv.Value, getlabel);
                                                }
                                            }
                                        }
                                        if (getlabel == "รายละเอียดข้อบกพร่อง")
                                        {
                                            foreach (var getv in valenc)
                                            {
                                                if (getv.Key == "สรุปผลการตรวจ")
                                                {
                                                    Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, getv.Value, getlabel);
                                                }
                                            }
                                        }
                                        if (getlabel == "แหล่งที่มา")
                                        {
                                            Getadvancetemcar = ReplaceDataProcessBC(Getadvancetemcar, sould, getlabel);
                                        }

                                    }

                                    objMemo.MAdvancveForm = Getadvancetemcar;
                                    dbWolf.TRNMemos.InsertOnSubmit(objMemo);
                                    dbWolf.SubmitChanges();

                                    List<ApprovalDetail> lineapprove = new List<ApprovalDetail>();
                                    var getlogic = dbWolf.MSTTemplateLogics.Where(x => x.TemplateId == gettemcar.TemplateId && x.logictype == "datalineapprove").ToList();
                                    var getallviewseq4 = dbWolf.ViewEmployees.Where(x => x.NameEn == nonTableData[5].Value).FirstOrDefault();
                                    foreach (var loadlocgic in getlogic)
                                    {
                                        var logicType = JObject.Parse(loadlocgic.jsonvalue);
                                        List<string> jvalue = new List<string>();
                                        string jlabel = logicType["label"].ToString();
                                        JToken conditionsToken = logicType["Conditions"];
                                        List<string> labelnew = new List<string>();
                                        if (conditionsToken is JArray jsonArray)
                                        {
                                            foreach (var item in jsonArray)
                                            {
                                                string label = item["label"].ToString();
                                                labelnew.Add(label);
                                            }
                                        }
                                        string getval1 = nonTableData[6].Key + "|" + nonTableData[6].Value;
                                        string getval2 = valenc[2].Key + "|" + valenc[2].Value;
                                        jvalue.Add(getval1);
                                        jvalue.Add(getval2);

                                        var getallview = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == empcode).ToList();
                                        var getlineapprove = GetLineapprove(getallview, gettemcar.TemplateId, jvalue, jlabel);

                                        lineapprove.AddRange(getlineapprove);
                                    }
                                    int i = 0;
                                    foreach (var getempcode in lineapprove)
                                    {
                                        var SignatureIdmst = dbWolf.MSTMasterDatas.Where(x => x.MasterId == getempcode.signature_id).FirstOrDefault();
                                        var SignatureIdmst2 = dbWolf.MSTMasterDatas.Where(x => x.Value1 == "Review Result").FirstOrDefault();
                                        TRNLineApprove trnLine3 = new TRNLineApprove();
                                        TRNLineApprove trnLine4 = new TRNLineApprove();
                                        if (i == 4)
                                        {

                                            trnLine3.MemoId = objMemo.MemoId;
                                            trnLine3.Seq = i;
                                            trnLine3.EmployeeId = getallviewseq4.EmployeeId;
                                            trnLine3.EmployeeCode = getallviewseq4.EmployeeCode;
                                            trnLine3.NameTh = getallviewseq4.NameTh;
                                            trnLine3.NameEn = getallviewseq4.NameEn;
                                            trnLine3.PositionTH = getallviewseq4.PositionNameTh;
                                            trnLine3.PositionEN = getallviewseq4.PositionNameEn;
                                            trnLine3.SignatureId = SignatureIdmst2.MasterId;
                                            trnLine3.IsParallel = null;
                                            trnLine3.IsApproveAll = null;
                                            trnLine3.ApproveSlot = null;
                                            trnLine3.SignatureTh = SignatureIdmst2.Value1;
                                            trnLine3.SignatureEn = SignatureIdmst2.Value2;
                                            trnLine3.IsActive = getallviewseq4.IsActive;

                                            dbWolf.TRNLineApproves.InsertOnSubmit(trnLine3);
                                            dbWolf.SubmitChanges();

                                            i++;
                                            var getempid = dbWolf.ViewEmployees.Where(x => x.EmployeeId == getempcode.emp_id).FirstOrDefault();
                                            trnLine4.MemoId = objMemo.MemoId;
                                            trnLine4.Seq = i;
                                            trnLine4.EmployeeId = getempcode.emp_id;
                                            trnLine4.EmployeeCode = getempid.EmployeeCode;
                                            trnLine4.NameTh = getempid.NameTh;
                                            trnLine4.NameEn = getempid.NameEn;
                                            trnLine4.PositionTH = getempid.PositionNameTh;
                                            trnLine4.PositionEN = getempid.PositionNameEn;
                                            trnLine4.SignatureId = SignatureIdmst.MasterId;
                                            trnLine4.IsParallel = getempcode.IsParallel;
                                            trnLine4.IsApproveAll = getempcode.IsApproveAll;
                                            trnLine4.ApproveSlot = getempcode.ApproveSlot;
                                            trnLine4.SignatureTh = SignatureIdmst.Value1;
                                            trnLine4.SignatureEn = SignatureIdmst.Value2;
                                            trnLine4.IsActive = getempid.IsActive;

                                            dbWolf.TRNLineApproves.InsertOnSubmit(trnLine4);
                                            dbWolf.SubmitChanges();
                                        }
                                        else
                                        {
                                            var getempid = dbWolf.ViewEmployees.Where(x => x.EmployeeId == getempcode.emp_id).FirstOrDefault();
                                            trnLine3.MemoId = objMemo.MemoId;
                                            trnLine3.Seq = i;
                                            trnLine3.EmployeeId = getempcode.emp_id;
                                            trnLine3.EmployeeCode = getempid.EmployeeCode;
                                            trnLine3.NameTh = getempid.NameTh;
                                            trnLine3.NameEn = getempid.NameEn;
                                            trnLine3.PositionTH = getempid.PositionNameTh;
                                            trnLine3.PositionEN = getempid.PositionNameEn;
                                            trnLine3.SignatureId = SignatureIdmst.MasterId;
                                            trnLine3.IsParallel = getempcode.IsParallel;
                                            trnLine3.IsApproveAll = getempcode.IsApproveAll;
                                            trnLine3.ApproveSlot = getempcode.ApproveSlot;
                                            trnLine3.SignatureTh = SignatureIdmst.Value1;
                                            trnLine3.SignatureEn = SignatureIdmst.Value2;
                                            trnLine3.IsActive = getempid.IsActive;

                                            dbWolf.TRNLineApproves.InsertOnSubmit(trnLine3);
                                            dbWolf.SubmitChanges();
                                        }



                                        i++;

                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                WriteLogFile("memo catch :" + getmemoid.MemoId.ToString());
                WriteLogFile("catch :" + ex.ToString());
                WriteLogFile("catch :" + ex.Message);
                WriteLogFile("catch :" + ex.InnerException);
            }


        }
        public static List<ApprovalDetail> GetLineapprove(List<ViewEmployee> lstEmp, int TemplateID, List<string> Type, string heardlabel)
        {
            //Type = Type.Replace("\\", "");

            Console.WriteLine("Type : " + Type);

            DataClasses1DataContext dbWolf = new DataClasses1DataContext(dbConnectionStringWolf);
            List<MSTTemplateLogic> LogicID = new List<MSTTemplateLogic>();
            string logic = string.Empty;
            if (dbWolf.Connection.State == ConnectionState.Open)
            {
                dbWolf.Connection.Close();
                dbWolf.Connection.Open();
            }
            else
            {

                dbWolf.Connection.Open();
                LogicID = dbWolf.MSTTemplateLogics.Where(a => a.TemplateId == TemplateID & a.logictype == "datalineapprove").ToList();
                foreach (var loadlocgic in LogicID)
                {
                    var logicType = JObject.Parse(loadlocgic.jsonvalue);

                    string jlabel = logicType["label"].ToString();
                    if (heardlabel.Contains(jlabel))
                    {
                        logic = loadlocgic.logicid.ToString();
                        break;
                    }
                }

            }
            string jsonFormatCondition = "{'logicid':'" + logic + "','conditions':[";

            foreach (var item in Type)
            {

                var parts = item.Split('|');
                if (parts.Length >= 2)
                {
                    jsonFormatCondition += "{'label':'" + parts[0].Replace("'", "\\'") + "','value':'" + parts[1].Replace("'", "\\'") + "'},";
                }
                else
                {
                    jsonFormatCondition += "{'label':'" + heardlabel.Replace("'", "\\'") + "','value':'" + item.Replace("'", "\\'") + "'},";
                }



            }

            // Remove the trailing comma if not needed
            if (jsonFormatCondition.EndsWith(","))
            {
                jsonFormatCondition = jsonFormatCondition.Remove(jsonFormatCondition.Length - 1);
            }

            jsonFormatCondition += "]}";

            // Replace single quotes with double quotes to make it a valid JSON string
            jsonFormatCondition = jsonFormatCondition.Replace("'", "\"");
            string FormatCondition = jsonFormatCondition;


            Console.WriteLine("FormatCondition : " + FormatCondition);

            ////return lstapprovalDetails;
            TemplateDetailFormPage templateDetailFormPage = new TemplateDetailFormPage();

            try
            {

                templateDetailFormPage.connectionString = dbConnectionStringWolf;
                templateDetailFormPage.lstTRNLineApprove = new List<ApprovalDetail>();
                templateDetailFormPage.templateForm = JsonConvert.DeserializeObject<CustomTemplate>(postAPI($"api/Template/TemplateByID", new CustomTemplate() { connectionString = dbConnectionStringWolf, TemplateId = TemplateID }));
                templateDetailFormPage.VEmployee = ConvertEmpToCustom(lstEmp.First());
                templateDetailFormPage.JsonCondition = FormatCondition;
                Console.WriteLine("Get TemplateLogic ID : " + TemplateID);

            }
            catch (Exception ex)
            {
                Console.WriteLine("GetLineapprove error :" + ex.Message.ToString());
                throw;
            }

            WriteLogFile(templateDetailFormPage.ToJson());
            return JsonConvert.DeserializeObject<List<ApprovalDetail>>(postAPI($"api/LineApprove/LineApproveWithTemplate", templateDetailFormPage));

        }

        public static CustomViewEmployee ConvertEmpToCustom(ViewEmployee viewEmployee)
        {
            CustomViewEmployee oResult = new CustomViewEmployee();
            oResult.EmployeeId = viewEmployee.EmployeeId;
            oResult.EmployeeCode = viewEmployee.EmployeeCode;
            oResult.Username = viewEmployee.Username;
            oResult.NameTh = viewEmployee.NameTh;
            oResult.NameEn = viewEmployee.NameEn;
            oResult.Email = viewEmployee.Email;
            oResult.IsActive = viewEmployee.IsActive;
            oResult.PositionId = viewEmployee.PositionId;
            oResult.PositionNameTh = viewEmployee.PositionNameTh;
            oResult.PositionNameEn = viewEmployee.PositionNameEn;
            oResult.DepartmentId = viewEmployee.DepartmentId;
            oResult.DepartmentNameTh = viewEmployee.DepartmentNameTh;
            oResult.DepartmentNameEn = viewEmployee.DepartmentNameEn;
            oResult.SignPicPath = viewEmployee.SignPicPath;
            oResult.Lang = viewEmployee.Lang;
            //AccountId = viewEmployee.AccountId;
            oResult.AccountCode = viewEmployee.AccountCode;
            oResult.AccountName = viewEmployee.AccountName;
            oResult.DefaultLang = viewEmployee.DefaultLang;
            oResult.RegisteredDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.RegisteredDate);
            oResult.ExpiredDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.ExpiredDate);
            oResult.CreatedDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.CreatedDate);
            oResult.CreatedBy = viewEmployee.CreatedBy;
            oResult.ModifiedDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.ModifiedDate);
            oResult.ModifiedBy = viewEmployee.ModifiedBy;
            oResult.ReportToEmpCode = viewEmployee.ReportToEmpCode;
            oResult.DivisionId = viewEmployee.DivisionId;
            oResult.DivisionNameTh = viewEmployee.DivisionNameTh;
            oResult.DivisionNameEn = viewEmployee.DivisionNameEn;
            oResult.ADTitle = viewEmployee.ADTitle;

            return oResult;
        }
        public static List<CustomViewEmployee> GetEmployeeByEmpCode(CustomViewEmployee empCode)
        {
            List<CustomViewEmployee> obj = JsonConvert.DeserializeObject<List<CustomViewEmployee>>(postAPI($"api/Employee/EmployeeByEmpCode", empCode));
            return obj == null ? new List<CustomViewEmployee>() : obj;
        }

        public static string postAPI(string subUri, Object obj)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.BaseAddress = new Uri(_BaseAPI);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    string json = new JavaScriptSerializer().Serialize(obj);
                    StringContent content = new StringContent(json, Encoding.UTF8, "application/json");

                    Task<HttpResponseMessage> response = client.PostAsync(subUri, content);

                    if (response.Result.IsSuccessStatusCode)
                    {
                        return response.Result.Content.ReadAsStringAsync().Result;
                    }
                    else
                    {
                        return "Not Found";
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public static string GenControlRunning(ViewEmployee Emp, string DocumentCode, TRNMemo objTRNMemo, DataClasses1DataContext db)
        {
            string TempCode = DocumentCode;
            String sPrefixDocNo = $"{TempCode}-{DateTime.Now.Year.ToString()}-";
            int iRunning = 1;
            List<TRNMemo> temp = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sPrefixDocNo.ToUpper())).ToList();
            if (temp.Count > 0)
            {
                String sLastDocumentNo = temp.OrderBy(a => a.DocumentNo).Last().DocumentNo;
                if (!String.IsNullOrEmpty(sLastDocumentNo))
                {
                    List<String> list_LastDocumentNo = sLastDocumentNo.Split('-').ToList();

                    if (list_LastDocumentNo.Count >= 3)
                    {
                        iRunning = checkDataIntIsNull(list_LastDocumentNo[list_LastDocumentNo.Count - 1]) + 1;
                    }
                }
            }
            String sDocumentNo = $"{sPrefixDocNo}{iRunning.ToString().PadLeft(6, '0')}";



            return sDocumentNo;
        }
        public static int checkDataIntIsNull(object Input)
        {
            int Results = 0;
            if (Input != null)
                int.TryParse(Input.ToString().Replace(",", ""), out Results);

            return Results;
        }
        public static string DocNoGenerate(string FixDoc, string DocCode, string CCode, string DCode, string DSCode, DataClasses1DataContext db)
        {
            string sDocumentNo = "";
            int iRunning;
            if (!string.IsNullOrWhiteSpace(FixDoc))
            {
                string y4 = DateTime.Now.ToString("yyyy");
                string y2 = DateTime.Now.ToString("yy");
                string CompanyCode = CCode;
                string DepartmentCode = DCode;
                string DivisionCode = DSCode;
                string FixCode = FixDoc;
                FixCode = FixCode.Replace("[CompanyCode]", CompanyCode);
                FixCode = FixCode.Replace("[DepartmentCode]", DepartmentCode);
                FixCode = FixCode.Replace("[DocumentCode]", DocCode);
                FixCode = FixCode.Replace("[DivisionCode]", DivisionCode);

                FixCode = FixCode.Replace("[YYYY]", y4);
                FixCode = FixCode.Replace("[YY]", y2);
                sDocumentNo = FixCode;
                List<TRNMemo> tempfixDoc = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sDocumentNo.ToUpper())).ToList();


                List<TRNMemo> tempfixDocByYear = db.TRNMemos.ToList();

                tempfixDocByYear = tempfixDocByYear.FindAll(a => a.DocumentNo != ("Auto Generate") & Convert.ToDateTime(a.RequestDate).Year.ToString().Equals(y4)).ToList();

                if (tempfixDocByYear.Count > 0)
                {
                    tempfixDocByYear = tempfixDocByYear.OrderByDescending(a => a.MemoId).ToList();

                    String sLastDocumentNofix = tempfixDocByYear.First().DocumentNo;
                    if (!String.IsNullOrEmpty(sLastDocumentNofix))
                    {
                        List<String> list_LastDocumentNofix = sLastDocumentNofix.Split('-').ToList();

                        if (list_LastDocumentNofix.Count >= 3)
                        {
                            iRunning = checkDataIntIsNull(list_LastDocumentNofix[list_LastDocumentNofix.Count - 1]) + 1;
                            sDocumentNo = $"{sDocumentNo}-{iRunning.ToString().PadLeft(3, '0')}";




                        }
                    }
                }
                else
                {
                    sDocumentNo = $"{sDocumentNo}-{1.ToString().PadLeft(3, '0')}";

                }
            }
            return sDocumentNo;
        }
        public static string ReplaceDataProcessBC(string DestAdvanceForm, string Value, string label)
        {
            JObject jsonAdvanceForm = JObject.Parse(DestAdvanceForm); // Assuming createJsonObject(string) does the same
            JArray itemsArray = (JArray)jsonAdvanceForm["items"];
            if (itemsArray == null) return DestAdvanceForm; // Early return if items array is not present

            foreach (JObject jItems in itemsArray)
            {
                JArray jLayoutArray = (JArray)jItems["layout"];
                if (jLayoutArray == null) continue; // Skip if layout array is not present

                foreach (JObject jLayout in jLayoutArray)
                {
                    JObject jTemplate = jLayout["template"] as JObject;

                    if (jTemplate != null && (string)jTemplate["label"] == label)
                    {
                        JObject jData = jLayout["data"] as JObject;
                        if (jData["row"] != null && jData["row"].Type == JTokenType.Null)
                        {
                            jData.Remove("row");
                            jData.Add("row", JArray.Parse(Value));
                            break;
                        }
                        else if (jData != null)
                        {
                            jData["value"] = Value;
                            break;
                        }
                    }
                }
            }
            return JsonConvert.SerializeObject(jsonAdvanceForm);
        }

        public static JObject createJsonObject(string jsonStr)
        {
            JObject json = null;
            try
            {
                json = (JObject)JProperty.Parse(jsonStr);
            }
            catch (Exception)
            {
                json = new JObject();
            }
            return json;
        }
        public static void WriteLogFile(String iText)
        {

            String LogFilePath = String.Format("{0}{1}_OrderLog.txt", _LogFile, DateTime.Now.ToString("yyyyMMdd"));

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(LogFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] ListText = iText.Split('|').ToArray();

                    foreach (String s in ListText)
                    {
                        sbLog.AppendLine(s);
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch { }
        }
        private static string _BaseAPI
        {
            get
            {
                var BaseAPI = System.Configuration.ConfigurationSettings.AppSettings["BaseAPI"];
                if (!string.IsNullOrEmpty(BaseAPI))
                {
                    return (BaseAPI);
                }
                return string.Empty;
            }
        }

        private static string _memoid
        {
            get
            {
                var memoid = System.Configuration.ConfigurationSettings.AppSettings["memoid"];
                if (!string.IsNullOrEmpty(memoid))
                {
                    return (memoid);
                }
                return string.Empty;
            }
        }
        private static string _runcatch
        {
            get
            {
                var runcatch = System.Configuration.ConfigurationSettings.AppSettings["runcatch"];
                if (!string.IsNullOrEmpty(runcatch))
                {
                    return (runcatch);
                }
                return string.Empty;
            }
        }

        public string GetControlRunningAutonum(CustomControlRunning icontrolRunning)
        {
            string result = string.Empty;
            try
            {
                DataClasses1DataContext dbWolf = new DataClasses1DataContext(dbConnectionStringWolf);

                if (string.IsNullOrEmpty(icontrolRunning.RunningNumber))
                {
                    string sPrefix = string.Empty;
                    int nRunning = 0;
                    int nDigit = 0;

                    string Result_Running = string.Empty;
                    sPrefix = icontrolRunning.PreFix;
                    nDigit = icontrolRunning.Digit;
                    int get_Digit = Convert.ToInt32(icontrolRunning.Digit);

                    List<TRNMemo> lstTrnMemo = new List<TRNMemo>();



                    List<TRNControlRunning> lst_ControlRunning = dbWolf.TRNControlRunnings.Where(x => x.TemplateId == icontrolRunning.TemplateId && x.Prefix == icontrolRunning.PreFix && x.Digit == icontrolRunning.Digit).ToList();
                    if (lst_ControlRunning.Count > 0)
                    {
                        nRunning = Convert.ToInt32(lst_ControlRunning.LastOrDefault().Running + 1);

                        var mstMasterData = dbWolf.MSTMasterDatas.FirstOrDefault(x => x.MasterType == "CHK_AT_NB" & x.IsActive == true);

                        nRunning = CheckAutoNumber(dbWolf, lst_ControlRunning.LastOrDefault().RunningNumber,
                            icontrolRunning.TemplateId,
                            nRunning,
                            sPrefix);
                    }

                    else
                    {
                        nRunning = 1;
                    }


                    Result_Running = sPrefix + nRunning.ToString().PadLeft(get_Digit, '0');


                    result = Result_Running;
                    List<TRNMemo> temp = dbWolf.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sPrefix.ToUpper())).ToList();
                    String sLastDocumentNo = temp.OrderBy(a => a.DocumentNo).Last().DocumentNo;
                    int iRunning = 1;
                    if (!String.IsNullOrEmpty(sLastDocumentNo))
                    {
                        List<String> list_LastDocumentNo = sLastDocumentNo.Split('-').ToList();

                        if (list_LastDocumentNo.Count >= 3)
                        {
                            iRunning = checkDataIntIsNull(list_LastDocumentNo[list_LastDocumentNo.Count - 1]) + 1;
                        }
                    }
                    TRNControlRunning controrun = new TRNControlRunning();
                    controrun.TemplateId = icontrolRunning.TemplateId;
                    controrun.Prefix = sPrefix;
                    controrun.Digit = nDigit;
                    controrun.Running = iRunning;
                    controrun.CreateBy = "1";
                    controrun.CreateDate = DateTime.Now;
                    controrun.RunningNumber = result;

                    dbWolf.TRNControlRunnings.InsertOnSubmit(controrun);
                    dbWolf.SubmitChanges();
                }
                else
                {
                    List<TRNControlRunning> ListRunning = dbWolf.TRNControlRunnings.Where(a => a.RunningNumber == icontrolRunning.RunningNumber && a.TemplateId == icontrolRunning.TemplateId).ToList();
                    if (ListRunning.Count > 0)
                    {

                        result = ListRunning.First().RunningNumber;
                    }
                    else
                    {
                        result = icontrolRunning.RunningNumber;
                    }

                }



            }
            catch (Exception ex)
            {
                result = string.Empty;
            }
            return result;
        }
        public static int CheckAutoNumber(DataClasses1DataContext db, string runningnumber, int templateID, int running, string prefix, string mode = "")
        {
            DataClasses1DataContext dbWolf = new DataClasses1DataContext(dbConnectionStringWolf);
            int nRunning = running;

            var mstMasterData = dbWolf.MSTMasterDatas.FirstOrDefault(x => x.MasterType == "CHK_AT_NB" & x.IsActive == true);

            if (mstMasterData != null)
            {
                var splitDocumentCode = mstMasterData.Value1.Split('|');

                var tempID_DarN = dbWolf.MSTTemplates.Where(x => splitDocumentCode.Contains(x.DocumentCode) && x.IsActive == true).ToList();


                if (tempID_DarN.Count() != 0)
                {
                    List<MSTTemplate> selectTempID = tempID_DarN.Where(x => x.TemplateId == templateID).ToList();

                    if (selectTempID.Count > 0)
                    {
                        var listRunningModel = new List<RunningModel>();
                        var itemMemo = new TRNMemo();
                        if (mode == "submit")
                        {
                            var listitemMemo = dbWolf.TRNMemos.Where(x => x.TemplateId == templateID && x.MAdvancveForm.Contains(prefix)).ToList();

                            if (listitemMemo.Count == 1)
                            {
                                itemMemo = listitemMemo.FirstOrDefault();
                            }

                            else
                            {
                                itemMemo = listitemMemo[listitemMemo.Count - 2];
                            }
                        }

                        else
                        {
                            itemMemo = dbWolf.TRNMemos.Where(x => x.TemplateId == templateID && x.MAdvancveForm.Contains(prefix)).ToList().LastOrDefault();
                        }

                        if (itemMemo != null)
                        {
                            var runningModel = new RunningModel();
                            runningModel.Status = itemMemo.StatusName;

                            if (runningModel.Status == Status._Rejected || runningModel.Status == Status._Cancelled)
                            {
                                return running - 1;
                            }

                            return running;
                        }
                    }
                }
            }

            return nRunning;
        }
        public class RunningModel
        {
            public string Status { get; set; }
            public string DocTypeCode { get; set; }
        }
        public static class Status
        {
            public static string _NewRequest = "New Request";
            public static string _WaitForRequestorReview = "Wait for Requestor Review";
            public static string _WaitForApprove = "Wait for Approve";
            public static string _WaitForComment = "Wait for Comment";
            public static string _Rework = "Rework";
            public static string _Draft = "Draft";
            public static string _Completed = "Completed";
            public static string _Rejected = "Rejected";
            public static string _Cancelled = "Cancelled";
            public static string _Pending = "Pending";
            public static string _RequestCancel = "Request Cancel";

            public static List<string> _EditableStatus = new List<string>() { Status._Draft, Status._Rework, Status._NewRequest, Status._WaitForRequestorReview };

        }
    }
}
