using Microsoft.Crm.Sdk.Messages;
//using Microsoft.Xrm.Client;
//using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Workflow.Runtime;
using Microsoft.Xrm.Sdk.Workflow.Activities;
using System.Data;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using OfficeOpenXml;
using Microsoft.Xrm.Sdk.Messages;
using System.Xml;
using System.Web.Services.Description;
using System.Xml.Serialization;
using System.Configuration;
using System.Xml.Linq;

namespace Linkdev.OHDREM.CustomStep.AttachFileToEmail
{
    public class SendMail : CodeActivity
    {


        [RequiredArgument]
        [Input("View Reference")]
        [ReferenceTarget("savedquery")]
        public InArgument<EntityReference> View { get; set; }

        [RequiredArgument]
        [Input("Mail Reference")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Mail { get; set; }


        protected override void Execute(CodeActivityContext wfContext)
        {
            ITracingService ITracingService = wfContext.GetExtension<ITracingService>();
            IWorkflowContext context = wfContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory IOrganizationServiceFactory = wfContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = IOrganizationServiceFactory.CreateOrganizationService(context.UserId);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
                EntityReference ViewEntityRef = View.Get<EntityReference>(wfContext);
                EntityReference MailEntityRef = Mail.Get<EntityReference>(wfContext);

                AttachViewAsExcel(service, ViewEntityRef, MailEntityRef);

                SendEmailStep(MailEntityRef.Id, service); //new Guid("00988493-33E9-E811-A845-000D3A2B2BE0")
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                string message = ex.Message;
                throw;
            }
        }


        //public DataTable GetDataTable(EntityCollection EC)
        //{
        //    //EntityCollection accountRecords = GetAccountRecords();
        //    DataTable dTable = new DataTable();
        //    int iElement = 0;

        //    if (EC.Entities.Count <= 0)
        //    {
        //        return null;
        //    }

        //    //Defining the ColumnName for the datatable
        //    for (iElement = 0; iElement <= EC.Entities[0].Attributes.Count - 1; iElement++)
        //    {
        //        string columnName = EC.Entities[0].Attributes.Keys.ElementAt(iElement);
        //        dTable.Columns.Add(columnName);
        //    }

        //    foreach (Entity entity in EC.Entities)
        //    {
        //        DataRow dRow = dTable.NewRow();
        //        for (int i = 0; i <= entity.Attributes.Count - 1; i++)
        //        {
        //            string colName = entity.Attributes.Keys.ElementAt(i);
        //            dRow[colName] = entity.Attributes.Values.ElementAt(i);
        //        }
        //        dTable.Rows.Add(dRow);
        //    }
        //    return dTable;
        //}
        //public DataTable ConvertToDataTable<T>(IList<T> data)
        //{
        //    PropertyDescriptorCollection properties =
        //       TypeDescriptor.GetProperties(typeof(T));
        //    DataTable table = new DataTable();
        //    foreach (PropertyDescriptor prop in properties)
        //        table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
        //    foreach (T item in data)
        //    {
        //        DataRow row = table.NewRow();
        //        foreach (PropertyDescriptor prop in properties)
        //            row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
        //        table.Rows.Add(row);
        //    }
        //    return table;

        //}

        public byte[] ExportToExcel(DataTable tbl, string ViewName)
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                ExcelWorkbook workbook = package.Workbook;

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(ViewName);

                //column headings
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = tbl.Columns[i].ColumnName.ToString();

                }

                // rows
                for (var i = 0; i < tbl.Rows.Count; i++)
                {
                    for (var j = 0; j < tbl.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = tbl.Rows[i][j];
                    }
                }
                //using (var rng = worksheet.Cells[worksheet.Dimension.Address])
                //{
                //    rng.AutoFitColumns(0);
                //}


                //worksheet.Calculate();
                //if (worksheet.Cells.Count() > 0)
                //{
                //    //worksheet.Cells[worksheet.Dimension.Address].
                //    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns(0,50);
                //    //worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                //}


                //for (int i = 0; i < tbl.Columns.Count; i++)
                //{
                //    worksheet.Column(i).AutoFit(1,50);
                //}


                byte[] result = package.GetAsByteArray();

                // Create a new memory stream.
                //MemoryStream outStream = new MemoryStream();

                return result;

            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        //public XmlDocument xDoc
        //{
        //    get { return m_xDoc; }
        //    set { value = m_xDoc; }
        //}
        public void AttachViewAsExcel(IOrganizationService service, EntityReference ViewEntityRef, EntityReference MailEntityRef)
        {

            var queryId = ViewEntityRef;// new Guid("F5BCAEBB-CD98-E911-A85B-000D3A2B2ACB");

            var savedQuery = service.Retrieve("savedquery", queryId.Id, new ColumnSet(new[] { "fetchxml", "layoutxml", "name" }));


            var queryFetchXml = savedQuery.Attributes["fetchxml"].ToString();
            var querylayoutxml = savedQuery.Attributes["layoutxml"].ToString();
            var ViewName = savedQuery.Attributes["name"].ToString();



            EntityCollection result = service.RetrieveMultiple(new FetchExpression(queryFetchXml));

            XDocument xdc = XDocument.Parse(querylayoutxml);
            var arrNames = xdc.Root
                .Descendants("cell")
                .Select(e => e.Attribute("name")).ToArray();

            //List<Entity> to DT
            DataTable table = new DataTable();

            table = convertEntityCollectionToDataTable(result, service, arrNames);

            //DT to EXCEL
            byte[] ArryOfBytes = ExportToExcel(table, ViewName);

            Entity attachment = createAttachment(ArryOfBytes, MailEntityRef, ViewName);
            service.Create(attachment);

        }

        public void FillExcelCells(ICollection<string> keys, EntityMetadata currentEntityAliased, DataTable dt, Entity myEntity, DataRow row,DataTable dtLogical, XAttribute[] arr)
        {
            int i = 0;
            for (i = 0; i < arr.Length; i++)
            {
                foreach (var attribute in currentEntityAliased.Attributes)
                {
                    string logicalname = attribute.SchemaName;
                    string columnName = "";
                    if (attribute.DisplayName.LocalizedLabels.Count() > 0)
                        columnName = attribute.DisplayName.LocalizedLabels[0].Label;
                    else
                        columnName = "";
                    
                    if (arr[i].Value.ToLower() == attribute.SchemaName.ToLower() && dt.Columns.IndexOf(columnName) == -1 && dtLogical.Columns.IndexOf(logicalname) == -1)
                    {
                        if (attribute.AttributeType == AttributeTypeCode.Money && dt.Columns.IndexOf("Currency") == -1 && dtLogical.Columns.IndexOf("TransactionCurrencyId") == -1)
                        {
                            dtLogical.Columns.Add("TransactionCurrencyId", Type.GetType("System.String"));
                            dt.Columns.Add("Currency", Type.GetType("System.String"));
                        }
                        if(attribute.DisplayName.LocalizedLabels[0].Label.ToLower() != "url")
                        {
                            dtLogical.Columns.Add(arr[i].Value, Type.GetType("System.String"));
                            dt.Columns.Add(columnName, Type.GetType("System.String"));
                        }
                    }
                }
            }
               

            foreach (var attribute in currentEntityAliased.Attributes)
            {
                
                if (attribute.DisplayName.LocalizedLabels.Count() > 0)
                {
                    string value = "";
                    var match = keys.FirstOrDefault(stringToCheck => stringToCheck.Contains(attribute.SchemaName.ToLower()));
                    if (match == null && keys.Contains(attribute.SchemaName.ToLower()))
                    {
                        value = getValuefromAttribute(myEntity.Attributes[attribute.SchemaName.ToLower()], match);
                    }
                    else
                    if (match != null && attribute.DisplayName.LocalizedLabels[0].Label.ToLower() != "url")
                    {
                        if (keys.Contains(match))
                        {
                            if(myEntity.FormattedValues.Keys.Contains(match))
                            {
                                value = getValuefromAttribute(myEntity.Attributes[match], myEntity.FormattedValues[match].ToString());
                            }
                            else
                            {
                                value = getValuefromAttribute(myEntity.Attributes[match], match);
                            }
                        }
                    }
                    Guid newGuid;
                    if (!Guid.TryParse(value, out newGuid) && value != "")
                    {
                        string logicalname = attribute.SchemaName;
                        string columnName = attribute.DisplayName.LocalizedLabels[0].Label;

                        int countOFColumn = match.Split('.').Count();
                        if (dt.Columns.IndexOf(columnName) == -1 && dtLogical.Columns.IndexOf(logicalname) == -1 && myEntity.Attributes.Keys.Contains(match.ToLower()) && match.ToLower() == logicalname.ToLower())
                        {
                            dtLogical.Columns.Add(logicalname, Type.GetType("System.String"));
                            dt.Columns.Add(columnName, Type.GetType("System.String"));

                        }
                        else if(dt.Columns.IndexOf(columnName) == -1 &&  keys.FirstOrDefault(stringToCheck => stringToCheck.Contains(attribute.SchemaName.ToLower())).Contains('.'))
                        {
                            dtLogical.Columns.Add(logicalname, Type.GetType("System.String"));
                            dt.Columns.Add(columnName, Type.GetType("System.String"));
                        }
                        

                        decimal DecimalValue;
                        if (dtLogical.Columns.IndexOf(logicalname) != -1 && dt.Columns.IndexOf(columnName) != -1 && myEntity.Attributes.Keys.Contains(match.ToLower()))
                        {
                            if (IsNumeric(value) && decimal.TryParse(value, out DecimalValue))
                            {
                                row[columnName] = Math.Round(DecimalValue, 2);
                            }
                            else
                            {
                                row[columnName] = value;
                            }
                        }

                    }
                }
            }
        }

        public DataTable convertEntityCollectionToDataTable(EntityCollection BEC, IOrganizationService service, XAttribute[] arr)
        {
            
            RetrieveEntityRequest metaDataRequest = new RetrieveEntityRequest();
            RetrieveEntityResponse metaDataResponse = new RetrieveEntityResponse();
            metaDataRequest.EntityFilters = EntityFilters.Attributes;
            metaDataRequest.LogicalName = BEC.EntityName;
            metaDataResponse = (RetrieveEntityResponse)service.Execute(metaDataRequest);
            EntityMetadata currentEntity = metaDataResponse.EntityMetadata;

            RetrieveEntityRequest metaDataRequestAliased = new RetrieveEntityRequest();
            RetrieveEntityResponse metaDataResponseAliased = new RetrieveEntityResponse();
            metaDataRequestAliased.EntityFilters = EntityFilters.Attributes;



            DataTable dtLogical = new DataTable();
            DataTable dt = new DataTable();
            int total = BEC.Entities.Count;

            for (int i = 0; i < total; i++)
            {
                DataRow row = dt.NewRow();
                Entity myEntity = (Entity)BEC.Entities[i];
                // get transaction currency 

                //XElement  root = new XElement("Root", arr);


                var values = myEntity.Attributes.Values;
                var keys = myEntity.Attributes.Keys;

                foreach (var item in values)
                {
                    if (item.GetType().ToString() == "Microsoft.Xrm.Sdk.AliasedValue")
                    {

                        metaDataRequestAliased.LogicalName = ((AliasedValue)item).EntityLogicalName.ToString();
                        metaDataResponseAliased = (RetrieveEntityResponse)service.Execute(metaDataRequestAliased);
                        EntityMetadata currentEntityAliased = metaDataResponseAliased.EntityMetadata;

                        FillExcelCells(keys, currentEntityAliased, dt, myEntity, row, dtLogical,arr);
                    }
                    else
                    {
                        FillExcelCells(keys, currentEntity, dt, myEntity, row, dtLogical,arr);
                    }
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public bool IsNumeric(string text)
        {
            double test;
            return double.TryParse(text, out test);
        }

        private string getValuefromAttribute(object p, string result = "")
        {
            if (p.ToString() == "Microsoft.Xrm.Sdk.EntityReference")
            {
                return ((EntityReference)p).Name;
            }
            if (p.ToString() == "Microsoft.Xrm.Sdk.OptionSetValue")
            {
                //return myEntity.FormattedValues[SchemaName].ToString();
                return result;
            }
            if (p.ToString() == "Microsoft.Xrm.Sdk.Money")
            {
                return ((Money)p).Value.ToString();
            }
            if (p.ToString() == "Microsoft.Xrm.Sdk.AliasedValue")
            {
                return ((Microsoft.Xrm.Sdk.AliasedValue)p).Value.ToString();
            }
            else
            {
                return p.ToString();
            }
        }
        public Entity createAttachment(byte[] ArryOfBytes, EntityReference MailEntityRef, string ViewName)
        {
            Entity attachment = new Entity("activitymimeattachment");
            attachment["subject"] = "Test";
            string fileName = "\"" + ViewName + ".xlsx" + "\"";
            //File.WriteAllText(filePath, csv.ToString());
            
            attachment["filename"] =  fileName ;
            //string.Format("inline; " + attachment["filename"] + " ={0}", fileName);
            byte[] fileStream = ArryOfBytes;

            attachment["body"] = Convert.ToBase64String(fileStream);
            attachment["mimetype"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            attachment["attachmentnumber"] = 1;
            attachment["objectid"] = new EntityReference("email", MailEntityRef.Id);
            attachment["objecttypecode"] = "email";

            return attachment;
        }

        public void SendEmailStep(Guid emailId, IOrganizationService service)
        {

            // Use the SendEmail message to send an e-mail message.
            SendEmailRequest sendEmailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };

            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);

        }
    }
}
