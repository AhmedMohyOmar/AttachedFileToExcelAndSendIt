using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Linkdev.MOE.CRM.BLL.CRMCommon;
//using Linkdev.MOE.CRM.DAL;
//using LinkDev.Libraries.Common;
using Microsoft.Xrm.Sdk;
using Moq;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
//using Linkdev.OHDREM.CustomStep.AttachFileToEmail;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using Linkdev.OHDREM.CustomStep.AttachFileToEmail;

namespace Linkdev.MOE.Test.Moq
{
    class Program
    {
        static void Main(string[] args)
        {


            #region Variables Intializations
            // { 0460d6d5 - 0b94 - e811 - 80ce - 000d3a393005}
            Guid AdminUserId = new Guid("22BF32C1-57B2-455C-AC52-BDD4B89553CC"); // crmadmin user
            //Guid AdminUserId = new Guid("66513DFC-4E23-E911-8125-000D3A393005"); //sectionhead
            
            //renewal
            //Guid PrimaryEntityID = new Guid("6CCAB8DE-B739-E911-A84E-000D3A2B2BE0");
            string PrimaryEntityName = "new_unit";

            #endregion

            //Activity activity = new CalculateFine();
            //Activity activity = new UpdateLicense();

            Activity activity = new SendMail();
            var invoker = new WorkflowInvoker(activity);

            //create our mocks
            var factoryMock = new Mock<IOrganizationServiceFactory>();
            var tracingServiceMock = new Mock<ITracingService>();
            var workflowContextMock = new Mock<IWorkflowContext>();
            //set up a mock service for CRM organization service
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;



            var connection = CrmConnection.Parse(@"Url=https://sbodre.crm4.dynamics.com/; Username=crmadmin@orascomdh.onmicrosoft.com; Password=Yabo7688;");
            //var connection2 = CrmConnection.Parse(@"AuthType=IFD;Url=https://moe.linkdev.com; HomeRealmUri=https://moe.linkdev.com/MOE/XRMServices/2011/Organization.svc;Domain=crm20160; Username=crmadmin; Password=P@ssw0rd@MOE");
            OrganizationService service = new OrganizationService(connection);

            //IOrganizationService service = new CRMAccess().GetAccessToCRM();
            //set up a mock workflowcontext


            workflowContextMock.Setup(t => t.InitiatingUserId).Returns(AdminUserId);
            workflowContextMock.Setup(t => t.CorrelationId).Returns(Guid.NewGuid());

            workflowContextMock.Setup(t => t.UserId).Returns(AdminUserId);
            //workflowContextMock.Setup(t => t.PrimaryEntityId).Returns(PrimaryEntityID);
            workflowContextMock.Setup(t => t.PrimaryEntityName).Returns(PrimaryEntityName);
            var workflowContext = workflowContextMock.Object;

            //set up a mock tracingservice - will write output to console
            tracingServiceMock.Setup(t => t.Trace(It.IsAny<string>(), It.IsAny<object[]>())).Callback<string, object[]>((t1, t2) => Console.WriteLine(t1, t2));
            var tracingService = tracingServiceMock.Object;
            //set up a mock servicefactory
            factoryMock.Setup(t => t.CreateOrganizationService(AdminUserId)).Returns(service);
            var factory = factoryMock.Object;

            invoker.Extensions.Add<ITracingService>(() => tracingService);
            invoker.Extensions.Add<IWorkflowContext>(() => workflowContext);
            invoker.Extensions.Add<IOrganizationServiceFactory>(() => factory);

            //var inputs = new Dictionary<string, object>
            //{

            //    {"paymentTransaction", new EntityReference("ldv_paymenttransaction", new Guid("{41791C92-6400-E811-82C3-000D3A24BD44}"))},
            //    {"FineAmount", new Money(77777) },
            //    { "IncrementalAmount",new Money(888)},
            //    { "SubmittedDate",new DateTime(2017,12,31)},
            //    { "ByPassFine",false},
            //    { "PaymentTotalFees",new Money(1350)},
            //    { "FinesType",new OptionSetValue(1)} //Daily

            //};

            //        enum RequestTypeEnum
            //{
            //    ChangeName = 1,
            //    ChangeManager = 2,
            //    ChangeOwner = 3,
            //    AddStage = 4,
            //    RemoveStage = 5
            //}

            //var inputs = new Dictionary<string, object>
            //    {

            //       // {"Account", new EntityReference("account", new Guid("{649F1E0D-1601-E911-811F-000D3A393005}"))}, / madonna
            //        //{"Account", new EntityReference("account", new Guid("{E935782D-4E46-E911-8156-000D3A393005}"))}, // link
            //        {"Account", new EntityReference("account", new Guid("{69D9674D-CDFE-E811-811C-000D3A393005}"))}, // nermine
            //        {"OldLicenseReference", new EntityReference("ldv_license", new Guid("{B2442930-397A-E911-81A9-000D3A393005}"))},
            //        //{"OldLicenseReference", new EntityReference("ldv_license", new Guid("{9582C502-474A-E911-8157-000D3A393005}"))},
            //        //{ "RequestType",new EntityReference("ldv_requesttype", new Guid("{886FC105-3B34-E911-8131-000D3A393005}"))} // changename
            //        //{ "RequestType",new EntityReference("ldv_requesttype", new Guid("{534914F7-3A34-E911-8131-000D3A393005}"))} // chamne manager
            //        //{ "RequestType",new EntityReference("ldv_requesttype", new Guid("{4BDDF123-3B34-E911-8131-000D3A393005}"))} // add stage
            //        //{ "RequestType",new EntityReference("ldv_requesttype", new Guid("{BDB9A432-3B34-E911-8131-000D3A393005}"))},  // remove stage
            //        //{ "RequestType",new EntityReference("ldv_requesttype", new Guid("{E3DA6A16-3B34-E911-8131-000D3A393005}"))},  // Owner stage
            //        { "RequestType","3"},  // owner stage
            //        { "RenewalPeriodDays",365} // remove stage


            //    };
            var inputs = new Dictionary<string, object>
            {

                //what is the GUID for the unit
                //villa : F5BCAEBB-CD98-E911-A85B-000D3A2B2ACB on staging
                //apartment : 8D724F91-D198-E911-A85B-000D3A2B2ACB onstaging
                {"View", new EntityReference("savedquery", new Guid("{8D724F91-D198-E911-A85B-000D3A2B2ACB}"))},
                {"Mail", new EntityReference("email", new Guid("{84E478D0-13B1-E911-A83B-000D3A29A761}"))}
            };

            //var inputs = new Dictionary<string, object> { };
            var outputs = invoker.Invoke(inputs);

        }
    }
}
