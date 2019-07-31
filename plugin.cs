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
//using Linkdev.MOE.CustomStep.UpdateLicense;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;


namespace Linkdev.MOE.Test.Moq
{
    class plugin
    {
         void Main(string[] args)
        {

            #region Variables Intializations

            //sms
            //Guid AdminUserId = new Guid("0460D6D5-0B94-E811-80CE-000D3A393005");
            //Guid PrimaryEntityID = new Guid("ED11608A-9419-E911-8124-000D3A393005");
            //string PrimaryEntityName = "ldv_sms";

            //portal
            Guid AdminUserId = new Guid("0460D6D5-0B94-E811-80CE-000D3A393005");
            Guid PrimaryEntityID = new Guid("76AD955A-6A5B-E911-8179-000D3A393005");
            string PrimaryEntityName = "ldv_portalnotification";

            //set up mock plugincontext with input/output parameters, etc.
            Entity targetEntity = new Entity(PrimaryEntityName);
            targetEntity.Attributes["regardingobjectid"] = new EntityReference("ldv_editlicenserequest", Guid.Parse("969eac99-787c-e911-81af-000d3a393005"));
            targetEntity.Attributes["ldv_contact"] = new EntityReference("contact", Guid.Parse("b7a66d03-5a29-e911-8127-000d3a393005"));//marina nasry contact
            targetEntity.Attributes["description"] = "test ";
            targetEntity.Attributes["ldv_arabicdescription"] = "fsdsdf";

            //targetEntity.Attributes["description"] = "";
            //targetEntity.Attributes["ldv_mobilenumber"] = "01009612167"; 
            //targetEntity.Attributes["description"] = new EntityReference("ldv_servicestatus", Guid.Parse("3CDC0E8D-0A31-E811-8254-000D3A28225B"));
            targetEntity.LogicalName = PrimaryEntityName;

            //Entity PreImageEntity = new Entity(PrimaryEntityName);
            //PreImageEntity.Attributes["ldv_servicestatus"] = new EntityReference("ldv_servicestatus", Guid.Parse("0C66A0D1-D72D-E811-824F-000D3A28225B"));
            //PreImageEntity.LogicalName = PrimaryEntityName;


            #endregion

            //create our mocks
            var factoryMock = new Mock<IOrganizationServiceFactory>();
            var tracingServiceMock = new Mock<ITracingService>();
            var workflowContextMock = new Mock<IWorkflowContext>();
            var serviceMock = new Mock<IOrganizationService>();
            var notificationServiceMock = new Mock<IServiceEndpointNotificationService>();
            var pluginContextMock = new Mock<IPluginExecutionContext>();
            var serviceProviderMock = new Mock<IServiceProvider>();

            //set up a mock service for CRM organization service


            //next - create an entity object that will allow us to capture the entity record that is passed to the Create method
            Entity actualEntity = new Entity();
            Guid idToReturn = Guid.NewGuid();
            //setup the CRM service mock
            serviceMock.Setup(t =>
                t.Update(It.IsAny<Entity>()))
                //when Create is called with any entity as an invocation parameter
                //.Returns(idToReturn) //return the idToReturn guid
                .Callback<Entity>(s => targetEntity = s); //store the Create method invocation parameter for inspection later

            //IOrganizationService service  = XrmConnectionProvider.Service;
            //IOrganizationService service = serviceMock.Object;
            var connection = CrmConnection.Parse(@"Url=https://moe.linkdev.com/XRMServices/2011/Organization.svc; Username=crm20160\crmadmin; Password=P@ssw0rd@MOE;");
            OrganizationService service = new OrganizationService(connection);

            //set up a mock servicefactory using the CRM service mock
            factoryMock.Setup(t => t.CreateOrganizationService(It.IsAny<Guid>())).Returns(service);
            var factory = factoryMock.Object;

            //set up a mock tracingservice - will write output to console
            tracingServiceMock.Setup(t => t.Trace(It.IsAny<string>(), It.IsAny<object[]>())).Callback<string, object[]>((t1, t2) => Console.WriteLine(t1, t2));
            var tracingService = tracingServiceMock.Object;

            //set up mock notificationservice - not going to do anything with this
            var notificationService = notificationServiceMock.Object;  


            //IOrganizationService service = new CRMAccess().GetAccessToCRM();

            //workflowContextMock.Setup(t => t.InitiatingUserId).Returns(AdminUserId);
            //workflowContextMock.Setup(t => t.CorrelationId).Returns(Guid.NewGuid());
            //workflowContextMock.Setup(t => t.UserId).Returns(AdminUserId);
            //workflowContextMock.Setup(t => t.PrimaryEntityId).Returns(PrimaryEntityID);
            //workflowContextMock.Setup(t => t.PrimaryEntityName).Returns(PrimaryEntityName);
            //var workflowContext = workflowContextMock.Object;


            ParameterCollection inputParameters = new ParameterCollection();
            inputParameters.Add("Target", targetEntity);
            //EntityImageCollection imageCollection = new EntityImageCollection();
            //imageCollection.Add("PreImage", PreImageEntity);
            ParameterCollection outputParameters = new ParameterCollection();
            outputParameters.Add("id", PrimaryEntityID);

            pluginContextMock.Setup(t => t.InputParameters).Returns(inputParameters);
            pluginContextMock.Setup(t => t.OutputParameters).Returns(outputParameters);
            //pluginContextMock.Setup(t => t.PreEntityImages).Returns(imageCollection);
            //pluginContextMock.Setup(t => t.PostEntityImages).Returns(imageCollection);
            pluginContextMock.Setup(t => t.UserId).Returns(AdminUserId);
            pluginContextMock.Setup(t => t.PrimaryEntityName).Returns(PrimaryEntityName);
            pluginContextMock.Setup(t => t.PrimaryEntityId).Returns(PrimaryEntityID);
            pluginContextMock.Setup(t => t.MessageName).Returns("create");
            var pluginContext = pluginContextMock.Object;

            //set up a serviceprovidermock
            serviceProviderMock.Setup(t => t.GetService(It.Is<Type>(i => i == typeof(IServiceEndpointNotificationService)))).Returns(notificationService);
            serviceProviderMock.Setup(t => t.GetService(It.Is<Type>(i => i == typeof(ITracingService)))).Returns(tracingService);
            serviceProviderMock.Setup(t => t.GetService(It.Is<Type>(i => i == typeof(IOrganizationServiceFactory)))).Returns(factory);
            serviceProviderMock.Setup(t => t.GetService(It.Is<Type>(i => i == typeof(IPluginExecutionContext)))).Returns(pluginContext);
            var serviceProvider = serviceProviderMock.Object;

            //PostCreateSMS PostCreateSMS = new PostCreateSMS();
            //PostCreateSMS.Execute(serviceProvider);

            //PostCreatePortalNotification PortalNotification = new PostCreatePortalNotification();
            //PortalNotification.Execute(serviceProvider);

        }
    }
}
