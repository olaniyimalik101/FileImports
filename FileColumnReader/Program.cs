using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;

namespace FileColumnReader
{
    class Program
    {
        //IOrganizationService service;
        static void Main(string[] args)
        {
            IOrganizationService service; //= GetOrganizationServiceClientSecret("e58f0b8e-5366-4e51-b7ed-04e975faaa18", "RHs8Q~aS.R.cHwsnxaAlBbwEA3dCuJzT5Wq48bww", "https://geniushub.crm4.dynamics.com", "https://login.microsoftonline.com/79db3a8e-ce69-4a1f-a805-1970f93dfd63");
            ////var clientId = "e58f0b8e-5366-4e51-b7ed-04e975faaa18";
            ////var clientSecret = "RHs8Q~aS.R.cHwsnxaAlBbwEA3dCuJzT5Wq48bww";
            ////var organizationUri  = "https://geniushub.crm4.dynamics.com";
            ////var authority = "https://login.microsoftonline.com/79db3a8e-ce69-4a1f-a805-1970f93dfd63";


            //var conn = new CrmServiceClient($"AuthType=ClientSecret;Url={organizationUri};ClientId={clientId};ClientSecret={clientSecret};Authority={authority};RequireNewInstance=False");
            var conn = new CrmServiceClient(@"AuthType=OAuth; Username=Shadrach@23we.onmicrosoft.com; Password=Dynamics100%; Url=https://geniushub.crm4.dynamics.com/; AppId=51f81489-12ee-4a9e-aaae-a2591f45987d; RedirectUri=app://58145B91-0C36-4500-8554-080854F2AC97;");
            service = conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            EntityReference entity = new EntityReference("sm_fileimport", new Guid("c27bee62-63a3-ef11-8a69-7c1e52506ab3"));
            byte[] downloadedFile = fileDownloader.DownloadFile(service, entity, "sm_uploadfile");

            if(downloadedFile != null)
            {
                FileHelper.ValidateFileHeaderForSelectedContentSize(downloadedFile);
            }

            var response = RecordCreationHelper.CreateRecordsInCRM(downloadedFile, service);

            // Access the response properties
            int rowCount = response.rowCount;
            int successCount = response.successCount;
            int failureCount = response.failureCount;
            List<string> failedRecordsDetails = response.failedRecordsDetails;


            RecordCreationHelper.AttachFailedRecordsToEntity(entity, service, failedRecordsDetails);

        }

    }
}
