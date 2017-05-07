using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Crm.Sdk.Messages;

namespace UpdateCRM
{
    public class Academics
    {
        private IOrganizationService _orgService;
        public Guid RetrievedContactID { get; set; }
        public string email { get; set; }

        #region RetrieveUpdateaAcademics
        /*METHOD TO RETRIEVE XL DATA AND PASSING IT TO RetrieveUpdate METHOD
        THE METHOD CALLS UPDATEHELPER CLASS WHICH RETRIEVES THE XL COLUMN DATA
        THEN PASSES THE XL DCOLUMN DATA TO RETRIEVEUPDATE*/
        public void RetrieveUpdateaAcademics()
        {
            UpdateHelper.UpdateAcademics academicHelper = new UpdateHelper.UpdateAcademics();
            var academicExcel = academicHelper.ReadAcademicData();

            var academics = from p in academicExcel.Worksheet<UpdateHelper.UpdateAcademics>("Academics")
                            select p;

            foreach (var contact in academics)
            {
                email = contact.Email;
                RetrievedContactID = new Guid(contact.ContactID);
                Academic(email, RetrievedContactID);
            }



        }
        #endregion

        #region RetrieveUpdateaAcademics
        /*METHOD TO CONNECT WITH CRM,RETRIEVE & UPDATE DATA.
        RETRIEVE DATA VIA LINQ QUERY
        PASSING EMAIL/CONTACT ID TO UPDATE CRM RECORD
        */
        private void Academic(string email, Guid Retrieved, String connectionString, bool promptforDelete)
        {
             
             //string connStr = ConfigurationManager.ConnectionStrings[0].ConnectionString;

            CrmServiceClient conn = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connectionString);

            {
                //Guid contactID;
               

                BUCKSdev svcContext = new BUCKSdev();

                var contacts = from c in svcContext.ContactSet
                               where c.ContactId == RetrievedContactID
                               select c;

                foreach (var contact in contacts)
                {
                    Contact updateContact = new Contact
                    {
                        ContactId = RetrievedContactID,
                        EMailAddress1 = email

                    };

                    //_serviceProxy.Update(updateContact);
                    Console.WriteLine("Academic {0} updated", RetrievedContactID);

                }

            }
        }
        #endregion
    }
}
