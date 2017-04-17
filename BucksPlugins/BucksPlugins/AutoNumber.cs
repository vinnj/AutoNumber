using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace BucksPlugins
{
    public class AutoNumber : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            Microsoft.Xrm.Sdk.IPluginExecutionContext context = (Microsoft.Xrm.Sdk.IPluginExecutionContext)
                serviceProvider.GetService(typeof(Microsoft.Xrm.Sdk.IPluginExecutionContext));

            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(
            typeof(IOrganizationServiceFactory));

            //get service object from service factory
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];
                //</snippetAccountNumberPlugin2>

                // Verify that the target entity represents an account.
                // If not, this plug-in was not registered correctly.

                /*Entity retrievedautoNumber = (Entity)service.Retrieve("bnu_autonumber",
                       new Guid("241563DF-3820-E711-80FD-5065F38A1B01"), new ColumnSet(new string[] { "bnu_initialnumber", "bnu_prefix",
                       "bnu_entityname", "bnu_conditionaloptionset", "bnu_optionsetvaluenumerical", "bnu_entityname", "bnu_recentnumber",
                       "bnu_updatefield" }));
                */

                QueryExpression query = new QueryExpression
                {
                    EntityName = "bnu_autonumber",
                    ColumnSet = new ColumnSet("bnu_initialnumber", "bnu_prefix",
                       "bnu_entityname", "bnu_conditionaloptionset", "bnu_optionsetvaluenumerical", "bnu_entityname", "bnu_recentnumber",
                       "bnu_updatefield", "bnu_leadingzeros")
                };

                EntityCollection entityCollection = service.RetrieveMultiple(query);

                Entity autoNumberupdate = new Entity("bnu_autonumber");

                string relationValue = null;
                int conditionOptionsetvalue = 0;
                int initialNumber = 0;
                string prefix = null;
                string updatedField = null;
                Guid recordId = new Guid();
                int leadingzeros = 0;
                int recentNumber = 0;

                foreach (Entity result in entityCollection.Entities)
                {
                   
                    relationValue = result.GetAttributeValue<string>("bnu_conditionaloptionset"); //customertypecode
                    conditionOptionsetvalue = result.GetAttributeValue<int>("bnu_optionsetvaluenumerical"); //22
                    initialNumber = Int32.Parse(result.GetAttributeValue<string>("bnu_initialnumber")); //0001
                    prefix = result.GetAttributeValue<string>("bnu_prefix").ToString(); //PAR-
                    updatedField = result.GetAttributeValue<string>("bnu_updatefield");
                    leadingzeros = result.GetAttributeValue<int>("bnu_leadingzeros");
                    recentNumber = Int32.Parse(result.GetAttributeValue<string>("bnu_recentnumber"));
                    recordId = result.Id;

                    bool customerTypeCodeFromEntity = entity.Contains(relationValue);

                    if (customerTypeCodeFromEntity)
                    {
                        int conditionOptionsetValuefromEntity = ((OptionSetValue)entity[relationValue]).Value;

                        if (conditionOptionsetValuefromEntity == conditionOptionsetvalue)
                        {
                            if (recentNumber != 0)
                            {
                                int incrementNumber = recentNumber + 1;

                                string incrementedNumber = (recentNumber + 1).ToString();

                                string allocatedNumber = prefix + incrementedNumber.PadLeft(leadingzeros, '0');

                                entity.Attributes.Add(updatedField, allocatedNumber);

                                Entity autoNumber = new Entity("bnu_autonumber");

                                autoNumber.Id = recordId;

                                autoNumber["bnu_recentnumber"] = incrementedNumber;

                                service.Update(autoNumber);

                                break;
                            }

                            else
                            {
                                int incrementNumber = initialNumber + 1;

                                string incrementedNumber = incrementNumber.ToString();

                                string allocatedNumber = prefix + incrementedNumber.PadLeft(leadingzeros, '0');

                                entity.Attributes.Add(updatedField, allocatedNumber);

                                Entity autoNumber = new Entity("bnu_autonumber");

                                autoNumber.Id = recordId;

                                autoNumber["bnu_recentnumber"] = incrementedNumber;

                                service.Update(autoNumber);

                                break;
                            }

                            
                        }
                        
                    }
                                     
                }

            }
        }
    }
}

