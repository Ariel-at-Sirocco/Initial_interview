using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Configuration;
using System.Linq;


namespace ConsoleApplication
{
    internal class Utils
    {
        /* <summary>
         *  Internal class with static methods with useful funciontality 
         * </summary>   
         */


        // Const definition

        //app configuration file
        const string APP_CONFIG_FILE = "App.config";
        //setting keys in config file
        const string URL_KEY         = "url";
        const string USER_NAME_KEY   = "userName";
        const string PASSWORD_KEY    = "password";

        //static private methods
        static private string GetConnectionString(string url, string userName, string password)
        {
            /**
             * <summary>
             *  Generates connection string with connection information. 
             * </summary>           
             */
            string connectionString = $@"
                                    AuthType = 'OAuth';
                                    Url = {url};
                                    UserName = {userName};
                                    Password = {password};
                                    RedirectUri = http://localhost;
                                    LoginPrompt = Auto;
                                    RequireNewInstance = True";

            return (connectionString);
        }

        // static public methods
        static public string GetAppConfigPassword()
        {
            /**
            * <summary>
            *  Reads App.config files and returns Password info. 
            * </summary>           
            */

            return (ConfigurationManager.AppSettings[PASSWORD_KEY]);
        }

        static public string GetAppConfigURL()
        {
            /**
             * <summary>
             *  Reads App.config files and returns URL info.  
             * </summary>           
             */
            
            return (ConfigurationManager.AppSettings[URL_KEY]);
        }
        static public string GetAppConfigUserName()
        {
            /**
             * <summary>
             *  Reads App.config files and returns UserName info. 
             * </summary>           
             */

            return (ConfigurationManager.AppSettings[USER_NAME_KEY]);
        }

        static public ServiceClient GetServiceClient(string url, string userName, string password)
        {
            /**
            * <summary>
            * returns ServiceClient related to the provided endpoint information.
            * </summary>           
            */

            string connectionString = GetConnectionString(url, userName, password);

            return (new ServiceClient(connectionString));
        }

        static public EntityReferenceCollection GetAccountsReferencesByName(string accountName, ServiceClient svc)
        {
            /*
             * returns those account references defined by account name, using provided Service Client instance
             */
            
            // Get accounts
            var query = new QueryByAttribute("account")
            {
                ColumnSet = new ColumnSet("name")
            };
            query.AddAttributeValue("name", accountName);
            EntityCollection accountsCollection = svc.RetrieveMultiple(query);

            //Convert the EntityCollection to a EntityReferenceCollection for Child Account
            EntityReferenceCollection accountReferencesCollection = new EntityReferenceCollection();

            accountsCollection.Entities.ToList().ForEach(x => {
                accountReferencesCollection.Add(x.ToEntityReference());
            });

            return (accountReferencesCollection);
        }
    }

    class Program
    {
        //obtained from App.config
        //static string USER_NAME = "mail@mail.com";
        //static string PASSWORD  = "pass";
        //static string URL       = "url"; 

        //info for Entities
        private const string ACCOUNT_ENTITY = "account";
        private const string CONTACT_ENTITY = "contact";
        private const string NOTE_ENTITY    = "annotation";

        // Relationship keys:
        private const string ACCOUNT_PARENT_ACCOUNT_REL  = "account_parent_account";
        private const string ACCOUNT_PRIMARY_CONTACT_REL = "account_primary_contact";
        private const string NOTE_ACCOUNT_REL = "objectid_account";
        private const string NOTE_CONTACT_REL = "objectid_account";

        static void Main(string[] args)
        {
            try
            {
                /***
                 *  Step 5: connect to CRM
                 */
                Console.WriteLine("-----------------");
                Console.WriteLine("*    STEP 5     *");
                Console.WriteLine("-----------------" + '\n');

                ServiceClient serviceClient = Utils.GetServiceClient(Utils.GetAppConfigURL(),
                                                                     Utils.GetAppConfigUserName(),
                                                                     Utils.GetAppConfigPassword()
                                                                     );

                if (serviceClient.IsReady)
                {
                    WhoAmIResponse response = (WhoAmIResponse)serviceClient.Execute(new WhoAmIRequest());
                    
                    /***
                     *  Step 7: Create 2 account and parent-child rel. use Late Binding.
                     *  In Production environment I would have modeled and included entities in 
                     *  JSON structures for processing them in a loop. Also for performance concerns                 
                     */
                    Console.WriteLine("\n");
                    Console.WriteLine("-----------------");
                    Console.WriteLine("*    STEP 7     *");
                    Console.WriteLine("-----------------" + '\n');

                    // Parent account  
                    var parentAccount = new Entity();
                    //set entity logical name  & values
                    parentAccount.LogicalName = ACCOUNT_ENTITY;
                    parentAccount["name"]     = "Rock Inc";

                    // child account
                    var childAccount = new Entity();
                    //set entity logical name & values
                    childAccount.LogicalName = ACCOUNT_ENTITY;
                    childAccount["name"]     = "Beatles Inc";

                    // Add the child account to an EntityCollection
                    EntityCollection childAccountCollection = new EntityCollection();
                    childAccountCollection.Entities.Add(childAccount);

                    // Set  relationship
                    Relationship accountRelationship      = new Relationship(ACCOUNT_PARENT_ACCOUNT_REL);
                    accountRelationship.PrimaryEntityRole = EntityRole.Referenced;
                    parentAccount.RelatedEntities.Add(accountRelationship, 
                                                      childAccountCollection
                                                      );

                    // Create the related accounts in CRM
                    Guid parentAccountGuid = serviceClient.Create(parentAccount);

                    Console.WriteLine("done");

                    /***
                    *  Setp 8: Create two contacts and associate the first one with the first account and the second one with the second account                
                    *  
                    */
                    Console.WriteLine("\n");
                    Console.WriteLine("-----------------");
                    Console.WriteLine("*    STEP 8     *");
                    Console.WriteLine("-----------------" + '\n');

                    // Parent account contact    
                    var parentContact = new Entity();
                    //set entity logical name & value
                    parentContact.LogicalName  = CONTACT_ENTITY;
                    parentContact["firstname"] = "Jack";
                    parentContact["lastname"]  = "Black";

                    // Child account contact    
                    var childContact = new Entity();
                    //set entity logical name
                    childContact.LogicalName  = CONTACT_ENTITY;
                    childContact["firstname"] = "Ringo";
                    childContact["lastname"]  = "Star";

                    //Create contacts
                    Guid parentContactGuid = serviceClient.Create(parentContact);
                    Guid childContactGuid  = serviceClient.Create(childContact);

                    // get Account reference by name    
                    EntityReferenceCollection parentAccountReferenceCollection = Utils
                        .GetAccountsReferencesByName((string)parentAccount["name"],
                                                             serviceClient
                                                    );

                    EntityReferenceCollection childAccountReferenceCollection = Utils
                        .GetAccountsReferencesByName((string)childAccount["name"],
                                                             serviceClient
                                                    );

                    // Use the Associate method for association to parent account
                    AssociateRequest associationRequest = new AssociateRequest()
                    {
                        RelatedEntities = parentAccountReferenceCollection,
                        Relationship    = new Relationship(ACCOUNT_PRIMARY_CONTACT_REL),
                        Target          = new EntityReference(CONTACT_ENTITY, parentContactGuid)
                    };
                    serviceClient.Execute(associationRequest);

                    // Use the Associate method for association to child account
                    associationRequest = new AssociateRequest()
                    {
                        RelatedEntities = childAccountReferenceCollection,
                        Relationship    = new Relationship(ACCOUNT_PRIMARY_CONTACT_REL),
                        Target          = new EntityReference(CONTACT_ENTITY, childContactGuid)
                    };
                    serviceClient.Execute(associationRequest);

                    Console.WriteLine("done");

                    /***
                     *  Step 9. update fields in one account (parent) and one contact (parent contact             
                     *  
                     */
                    Console.WriteLine("\n");
                    Console.WriteLine("-----------------");
                    Console.WriteLine("*    STEP 9     *");
                    Console.WriteLine("-----------------" + '\n');

                    // update parent account 
                    //get old values
                    Entity retrievedAccount = serviceClient.Retrieve(ACCOUNT_ENTITY,
                                                                    parentAccountGuid,
                                                                    new ColumnSet("name")
                                                                    );
                    Entity modifiedAccount  = new Entity(ACCOUNT_ENTITY);
                    modifiedAccount.Id      = parentAccountGuid;
                    modifiedAccount["name"] = retrievedAccount["name"];
                    //new value
                    modifiedAccount["numberofemployees"] = 500;
                    serviceClient.Update(modifiedAccount);

                    // update parent contact 
                    //get old values
                    Entity retrievedContact = serviceClient.Retrieve(CONTACT_ENTITY,
                                                                     parentContactGuid,
                                                                     new ColumnSet("firstname","lastname")
                                                                     );

                    Entity modifiedContact       = new Entity(CONTACT_ENTITY);
                    modifiedContact.Id           = parentContactGuid;
                    modifiedContact["firstname"] = retrievedContact["firstname"];
                    //new value
                    modifiedContact["lastname"] = retrievedContact["lastname"] + " (updated)";
                    serviceClient.Update(modifiedContact);

                    Console.WriteLine("done");

                    /*** 
                     * Step 10: Create a note and associate it with the parent account
                     */
                    Console.WriteLine("\n");
                    Console.WriteLine("----------------------");
                    Console.WriteLine("*    STEP 10  & 11   *");
                    Console.WriteLine("----------------------" + '\n');

                    // Note for parent account
                    Entity noteForParentAccount = new Entity();
                    //set entity logical name
                    noteForParentAccount.LogicalName = NOTE_ENTITY;
                    noteForParentAccount["subject"]  = "Hey";
                    noteForParentAccount["notetext"] = "this is note 1";
                    //reference from annotation to account
                    noteForParentAccount["objectid"]       = new EntityReference(ACCOUNT_ENTITY, parentAccountGuid);
                    noteForParentAccount["objecttypecode"] = ACCOUNT_ENTITY;

                    /***
                     * Setp  11.	Create two notes and associate them with the second contact
                     * 
                     * all notes all created at once usign Execute Multiple Request object 
                     * (due to performance concerns & demo purpose)
                     * 
                     */

                    // Note 1 for child contact
                    Entity noteForChildContact1 = new Entity();
                    //set entity logical name
                    noteForChildContact1.LogicalName = NOTE_ENTITY;
                    noteForChildContact1["subject"]  = "Note 1";
                    noteForChildContact1["notetext"] = "this is note 1 for Child contact";
                    //reference from annotation to account
                    noteForChildContact1["objectid"]       = new EntityReference(CONTACT_ENTITY, childContactGuid);
                    noteForChildContact1["objecttypecode"] = CONTACT_ENTITY;

                    // Note 2 for child contact
                    Entity noteForChildContact2 = new Entity();
                    //set entity logical name
                    noteForChildContact2.LogicalName = NOTE_ENTITY;
                    noteForChildContact2["subject"]  = "Note 2";
                    noteForChildContact2["notetext"] = "this is note 2 for Child contact";
                    //reference from annotation to account
                    noteForChildContact2["objectid"]       = new EntityReference(CONTACT_ENTITY, childContactGuid);
                    noteForChildContact2["objecttypecode"] = CONTACT_ENTITY;

                    // Create an ExecuteMultipleRequest object.
                    ExecuteMultipleRequest notesCreationRequest = new ExecuteMultipleRequest()
                    {
                        // Assign settings that define execution behavior: continue on error, return responses. 
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = true
                        },
                        // Create an empty organization request collection.
                        Requests = new OrganizationRequestCollection()
                    };

                    // add note entities in a collection. 
                    EntityCollection notesCollection = new EntityCollection();
                    notesCollection.Entities.Add(noteForParentAccount);
                    notesCollection.Entities.Add(noteForChildContact1);
                    notesCollection.Entities.Add(noteForChildContact2);

                    // Add a CreateRequest for each entity to the request collection.
                    foreach (var note in notesCollection.Entities)
                    {
                        CreateRequest createRequest = new CreateRequest { Target = note };
                        notesCreationRequest.Requests.Add(createRequest);
                    }

                    // Execute all the requests in the request collection using a single web method call.
                    ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)serviceClient
                        .Execute(notesCreationRequest);

                    Console.WriteLine("done");

                    /***
                     * 12.	Query the database for all contacts and all accounts and all notes. This should be done in one query. Create a list containing “name” (account or contact) and “notetext”. 
                     */
                    Console.WriteLine("\n");
                    Console.WriteLine("-----------------");
                    Console.WriteLine("*    STEP 12     *");
                    Console.WriteLine("-----------------" + '\n');

                    using (OrganizationServiceContext orgServiceContext = new OrganizationServiceContext(serviceClient))
                    {
                        // qhelp contains name and notetext
                        var qhelp = (from con in orgServiceContext.CreateQuery(CONTACT_ENTITY)
                                     join note in orgServiceContext.CreateQuery(NOTE_ENTITY)
                                       on con["contactid"] equals note["objectid"]
                                     select new
                                     {
                                         name     = con["firstname"],  
                                         notetext = note["notetext"]
                                     }).Distinct();
                    
                    }

                    Console.WriteLine("done");
                    Console.WriteLine("-- END ---");
                }
                else
                {
                    Console.WriteLine("A web service connection was not established.");
                }

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // Pause the console so it does not close.
                Console.WriteLine("Press any key to exit.");
                Console.ReadLine();
            };
        }
    }
}


