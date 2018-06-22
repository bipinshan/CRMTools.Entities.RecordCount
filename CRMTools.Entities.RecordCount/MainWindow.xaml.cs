using CRMTools.Entities.RecordCount.LoginWindow;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CRMTools.Entities.RecordCount
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        CrmServiceClient svcClient;
        public MainWindow()
        {
            InitializeComponent();
            this.label.Visibility = Visibility.Hidden;
            this.label1.Visibility = Visibility.Hidden;
            this.label2.Visibility = Visibility.Hidden;
            this.comboBox.Visibility = Visibility.Hidden;
            this.comboBox1.Visibility = Visibility.Hidden;
            this.listView.Visibility = Visibility.Hidden;
            this.listView1.Visibility = Visibility.Hidden;
            this.listView2.Visibility = Visibility.Hidden;
            this.listView3.Visibility = Visibility.Hidden;
            this.button.Visibility = Visibility.Hidden;
            this.label3.Visibility = Visibility.Hidden;
            this.comboBox2.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// Button to login to CRM and create a CrmService Client 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            List<EntityObj> lst = new List<EntityObj>();

            List<SystemUser> lstActiveUsr = new List<SystemUser>();


            #region Login Control
            // Establish the Login control
            CrmLogin ctrl = new CrmLogin();
            // Wire Event to login response. 
            ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
            // Show the dialog. 
            ctrl.ShowDialog();

            // Handel return. 
            if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                MessageBox.Show("Good Connect");
            else
                MessageBox.Show("BadConnect");

            #endregion

            #region CRMServiceClient
            if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            {
                svcClient = ctrl.CrmConnectionMgr.CrmSvc;
                ctrl.Close();
                this.loginbutton.Visibility = Visibility.Hidden;


                if (svcClient.IsReady)
                {
                    this.label.Visibility = Visibility.Visible;
                    this.comboBox.Visibility = Visibility.Visible;

                    //RetrieveEntityRequest req = new RetrieveEntityRequest()
                    //{
                    //    EntityFilters = EntityFilters.All,
                    //    RetrieveAsIfPublished=true,
                    //    LogicalName="account"
                    //};

                    //RetrieveEntityResponse res = (RetrieveEntityResponse)svcClient.OrganizationServiceProxy.Execute(req);



                    RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest()
                    {
                        EntityFilters = EntityFilters.Entity,
                        RetrieveAsIfPublished = true
                    };

                    // Retrieve the MetaData.
                    RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)svcClient.OrganizationServiceProxy.Execute(request);

                    foreach (EntityMetadata currentEntity in response.EntityMetadata)
                    {

                        //if (currentEntity.DisplayName.UserLocalizedLabel!=null)
                        //{
                        EntityObj obj = new EntityObj();
                        obj.EntityName = currentEntity.LogicalName;
                        obj.EntityOC = currentEntity.ObjectTypeCode.ToString();

                        lst.Add(obj);
                        //}
                    }

                    lst = lst.OrderBy(o => o.EntityName).ToList();
                    this.comboBox.DisplayMemberPath = "EntityName";
                    this.comboBox.SelectedValuePath = "EntityOC";
                    this.comboBox.ItemsSource = lst;
                    this.comboBox.Text = "Select Entity Name";

                    this.label3.Visibility = Visibility.Visible;
                    this.comboBox2.Visibility = Visibility.Visible;

                    DataCollection<Entity> activeUsrsColl = MainWindow.FindEnabledUsers(svcClient.OrganizationServiceProxy);

                    foreach (Entity user in activeUsrsColl)
                    {
                        SystemUser sysUser = new SystemUser();
                        sysUser.FullName = user["fullname"].ToString();
                        sysUser.SystemUserID = user["systemuserid"].ToString();

                        lstActiveUsr.Add(sysUser);
                    }

                    lstActiveUsr = lstActiveUsr.OrderBy(o => o.FullName).ToList();
                    this.comboBox2.DisplayMemberPath = "FullName";
                    this.comboBox2.SelectedValuePath = "SystemUserID";
                    this.comboBox2.ItemsSource = lstActiveUsr;
                    this.comboBox2.Text = "Select Active User Name";

                    // Get data from CRM . 
                    //string FetchXML =
                    //    @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //                    <entity name='account'>
                    //                        <attribute name='name' />
                    //                        <attribute name='primarycontactid' />
                    //                        <attribute name='telephone1' />
                    //                        <attribute name='accountid' />
                    //                        <order attribute='name' descending='false' />
                    //                      </entity>
                    //                    </fetch>";

                    //var Result = svcClient.GetEntityDataByFetchSearchEC(FetchXML);
                    //if (Result != null)
                    //{
                    //    MessageBox.Show(string.Format("Found {0} records\nFirst Record name is {1}", Result.Entities.Count, Result.Entities.FirstOrDefault().GetAttributeValue<string>("name")));
                    //}


                    //// Core API using SDK OOTB 
                    //CreateRequest req = new CreateRequest();
                    //Entity accENt = new Entity("account");
                    //accENt.Attributes.Add("name", "TESTFOO");
                    //req.Target = accENt;
                    //CreateResponse res = (CreateResponse)svcClient.OrganizationServiceProxy.Execute(req);
                    ////CreateResponse res = (CreateResponse)svcClient.ExecuteCrmOrganizationRequest(req, "MyAccountCreate");
                    //MessageBox.Show(res.id.ToString());



                    //// Using Xrm.Tooling helpers. 
                    //Dictionary<string, CrmDataTypeWrapper> newFields = new Dictionary<string, CrmDataTypeWrapper>();
                    //// Create a new Record. - Account 
                    //newFields.Add("name", new CrmDataTypeWrapper("CrudTestAccount", CrmFieldType.String));
                    //Guid guAcctId = svcClient.CreateNewRecord("account", newFields);

                    //MessageBox.Show(string.Format("New Record Created {0}", guAcctId));
                }
            }
            #endregion


        }

        /// <summary>
        /// Raised when the login form process is completed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CrmLogin)
            {
                this.Dispatcher.Invoke(() =>
                {
                    ((CrmLogin)sender).Close();
                });
            }
        }

        private void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int totalCount = 0;
            try
            {
                EntityObj obj = (EntityObj)this.comboBox.SelectedItem;
                QueryExpression query = new QueryExpression(obj.EntityName);
                query.ColumnSet = new ColumnSet(true);
                query.Distinct = true;
                //query.ColumnSet.AddColumn(obj.EntityOC);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;

                EntityCollection entityCollection = svcClient.OrganizationServiceProxy.RetrieveMultiple(query);
                totalCount = entityCollection.Entities.Count;

                while (entityCollection.MoreRecords)
                {
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                    entityCollection = svcClient.OrganizationServiceProxy.RetrieveMultiple(query);
                    totalCount = totalCount + entityCollection.Entities.Count;
                }

                MessageBox.Show("Total Record Count :" + totalCount.ToString());


                // Retrieve Views
                //<snippetWorkWithViews2>
                QueryExpression mySavedQuery = new QueryExpression
                {
                    ColumnSet = new ColumnSet("savedqueryid", "name", "querytype", "isdefault", "returnedtypecode", "isquickfindquery"),
                    EntityName = "savedquery",
                    Criteria = new FilterExpression
                    {
                        Conditions =
            {
                //new ConditionExpression
                //{
                //    AttributeName = "querytype",
                //    Operator = ConditionOperator.Equal,
                //    Values = {0}
                //},
                new ConditionExpression
                {
                    AttributeName = "returnedtypecode",
                    Operator = ConditionOperator.Equal,
                    Values = { obj.EntityName }
                }
            }
                    }
                };
                RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = mySavedQuery };

                RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)svcClient.OrganizationServiceProxy.Execute(retrieveSavedQueriesRequest);

                DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
                this.label1.Visibility = Visibility.Visible;
                this.comboBox1.Visibility = Visibility.Visible;
                this.label2.Visibility = Visibility.Visible;
                this.listView.Visibility = Visibility.Visible;
                this.button.Visibility = Visibility.Visible;

                //Display the Retrieved views
                ArrayList viewList = new ArrayList();

                foreach (Entity ent in savedQueries)
                {
                    viewList.Add(ent["name"].ToString());
                }

                viewList.Sort();
                listView.ItemsSource = viewList;

                RetrieveEntityRequest reqAttr = new RetrieveEntityRequest();
                reqAttr.EntityFilters = EntityFilters.Attributes;
                reqAttr.LogicalName = obj.EntityName;
                reqAttr.RetrieveAsIfPublished = true;

                RetrieveEntityResponse resAttr = (RetrieveEntityResponse)svcClient.OrganizationServiceProxy.Execute(reqAttr);
                EntityMetadata metaAttr = resAttr.EntityMetadata;

                List<AttributeObj> lstAttr = new List<AttributeObj>();
                foreach (AttributeMetadata attrMeta in metaAttr.Attributes)
                {
                    AttributeObj attrobj = new AttributeObj();
                    attrobj.AttributeName = attrMeta.SchemaName;
                    attrobj.AttributeSchemaName = attrMeta.LogicalName;

                    lstAttr.Add(attrobj);
                }

                lstAttr = lstAttr.OrderBy(o => o.AttributeName).ToList();
                this.comboBox1.DisplayMemberPath = "AttributeName";
                this.comboBox1.SelectedValuePath = "AttributeSchemaName";
                this.comboBox1.ItemsSource = lstAttr;
                this.comboBox1.Text = "Select Attribute Name";

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Message: " + ex.Message);
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.listView1.Visibility = Visibility.Visible;
                List<WebResource> lstWebRes = new List<WebResource>();
                lstWebRes = MainWindow.GetWebResources(svcClient.OrganizationServiceProxy);
                lstWebRes = lstWebRes.OrderBy(o => o.DisplayName).ToList();
                this.listView1.ItemsSource = lstWebRes;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error Message: " + ex.Message);
            }


        }

        public static List<WebResource> GetWebResources(IOrganizationService OrgService)
        {
            var fetchQuery = @"<fetch mapping='logical' version='1.0'>
                        <entity name='webresource'>
                            <attribute name='name' />
                            <attribute name='displayname' />
                            <attribute name='webresourcetype' />
                        </entity>
                    </fetch>";

            EntityCollection result = OrgService.RetrieveMultiple(new FetchExpression(fetchQuery));
            List<WebResource> WebResourceList = new List<WebResource>();

            foreach (var webresource in result.Entities)
            {
                WebResourceList.Add(new WebResource() { WebResourceName = webresource.Attributes["name"].ToString(), Type = ((OptionSetValue)webresource.Attributes["webresourcetype"]).Value.ToString(), DisplayName = webresource.Attributes["displayname"].ToString() });
            }

            return WebResourceList;
        }

        public static DataCollection<Entity> FindEnabledUsers(IOrganizationService service)
        {
            try
            {
                // Query to retrieve other users.
                QueryExpression querySystemUser = new QueryExpression
                {
                    EntityName = "systemuser",
                    ColumnSet = new ColumnSet(new String[] { "systemuserid", "fullname" }),
                    Criteria = new FilterExpression()
                };

                querySystemUser.Criteria.AddCondition("isdisabled",
                    ConditionOperator.Equal, "0");
                // Excluding SYSTEM user.
                querySystemUser.Criteria.AddCondition("lastname",
                    ConditionOperator.NotEqual, "SYSTEM");
                // Excluding INTEGRATION user.
                querySystemUser.Criteria.AddCondition("lastname",
                    ConditionOperator.NotEqual, "INTEGRATION");

                return service.RetrieveMultiple(querySystemUser).Entities;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void comboBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();

            ArrayList lstUserRoles = new ArrayList();
            SystemUser objSysUsr = new SystemUser();
            objSysUsr = (SystemUser)this.comboBox2.SelectedItem;
            //changeRequest.Target = new EntityReference("systemuser", Guid.Parse(objSysUsr.SystemUserID));

            //RetrieveRecordChangeHistoryResponse changeResponse = (RetrieveRecordChangeHistoryResponse)svcClient.OrganizationServiceProxy.Execute(changeRequest);

            QueryExpression queryExpression = new QueryExpression();
            queryExpression.EntityName = "role"; //role entity name
            ColumnSet cols = new ColumnSet();
            cols.AddColumn("name"); //We only need role name
            queryExpression.ColumnSet = cols;
            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = "systemuserid";
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(objSysUsr.SystemUserID);
            //system roles
            LinkEntity lnkEntityRole = new LinkEntity();
            lnkEntityRole.LinkFromAttributeName = "roleid";
            lnkEntityRole.LinkFromEntityName = "role"; //FROM
            lnkEntityRole.LinkToEntityName = "systemuserroles";
            lnkEntityRole.LinkToAttributeName = "roleid";
            //system users
            LinkEntity lnkEntitySystemusers = new LinkEntity();
            lnkEntitySystemusers.LinkFromEntityName = "systemuserroles";
            lnkEntitySystemusers.LinkFromAttributeName = "systemuserid";
            lnkEntitySystemusers.LinkToEntityName = "systemuser";
            lnkEntitySystemusers.LinkToAttributeName = "systemuserid";
            lnkEntitySystemusers.LinkCriteria = new FilterExpression();
            lnkEntitySystemusers.LinkCriteria.Conditions.Add(ce);
            lnkEntityRole.LinkEntities.Add(lnkEntitySystemusers);
            queryExpression.LinkEntities.Add(lnkEntityRole);
            EntityCollection entColRoles = svcClient.OrganizationServiceProxy.RetrieveMultiple(queryExpression);
            if (entColRoles != null && entColRoles.Entities.Count > 0)
            {
                foreach (Entity entRole in entColRoles.Entities)
                {
                    lstUserRoles.Add(entRole["name"].ToString());
                }
            }

            lstUserRoles.Sort();
            
            this.listView2.Visibility = Visibility.Visible;
            this.listView2.ItemsSource = lstUserRoles;

            QueryExpression query = new QueryExpression("team");
            query.ColumnSet = new ColumnSet(true);
            LinkEntity link = query.AddLink("teammembership", "teamid", "teamid");
            link.LinkCriteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, objSysUsr.SystemUserID));

            ArrayList lstTeam = new ArrayList();
            EntityCollection entColTeams = svcClient.OrganizationServiceProxy.RetrieveMultiple(query);
            foreach (Entity entTeams in entColTeams.Entities)
            {
                lstTeam.Add(entTeams["name"].ToString());
            }

            lstTeam.Sort();
            this.listView3.Visibility = Visibility.Visible;
            this.listView3.ItemsSource = lstTeam;
        }
    }

    public class EntityObj
    {
        public string EntityName { get; set; }
        public string EntityOC { get; set; }
    }

    public class AttributeObj
    {
        public string AttributeName { get; set; }
        public string AttributeSchemaName { get; set; }
    }

    public class WebResource
    {
        public string WebResourceName { get; set; }
        public string Type { get; set; }
        public string DisplayName { get; set; }

    }

    public class SystemUser
    {
        public string FullName { get; set; }
        public string SystemUserID { get; set; }
    }
}
