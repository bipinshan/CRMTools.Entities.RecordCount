using CRMTools.Entities.RecordCount.LoginWindow;
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
        }

        /// <summary>
        /// Button to login to CRM and create a CrmService Client 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            List<EntityObj> lst = new List<EntityObj>();


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

                MessageBox.Show("Total Record Count :" +totalCount.ToString());


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

                lstAttr=lstAttr.OrderBy(o => o.AttributeName).ToList();
                this.comboBox1.DisplayMemberPath = "AttributeName";
                this.comboBox1.SelectedValuePath = "AttributeSchemaName";
                this.comboBox1.ItemsSource = lstAttr;
                this.comboBox1.Text = "Select Attribute Name";

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Message: "+ex.Message);
            }
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
}
