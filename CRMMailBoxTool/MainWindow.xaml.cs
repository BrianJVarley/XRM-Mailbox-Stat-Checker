using CRMMailBoxTool.LoginWindow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace CRMMailBoxTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private List<Entity> QueryResult;
        private int CaseSwitch = 0;
        private CrmLogin Ctrl;


        private ObservableCollection<string> _resultList;
        public ObservableCollection<string> ResultList
        {
            get
            {
                return this._resultList;
            }
            set
            {
                this._resultList = value;
            }

        }


        private string _organizationName;
        public string OrganizationName
        {
            get
            {
                return this._organizationName;
            }
            set
            {
                this._organizationName = value;
            }

        }
        

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            // Establish the Login control
            ResultList = new ObservableCollection<string>();
            InitQueryList();
            Ctrl = new CrmLogin();

        }


        private void InitQueryList()
        {
            //Queries on queue entity
            _resultList.Add("Queues with unapproved email routers");
            _resultList.Add("Queues with disabled mailboxes");
            _resultList.Add("Queues with emails in pending send status");
            _resultList.Add("Queues with emails in failed send status");
            _resultList.Add("Queues with emails in pending send status > 15 mins");
    
            //Queries on user entity
            _resultList.Add("Users with unapproved email routers");
            _resultList.Add("Users with disabled mailboxes");
            _resultList.Add("Users with emails in pending send status");
            _resultList.Add("Users with emails in failed send status");
            _resultList.Add("Users with emails in pending send status > 15 mins");
    
        }


        /// <summary>
        /// Button to login to CRM and create a CrmService Client 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            #region Login Control
            // Wire Event to login response. 
            Ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
            // Show the dialog. 
            Ctrl.ShowDialog();

            // Handel return. 
            if (Ctrl.CrmConnectionMgr != null && Ctrl.CrmConnectionMgr.CrmSvc != null && Ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            {
                OrganizationLbl.Content = Ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgFriendlyName;
                MessageBox.Show("Good Connect");
                loginBtn.Visibility = Visibility.Hidden;
                queryBtn.Visibility = Visibility.Visible;
                queryComboBox.Visibility = Visibility.Visible;
                OrganizationLbl.Visibility = Visibility.Visible;
            }
            else
                MessageBox.Show("BadConnect");

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

        private void queryBtn_Click(object sender, RoutedEventArgs e)
        {
            ExecuteQueueQuery();               
        }


     

        /// <summary>
        /// Execute query based on selected query in combo box
        /// </summary>
        private void ExecuteQueueQuery()
        {
            CaseSwitch = queryComboBox.SelectedIndex;


            #region CRMServiceClient
            if (Ctrl.CrmConnectionMgr != null && Ctrl.CrmConnectionMgr.CrmSvc != null && Ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            {
                CrmServiceClient svcClient = Ctrl.CrmConnectionMgr.CrmSvc;

                if (svcClient.IsReady)
                {

                    using (OrganizationServiceContext sourceContext = new OrganizationServiceContext(svcClient.OrganizationServiceProxy))
                    {

                        var queueAttr = "name";
                        var userAttr = "fullname";

                        var queueEntityName = "queue";
                        var userEntityName = "systemuser";

                        var userIdName = "systemuserid";
                        var queueIdName = "queueid";


                        switch (CaseSwitch)
                        {
                            case 0:

                                //Query queues with unapproved email routers
                                QueryResult = (from q in
                                                   sourceContext.CreateQuery(queueEntityName)
                                      where q.GetAttributeValue<OptionSetValue>("emailrouteraccessapproval").Value == 2 ||
                                      q.GetAttributeValue<OptionSetValue>("emailrouteraccessapproval").Value == 3
                                      select q).ToList();

                                QueryResult = QueryResult.OrderBy(x => x.Attributes[queueAttr]).ToList();


                                NavigateToQueryResult(QueryResult, queueEntityName, queueAttr, queueIdName);


                                break;

                            case 1:

                                //Query queues with unenabled mailboxes
                                QueryResult = (from q in sourceContext.CreateQuery(queueEntityName)
                                               join m in sourceContext.CreateQuery("mailbox") on
                                               q.GetAttributeValue<EntityReference>("defaultmailbox").Id equals m.GetAttributeValue<Guid>("mailboxid")
                                               where m.GetAttributeValue<OptionSetValue>("statecode").Value == 1
                                               select q).ToList();

                                QueryResult = QueryResult.OrderBy(x => x.Attributes[queueAttr]).ToList();

                                NavigateToQueryResult(QueryResult, queueEntityName, queueAttr, queueIdName);


                                break;
                            


                            case 2:
                                ///Query queues with emails in failed send status
                                QueryResult = (from q in sourceContext.CreateQuery(queueEntityName)
                                      join em in sourceContext.CreateQuery("email") on
                                      q.GetAttributeValue<EntityReference>("queueid").Id equals em.GetAttributeValue<Guid>("emailsender")
                                      where em.GetAttributeValue<OptionSetValue>("statuscode").Value == 8
                                      select q).ToList();
                                QueryResult = QueryResult.OrderBy(x => x.Attributes[queueAttr]).ToList();

                                NavigateToQueryResult(QueryResult, queueEntityName, queueAttr, queueIdName);


                                break;

                            case 3:
                                ///Query queues with emails in pending send status > 15 mins
                                QueryResult = (from q in sourceContext.CreateQuery(queueEntityName)
                                               join em in sourceContext.CreateQuery("email") on
                                               q.GetAttributeValue<EntityReference>("queueid").Id equals em.GetAttributeValue<Guid>("emailsender")
                                               where em.GetAttributeValue<OptionSetValue>("statuscode").Value == 6 && em.GetAttributeValue<OptionSetValue>("actualdurationminutes").Value == 15
                                               select q).ToList();
                                QueryResult = QueryResult.OrderBy(x => x.Attributes[queueAttr]).ToList();

                                NavigateToQueryResult(QueryResult, queueEntityName, queueAttr, queueIdName);


                                break;

                            case 4:
                                ///Query queues with emails in pending send status
                                QueryResult = (from q in sourceContext.CreateQuery(queueEntityName)
                                               join em in sourceContext.CreateQuery("email") on
                                               q.GetAttributeValue<EntityReference>("queueid").Id equals em.GetAttributeValue<Guid>("emailsender")
                                               where em.GetAttributeValue<OptionSetValue>("statuscode").Value == 6
                                               select q).ToList();
                                QueryResult = QueryResult.OrderBy(x => x.Attributes[queueAttr]).ToList();

                                NavigateToQueryResult(QueryResult, queueEntityName, queueAttr, queueIdName);


                                break;




                            case 5:

                                //Query users with unapproved email routers
                                QueryResult = (from q in
                                                   sourceContext.CreateQuery(userEntityName)
                                               where q.GetAttributeValue<OptionSetValue>("emailrouteraccessapproval").Value == 2 ||
                                               q.GetAttributeValue<OptionSetValue>("emailrouteraccessapproval").Value == 3
                                               select q).ToList();

                                QueryResult = QueryResult.OrderBy(x => x.Attributes[userAttr]).ToList();


                                NavigateToQueryResult(QueryResult, userEntityName, userAttr, userIdName);


                                break;


                            case 6:

                                //Query users with unenabled mailboxes
                                QueryResult = (from q in sourceContext.CreateQuery(userEntityName)
                                               join m in sourceContext.CreateQuery("mailbox") on
                                               q.GetAttributeValue<EntityReference>("defaultmailbox").Id equals m.GetAttributeValue<Guid>("mailboxid")
                                               where m.GetAttributeValue<OptionSetValue>("statecode").Value == 1
                                               select q).ToList();

                                QueryResult = QueryResult.OrderBy(x => x.Attributes[userAttr]).ToList();

                                NavigateToQueryResult(QueryResult, userEntityName, userAttr, userIdName);


                                break;


                            case 7:
                                ///Query users with emails in failed send status
                                QueryResult = (from q in sourceContext.CreateQuery(userEntityName)
                                               join em in sourceContext.CreateQuery("email") on
                                               q.GetAttributeValue<EntityReference>("queueid").Id equals em.GetAttributeValue<Guid>("emailsender")
                                               where em.GetAttributeValue<OptionSetValue>("statuscode").Value == 8
                                               select q).ToList();
                                QueryResult = QueryResult.OrderBy(x => x.Attributes[userAttr]).ToList();

                                NavigateToQueryResult(QueryResult, userEntityName, userAttr, userIdName);


                                break;

                            case 8:
                                ///Query users with emails in pending send status > 15 mins
                                QueryResult = (from q in sourceContext.CreateQuery(userEntityName)
                                               join em in sourceContext.CreateQuery("email") on
                                               q.GetAttributeValue<EntityReference>("queueid").Id equals em.GetAttributeValue<Guid>("emailsender")
                                               where em.GetAttributeValue<OptionSetValue>("statuscode").Value == 6 && em.GetAttributeValue<OptionSetValue>("actualdurationminutes").Value == 15
                                               select q).ToList();
                                QueryResult = QueryResult.OrderBy(x => x.Attributes[userAttr]).ToList();

                                NavigateToQueryResult(QueryResult, userEntityName, userAttr, userIdName);


                                break;

                            case 9:
                                ///Query users with emails in pending send status
                                QueryResult = (from q in sourceContext.CreateQuery(userEntityName)
                                               join em in sourceContext.CreateQuery("email") on
                                               q.GetAttributeValue<EntityReference>("queueid").Id equals em.GetAttributeValue<Guid>("emailsender")
                                               where em.GetAttributeValue<OptionSetValue>("statuscode").Value == 6
                                               select q).ToList();
                                QueryResult = QueryResult.OrderBy(x => x.Attributes[userAttr]).ToList();

                                NavigateToQueryResult(QueryResult, userEntityName, userAttr, userIdName);


                                break;


                            default:
                                Console.WriteLine("Default case");
                                break;
                        }

                    }

                }
            }
            #endregion
        }

        /// <summary>
        /// Navigation call to query result view
        /// </summary>
        /// <param name="queryResult"></param>
        private void NavigateToQueryResult(List<Entity> queryResult, string entityName, string attrName, string idName)
        {
            this.Cursor = Cursors.Wait;
            QueryResultWindow queryWindow = new QueryResultWindow(queryResult, entityName, attrName, idName, Ctrl);
            queryWindow.ShowDialog();
            this.Cursor = Cursors.Arrow;
        }
   

    }
}
