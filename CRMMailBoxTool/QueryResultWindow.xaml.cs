using CRMMailBoxTool.LoginWindow;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;
using System.Diagnostics;
using System.Windows.Navigation;

namespace CRMMailBoxTool
{
    /// <summary>
    /// Interaction logic for QueryResultWindow.xaml
    /// </summary>
    public partial class QueryResultWindow : Window
    {
        private Entity PassedDownEntity = new Entity();
        private List<Entity> _queryResult;
        private List<string> QueryStringList = new List<string>();
        private string ResultCount = "Result Count: ";
        private CrmLogin _ctrl = null ;
        private CrmServiceClient SvcClient;
        private OrganizationService SourceService = null;
        private string _attrName = string.Empty;
        private string _entityName = string.Empty;
        private string _idName = string.Empty;


        private enum EmailStates 
        {
            Empty = 0,
            Approved = 1,
            PendingApproval = 2,
            Rejected = 3
        };


        private enum MailBoxStates
        {
            Enabled = 0,
            Disabled = 1
        };




        public QueryResultWindow(List<Entity> queryResult, string entityName, string attrName, string idName, CrmLogin ctrl)
        {
            InitializeComponent();

            this.DataContext = this;
            this._queryResult = queryResult;
            this._entityName = entityName;
            this._attrName = attrName;
            this._idName = idName;
            this._ctrl = ctrl;

            InitViewBindings();

            SvcClient = ctrl.CrmConnectionMgr.CrmSvc;
           
        }

        /// <summary>
        /// Initialize the view controls data bindings
        /// </summary>
        private void InitViewBindings()
        {
            foreach (Entity result in _queryResult)
            {
                QueryStringList.Add(result.Attributes[_attrName].ToString());
            }

            //Assign list box values
            queryResultListBox.ItemsSource = QueryStringList;
            resultCountLbl.Content = ResultCount + QueryStringList.Count;
        }


        /// <summary>
        /// Approve the selected queue's email
        /// </summary>
        private void ApproveEmail()
        {

            int selectedIndexEmail = queryResultListBox.SelectedIndex;
            Entity selectedEmail = _queryResult[selectedIndexEmail];
            Guid selectedGuid = (Guid)selectedEmail.Attributes[_idName];

            string attrToChange = "emailrouteraccessapproval";
            ColumnSet attributes = new ColumnSet(new string[] { attrToChange });
            int attrModState = (int)EmailStates.Approved;

            ChangeEmailState(selectedGuid, _entityName, attrToChange, attributes, attrModState);
        }



        private void ChangeEmailState(Guid selectedGuid, string entityName, string attrToChange, ColumnSet attributes, int attrModState)
        {
            try
            {
                if (SvcClient.IsReady)
                {

                    using (SourceService = new OrganizationService(SvcClient.OrganizationServiceProxy))
                    {
                        PassedDownEntity = SourceService.Retrieve(entityName, selectedGuid, attributes);

                        if (PassedDownEntity != null)
                        {
                            PassedDownEntity.Attributes[attrToChange] = new OptionSetValue(attrModState);
                            PassedDownEntity.EntityState = EntityState.Changed; //need to mark Entity as changed before update 
                            SourceService.Update(PassedDownEntity);

                            MessageBox.Show("Email Approved");

                            //Refresh the queue list after update
                            InitViewBindings();
                        }

                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("State change failed");
                //ErrorLogger.Log(ex);
            }
        }


        /// <summary>
        /// Enable the selected queue's mailbox
        /// </summary>
        private void EnableMailbox()
        {
            int selectedIndexEmail = queryResultListBox.SelectedIndex;
            Entity selectedEmail = _queryResult[selectedIndexEmail];
            Guid selectedMailboxGuid = selectedEmail.GetAttributeValue<EntityReference>("defaultmailbox").Id;


            string attrToChange = "statecode";
            string entityToChange = "mailbox";

            ColumnSet attributes = new ColumnSet(new string[] { attrToChange });

            int attrModState = (int)MailBoxStates.Enabled;

            ChangeMailboxState(selectedMailboxGuid, entityToChange, attrToChange, attributes, attrModState);

        }




        private void ChangeMailboxState(Guid selectedGuid, string entityToChange, string attrToChange, ColumnSet attributes, int attrModState)
        {

            try
            {
                if (SvcClient.IsReady)
                {

                    using (SourceService = new OrganizationService(SvcClient.OrganizationServiceProxy))
                    {
                        PassedDownEntity = SourceService.Retrieve(entityToChange, selectedGuid, attributes);

                        if (PassedDownEntity != null)
                        {
                            PassedDownEntity.Attributes[attrToChange] = new OptionSetValue(attrModState);
                            PassedDownEntity.EntityState = EntityState.Changed; //need to mark Entity as changed before update 
                            SourceService.Update(PassedDownEntity);

                            MessageBox.Show("Mailbox Enabled");
                            //Refresh the queue list after update
                            InitViewBindings();
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("State change failed");
                //ErrorLogger.Log(ex);
            }
        }

       


        private void approveEmailBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            ApproveEmail();
            this.Cursor = Cursors.Arrow;
        }

        private void enableMailboxBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            EnableMailbox();
            this.Cursor = Cursors.Arrow;
        }

     

        private void HelpHyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

              
    }
}
