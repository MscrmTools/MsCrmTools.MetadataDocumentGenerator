using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using MsCrmTools.MetadataDocumentGenerator.Helper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.ServiceModel;
using System.Windows.Forms;

namespace MsCrmTools.MetadataDocumentGenerator.Forms
{
    public partial class SolutionPicker : Form
    {
        private readonly IOrganizationService innerService;

        /// <summary>
        /// Variable containing the solutions found in the environment.
        /// </summary>
        private List<Entity> environmentSolutions = new List<Entity>();
        public SolutionPicker(IOrganizationService service)
        {
            InitializeComponent();

            innerService = service;
        }

        public List<Entity> SelectedSolutions { get; set; } = new List<Entity>();

        /// <summary>
        /// Method that filters solutions that contain part of the text entered in the textbox. The search is made when there are more than three characters entered in the textbox.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        private void filtersolution_TextChanged(object sender, EventArgs e)
        {
            var tempList = new List<Entity>();

            var text = filtersolution.Text.ToLower().Trim();
            
            tempList = (text.Length < 3) ? environmentSolutions :
                environmentSolutions.Where(c => c["friendlyname"].ToString().ToLower().Trim().Contains(text)).ToList();

            lstSolutions.Items.Clear();

            foreach (Entity solution in tempList)
            {
                ListViewItem item = new ListViewItem(solution["friendlyname"].ToString());
                item.SubItems.Add(solution["version"].ToString());
                item.SubItems.Add(((EntityReference)solution["publisherid"]).Name);
                item.Tag = solution;

                lstSolutions.Items.Add(item);
            }

            lbltotalsolutions.Text = $"{tempList.Count} - Solutions";
        }

        private void btnSolutionPickerCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnSolutionPickerValidate_Click(object sender, EventArgs e)
        {
            if (lstSolutions.SelectedItems.Count > 0)
            {
                SelectedSolutions.AddRange(lstSolutions.SelectedItems.Cast<ListViewItem>().Select(i => (Entity)i.Tag));
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show(this, @"Please select a solution!", @"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void lstSolutions_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            var list = (ListView)sender;
            list.Sorting = list.Sorting == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            list.ListViewItemSorter = new ListViewItemComparer(e.Column, list.Sorting);
        }

        private void lstSolutions_DoubleClick(object sender, EventArgs e)
        {
            btnSolutionPickerValidate_Click(null, null);
        }

        private EntityCollection RetrieveSolutions()
        {
            try
            {
                QueryExpression qe = new QueryExpression("solution");
                qe.Distinct = true;
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria = new FilterExpression();
                qe.Criteria.AddCondition(new ConditionExpression("isvisible", ConditionOperator.Equal, true));
                qe.Criteria.AddCondition(new ConditionExpression("uniquename", ConditionOperator.NotEqual, "Default"));

                return innerService.RetrieveMultiple(qe);
            }
            catch (Exception error)
            {
                if (error.InnerException is FaultException)
                {
                    throw new Exception("Error while retrieving solutions: " + error.InnerException.Message);
                }

                throw new Exception("Error while retrieving solutions: " + error.Message);
            }
        }

        private void SolutionPicker_Load(object sender, EventArgs e)
        {
            lstSolutions.Items.Clear();

            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync();
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = RetrieveSolutions();
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //Initializes the environmentSolutions property with the list of solutions found in the environment. It will be used to filter the solutions.
            environmentSolutions = ((EntityCollection)e.Result).Entities.ToList();

            foreach (Entity solution in ((EntityCollection)e.Result).Entities)
            {
                ListViewItem item = new ListViewItem(solution["friendlyname"].ToString());
                item.SubItems.Add(solution["version"].ToString());
                item.SubItems.Add(((EntityReference)solution["publisherid"]).Name);
                item.Tag = solution;

                lstSolutions.Items.Add(item);
            }
            
            //Displays the total number of solutions found.
            lbltotalsolutions.Text = $"{environmentSolutions.Count} - Solutions";
            
            lstSolutions.Enabled = true;
            btnSolutionPickerValidate.Enabled = true;
        }
    }
}