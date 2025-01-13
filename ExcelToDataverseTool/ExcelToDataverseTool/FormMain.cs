using Newtonsoft.Json;

namespace ExcelToDataverseTool
{
    public partial class FormMain : Form
    {
        private ImportTaskConfig g_config;
        private ImportItemConfig g_currentItemConfig;
        private bool g_editing = false;

        private TaskManager taskManager;

        public FormMain()
        {
            InitializeComponent();
            InitTaskConfig();
            ToggleEnabledState(g_editing);
            comboBoxDataverseColumnType.DataSource = DataverseColumn.GetColumnTypeList();
            comboBoxDataverseColumnType.SelectedIndex = 0;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            taskManager = new TaskManager();
            taskManager.Log = LogToTextBox;
            taskManager.ChangeControlsEnabledStatus = ChangeTaskControlsEnabledStatus;

        }

        public void ChangeTaskControlsEnabledStatus(bool enabled)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(ChangeTaskControlsEnabledStatus), enabled);
            }
            else
            {
                buttonRunAll.Enabled = enabled;
                buttonRunTask.Enabled = enabled;
                buttonSave.Enabled = enabled;

                buttonNewTask.Enabled = enabled;
                buttonSaveTask.Enabled = enabled;
                buttonDeleteTask.Enabled = enabled;

                buttonBrowseLocalFile.Enabled = enabled;
                buttonAddMapping.Enabled = enabled;
            }
        }



        private void LogToTextBox(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(LogToTextBox), message);
            }
            else
            {
                richTextBoxLogs.AppendText(message + Environment.NewLine);
                richTextBoxLogs.ScrollToCaret();
            }
        }

        void InitTaskConfig()
        {
            string profile = EnvTool.GetDefaultProfilePath();
            if (!string.IsNullOrEmpty(profile) && File.Exists(profile))
            {
                string jsonconent = File.ReadAllText(profile);
                g_config = JsonConvert.DeserializeObject<ImportTaskConfig>(jsonconent) ?? new ImportTaskConfig().GenerateSampleTaskConfig();
            }
            else
            {
                g_config = new ImportTaskConfig().GenerateSampleTaskConfig();
            }

            ImportTaskConfigToControls();
        }

        void ImportTaskConfigToControls()
        {
            if (g_config != null)
            {
                textBoxEnvUrl.Text = g_config.environmentUrl;
                textBoxTenantId.Text = g_config.tenantId;
                textBoxClientId.Text = g_config.clientId;
                textBoxClientSecret.Text = g_config.clientSecret;

                treeViewTasks.Nodes.Clear();

                foreach (var itemConfig in g_config.importItemConfigs)
                {
                    // First level node
                    TreeNode entityNode = new TreeNode(itemConfig.targetEntityName);
                    entityNode.Tag = itemConfig;

                    // Second level nodes
                    entityNode.Nodes.Add(new TreeNode($"Source File: {itemConfig.sourceFilePath}"));
                    entityNode.Nodes.Add(new TreeNode($"Target Entity Name: {itemConfig.targetEntityName}"));
                    entityNode.Nodes.Add(new TreeNode($"Primary Key: {itemConfig.primaryKey.Key} => {itemConfig.primaryKey.Value.Name}"));

                    // Column Mappings node
                    TreeNode columnMappingsNode = new TreeNode("Column Mappings");
                    foreach (var columnMapping in itemConfig.columnMappings)
                    {
                        columnMappingsNode.Nodes.Add(new TreeNode($"{columnMapping.Key} => {columnMapping.Value.Name}"));
                    }
                    entityNode.Nodes.Add(columnMappingsNode);

                    // Add the entity node to the TreeView
                    treeViewTasks.Nodes.Add(entityNode);
                    treeViewTasks.ExpandAll();
                }
            }
        }

        private void ToggleEnabledState(bool enabled)
        {
            textBoxEnvUrl.Enabled = enabled;
            textBoxTenantId.Enabled = enabled;
            textBoxClientId.Enabled = enabled;
            textBoxClientSecret.Enabled = enabled;

            if (enabled)
            {
                buttonSave.Text = "Save";
                buttonRunAll.Visible = false;
            }
            else
            {
                buttonSave.Text = "Edit";
                buttonRunAll.Visible = true;
            }

        }

        private void buttonSave_Click(object sender, EventArgs e)
        {   
          
            if (g_editing == true)
            {              
                g_config.environmentUrl = textBoxEnvUrl.Text;
                g_config.tenantId = textBoxTenantId.Text;
                g_config.clientId = textBoxClientId.Text;
                g_config.clientSecret = textBoxClientSecret.Text;

                g_config.BuildConnectionString();
                taskManager.SetTaskConfig(g_config);
                if (taskManager.CheckConnection() == false)
                {
                    MessageBox.Show("Connection failed, please check the connection settings.");
                    return;
                }
                string profile = EnvTool.GetDefaultProfilePath();
                g_config.SaveToFile(profile);
            }

            g_editing = !g_editing;
            ToggleEnabledState(g_editing);

        }


        private void treeViewTasks_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode selectedNode = e.Node;

            if (e.Node == null)
            {
                return;
            }

            while (selectedNode.Parent != null)
            {
                selectedNode = selectedNode.Parent;
            }

            if (selectedNode.Tag != null)
            {
                ImportItemConfig itemConfig = (ImportItemConfig)selectedNode.Tag;
                g_currentItemConfig = itemConfig;
                LoadItemConfigToControls(itemConfig);
                //switch to tabpageConfig
                tabControlTask.SelectedTab = tabPageConfig;
            }

        }

        private void LoadItemConfigToControls(ImportItemConfig itemConfig)
        {
            if (itemConfig != null)
            {
                textBoxExcelFile.Text = itemConfig.sourceFilePath;
                textBoxEntityName.Text = itemConfig.targetEntityName;

                //add primary key and all column mappings to the datagridview,the primary's columnIsPrimaryKey is true,other columns are false
                dataGridViewColumnMapping.Rows.Clear();
                if (itemConfig.primaryKey.Key != null)
                    dataGridViewColumnMapping.Rows.Add(itemConfig.primaryKey.Key, itemConfig.primaryKey.Value.Name, DataverseColumn.GetColumnTypeString(itemConfig.primaryKey.Value.Type), true);
                if (itemConfig.columnMappings == null)
                {
                    itemConfig.columnMappings = new Dictionary<string, DataverseColumn>();
                }
                foreach (var item in itemConfig.columnMappings)
                {
                    if (item.Key == itemConfig.primaryKey.Key)
                    {
                        continue;
                    }
                    dataGridViewColumnMapping.Rows.Add(item.Key, item.Value.Name, DataverseColumn.GetColumnTypeString(item.Value.Type), false);
                }

                //set comboboxactiontype to the itemConfig's actionType
                comboBoxActionType.Items.Clear();
                comboBoxActionType.Items.AddRange(ActionType.actions);
                comboBoxActionType.SelectedIndex = (int)itemConfig.actionType;

                textBoxExcelColumnName.Text = "";
                textBoxDataverseColumnName.Text = "";
                comboBoxDataverseColumnType.SelectedIndex = 0;
                checkBoxPrimaryKey.Checked = false;

                //remove auto complete source for entity name and dataverse column name
                textBoxEntityName.AutoCompleteCustomSource = null;
                textBoxEntityName.AutoCompleteMode = AutoCompleteMode.None;
                textBoxEntityName.AutoCompleteSource = AutoCompleteSource.None;

                textBoxDataverseColumnName.AutoCompleteCustomSource = null;
                textBoxDataverseColumnName.AutoCompleteMode = AutoCompleteMode.None;
                textBoxDataverseColumnName.AutoCompleteSource = AutoCompleteSource.None;

                textBoxExcelColumnName.AutoCompleteCustomSource = null;
                textBoxExcelColumnName.AutoCompleteMode = AutoCompleteMode.None;
                textBoxExcelColumnName.AutoCompleteSource = AutoCompleteSource.None;


            }
        }


        private void dataGridViewColumnMapping_KeyUp(object sender, KeyEventArgs e)
        {
            if (dataGridViewColumnMapping.Rows.Count == 0 || g_currentItemConfig.columnMappings.Count != dataGridViewColumnMapping.Rows.Count || e.KeyCode != Keys.Delete)
            {
                return;
            }
            try
            {
                //get the current selected row's index
                int rowIndex = dataGridViewColumnMapping.CurrentCell.RowIndex;
                string key = dataGridViewColumnMapping.Rows[rowIndex].Cells[0].Value.ToString();
                bool isPrimary = dataGridViewColumnMapping.Rows[rowIndex].Cells[3].Value.ToString() == "True" ? true : false;

                if (rowIndex == 0 && isPrimary == true)
                {
                    //prompt to confirm if the primary key is removed
                    if (MessageBox.Show("Primary key will be removed, do you want to continue?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        //add the primary key back
                        dataGridViewColumnMapping.Rows.Add(g_currentItemConfig.primaryKey.Key, g_currentItemConfig.primaryKey.Value.Name, DataverseColumn.GetColumnTypeString(g_currentItemConfig.primaryKey.Value.Type), true);
                    }
                    else
                    {
                        //check if primary is existed, if yes, remove it
                        if (g_currentItemConfig.primaryKey.Key != null)
                        {
                            if (g_currentItemConfig.columnMappings.ContainsKey(g_currentItemConfig.primaryKey.Key))
                            {
                                g_currentItemConfig.columnMappings.Remove(g_currentItemConfig.primaryKey.Key);
                            }
                        }
                        g_currentItemConfig.primaryKey = new KeyValuePair<string, DataverseColumn>(null, null);
                    }
                }
                else if (rowIndex < g_currentItemConfig.columnMappings.Count)
                {
                    g_currentItemConfig.columnMappings.Remove(key);
                }
            }
            catch (Exception ex)
            {
                LLog.LogException(ex);
            }

            //reload the datagridview
            LoadItemConfigToControls(g_currentItemConfig);
        }

        private void buttonNewTask_Click(object sender, EventArgs e)
        {
            g_currentItemConfig = new ImportItemConfig();
            LoadItemConfigToControls(g_currentItemConfig);
        }

        private void buttonAddMapping_Click(object sender, EventArgs e)
        {
            if (g_currentItemConfig == null)
            {
                MessageBox.Show("Please create a new task or select a task first.");
                return;
            }

            //check if all the required fields are filled
            if (string.IsNullOrEmpty(textBoxExcelFile.Text) || string.IsNullOrEmpty(textBoxEntityName.Text) || string.IsNullOrEmpty(comboBoxDataverseColumnType.Text))
            {
                MessageBox.Show("Please fill in all the required fields for column mapping.");
                return;
            }

            //add the new column mapping to g_currentItemConfig,then reload the datagridview
            string key = textBoxExcelColumnName.Text;
            string value = textBoxDataverseColumnName.Text;
            DataverseColumnType type = DataverseColumn.GetTypeFromString(comboBoxDataverseColumnType.Text);

            if(string.IsNullOrEmpty(key) || string.IsNullOrEmpty(value))
            {
                MessageBox.Show("Please fill in all the required fields for column mapping.");
                return;
            }


            if (g_currentItemConfig.columnMappings == null)
            {
                g_currentItemConfig.columnMappings = new Dictionary<string, DataverseColumn>();
            }

            if (g_currentItemConfig.columnMappings.ContainsKey(key))
            {
                if (checkBoxPrimaryKey.Checked)
                {
                    //check if there are primary key already existed, if yes, prompt to confirm
                    if (g_currentItemConfig.primaryKey.Key != null)
                    {
                        if (MessageBox.Show("Primary key already exists, do you want to update it?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.No)
                        {
                            return;
                        }                        
                    }
                    //update the primary key
                    g_currentItemConfig.primaryKey = new KeyValuePair<string, DataverseColumn>(key, new DataverseColumn() { Name = value, Type = type });
                    g_currentItemConfig.columnMappings[key] = new DataverseColumn() { Name = value, Type = type };
                }

                //update the other columns
                g_currentItemConfig.columnMappings[key] = new DataverseColumn() { Name = value, Type = type };
                if (g_currentItemConfig.primaryKey.Key == key)
                {
                    g_currentItemConfig.primaryKey = new KeyValuePair<string, DataverseColumn>(key, new DataverseColumn() { Name = value, Type = type });
                }
            }
            else
            {
                //if checkbox is checked,update the primary key
                if (checkBoxPrimaryKey.Checked)
                {
                    //check if there are primary key already existed, if yes, prompt to confirm
                    if (g_currentItemConfig.primaryKey.Key != null)
                    {
                        if (MessageBox.Show("Primary key already exists, do you want to update it?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.No)
                        {
                            return;
                        }
                        else
                        {
                            //remove the current primary key
                            g_currentItemConfig.columnMappings.Remove(g_currentItemConfig.primaryKey.Key);
                        }
                    }

                    g_currentItemConfig.primaryKey = new KeyValuePair<string, DataverseColumn>(key, new DataverseColumn() { Name = value, Type = type });
                    g_currentItemConfig.columnMappings.Add(key, new DataverseColumn() { Name = value, Type = type });
                }
                else
                {
                    g_currentItemConfig.columnMappings.Add(key, new DataverseColumn() { Name = value, Type = type });
                }


            }

            //load the new column mappings to the datagridview
            LoadItemConfigToControls(g_currentItemConfig);
        }

        private void buttonSaveTask_Click(object sender, EventArgs e)
        {

            //check if the file is existed, if not,prompt and exit
            if (!File.Exists(textBoxExcelFile.Text))
            {
                MessageBox.Show("The source file is not existed, please check the file path.");
                return;
            }

            //check if primary key is set, if not,prompt and exit
            if (g_currentItemConfig.primaryKey.Key == null)
            {
                MessageBox.Show("Please set the primary key for the task first.");
                return;
            }


            //save the current item config to the import task config
            if (g_currentItemConfig != null)
            {
                if (g_config.importItemConfigs == null)
                {
                    g_config.importItemConfigs = new List<ImportItemConfig>();
                }
                if (g_config.importItemConfigs.Contains(g_currentItemConfig))
                {
                    g_config.importItemConfigs.Remove(g_currentItemConfig);
                }
                g_config.importItemConfigs.Add(g_currentItemConfig);
                g_config.SaveToFile(EnvTool.GetDefaultProfilePath());
                ImportTaskConfigToControls();
            }
        }

        private void textBoxExcelFile_TextChanged(object sender, EventArgs e)
        {
            if (g_currentItemConfig != null)
            {
                g_currentItemConfig.sourceFilePath = textBoxExcelFile.Text;
            }
        }

        private void textBoxEntityName_TextChanged(object sender, EventArgs e)
        {
            if (g_currentItemConfig != null)
            {
                g_currentItemConfig.targetEntityName = textBoxEntityName.Text;
            }
        }

        private void buttonBrowseLocalFile_Click(object sender, EventArgs e)
        {
            //open file dialog to select the excel file and csv file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx|CSV Files|*.csv";
            openFileDialog.Title = "Select the source file";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxExcelFile.Text = openFileDialog.FileName;
            }

        }

        private void dataGridViewColumnMapping_SelectionChanged(object sender, EventArgs e)
        {
            ////load the selected column mapping to the textboxes
            //if (dataGridViewColumnMapping.SelectedRows.Count > 0)
            //{
            //    DataGridViewRow row = dataGridViewColumnMapping.SelectedRows[0];
            //    textBoxExcelColumnName.Text = row.Cells[0].Value.ToString();
            //    textBoxDataverseColumnName.Text = row.Cells[1].Value.ToString();
            //    comboBoxDataverseColumnType.Text = row.Cells[2].Value.ToString();
            //    checkBoxPrimaryKey.Checked = (bool)row.Cells[3].Value;
            //}
        }


        private void buttonDeleteTask_Click(object sender, EventArgs e)
        {
            //if the current item config is not null,remove it from the import task config
            if (g_currentItemConfig != null)
            {
                //prompt to confirm
                if (MessageBox.Show("Are you sure to delete the task?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    if (g_config.importItemConfigs.Contains(g_currentItemConfig))
                    {
                        g_config.importItemConfigs.Remove(g_currentItemConfig);
                        g_config.SaveToFile(EnvTool.GetDefaultProfilePath());
                    }
                    ImportTaskConfigToControls();
                }
            }
        }

        private void buttonRunTask_Click(object sender, EventArgs e)
        {
            //call taskmanager to run the task, change the button to disabled until the task is finished
            buttonRunTask.Enabled = false;
            buttonRunAll.Enabled = false;
            if (g_currentItemConfig != null)
            {
                //taskManager.ImportData(g_currentItemConfig);
                Task.Run(() => taskManager.ImportData(g_currentItemConfig));
                tabControlTask.SelectedTab = tabPageLogs;
            }
            else
            {
                buttonRunTask.Enabled = true;
                buttonRunAll.Enabled = true;
            }

        }


        private async void textBoxEntityName_Enter(object sender, EventArgs e)
        {
            if (textBoxEntityName.AutoCompleteCustomSource == null || textBoxEntityName.AutoCompleteCustomSource.Count == 0)
            {
                string[] entities = await Task.Run(() => taskManager.GetAvailableEntities());
                if (entities != null)
                {
                    AutoCompleteStringCollection autoCompleteStringCollection = new AutoCompleteStringCollection();
                    autoCompleteStringCollection.AddRange(entities);
                    textBoxEntityName.AutoCompleteCustomSource = autoCompleteStringCollection;
                    textBoxEntityName.AutoCompleteMode = AutoCompleteMode.Suggest;
                    textBoxEntityName.AutoCompleteSource = AutoCompleteSource.CustomSource;
                }
            }
        }

        private async void textBoxExcelColumnName_Enter(object sender, EventArgs e)
        {
            if ((textBoxExcelColumnName.AutoCompleteCustomSource == null || textBoxExcelColumnName.AutoCompleteCustomSource.Count == 0) && !string.IsNullOrEmpty(textBoxExcelFile.Text))
            {
                string[] columns = await Task.Run(() => taskManager.GetHeadersFromExcelandCsv(textBoxExcelFile.Text));
                if (columns != null)
                {
                    AutoCompleteStringCollection autoCompleteStringCollection = new AutoCompleteStringCollection();
                    autoCompleteStringCollection.AddRange(columns);
                    textBoxExcelColumnName.AutoCompleteCustomSource = autoCompleteStringCollection;
                    textBoxExcelColumnName.AutoCompleteMode = AutoCompleteMode.Suggest;
                    textBoxExcelColumnName.AutoCompleteSource = AutoCompleteSource.CustomSource;
                }
            }
        }

        private async void textBoxDataverseColumnName_Enter(object sender, EventArgs e)
        {
            if ((textBoxDataverseColumnName.AutoCompleteCustomSource == null || textBoxDataverseColumnName.AutoCompleteCustomSource.Count == 0) && !string.IsNullOrEmpty(textBoxEntityName.Text))
            {
                string[] columns = await Task.Run(() => taskManager.GetAvailableColumns(textBoxEntityName.Text));
                if (columns != null)
                {
                    AutoCompleteStringCollection autoCompleteStringCollection = new AutoCompleteStringCollection();
                    autoCompleteStringCollection.AddRange(columns);
                    textBoxDataverseColumnName.AutoCompleteCustomSource = autoCompleteStringCollection;
                    textBoxDataverseColumnName.AutoCompleteMode = AutoCompleteMode.Suggest;
                    textBoxDataverseColumnName.AutoCompleteSource = AutoCompleteSource.CustomSource;
                }
            }
        }

        private void buttonRunAll_Click(object sender, EventArgs e)
        {
            buttonRunTask.Enabled = false;
            buttonRunAll.Enabled = false;
            if (g_config != null)
            {
                Task.Run(() => taskManager.RunImportTasks(g_config));
                tabControlTask.SelectedTab = tabPageLogs;
            }
            else
            {
                buttonRunTask.Enabled = true;
                buttonRunAll.Enabled = true;
            }
        }

        private void dataGridViewColumnMapping_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridViewColumnMapping.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dataGridViewColumnMapping.SelectedRows[0];
                textBoxExcelColumnName.Text = row.Cells[0].Value.ToString();
                textBoxDataverseColumnName.Text = row.Cells[1].Value.ToString();
                comboBoxDataverseColumnType.Text = row.Cells[2].Value.ToString();
                checkBoxPrimaryKey.Checked = (bool)row.Cells[3].Value;
            }
        }

        private async void textBoxDataverseColumnName_TextChanged(object sender, EventArgs e)
        {
            //check if the text value is matching one of the available columns,if yes, call taskmanager to get the column type
            if (!string.IsNullOrEmpty(textBoxEntityName.Text) && !string.IsNullOrEmpty(textBoxDataverseColumnName.Text))
            {
                if (textBoxDataverseColumnName.AutoCompleteCustomSource != null && textBoxDataverseColumnName.AutoCompleteCustomSource.Contains(textBoxDataverseColumnName.Text))
                {
                    DataverseColumnType type = await Task.Run(() => taskManager.GetColumnDataTypeFromDataverse(textBoxEntityName.Text, textBoxDataverseColumnName.Text));
                    if (type != DataverseColumnType.SingleLineOfText)
                    {
                        comboBoxDataverseColumnType.Text = DataverseColumn.GetColumnTypeString(type);
                    }
                }
            }
        }

        private void comboBoxActionType_SelectedIndexChanged(object sender, EventArgs e)
        {
            //update the action type of the current item config
            if (g_currentItemConfig != null)
            {
                g_currentItemConfig.actionType = ActionType.GetActionType(comboBoxActionType.Text);
            }
        }
    }
}
