namespace ExcelToDataverseTool
{
    partial class FormMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            panelTop = new Panel();
            buttonRunAll = new Button();
            buttonSave = new Button();
            textBoxClientSecret = new TextBox();
            textBoxClientId = new TextBox();
            textBoxTenantId = new TextBox();
            textBoxEnvUrl = new TextBox();
            labelClientSecret = new Label();
            labelClientId = new Label();
            labelTenantId = new Label();
            labelenvurl = new Label();
            groupBoxTasks = new GroupBox();
            treeViewTasks = new TreeView();
            panelBody = new Panel();
            tabControlTask = new TabControl();
            tabPageLogs = new TabPage();
            richTextBoxLogs = new RichTextBox();
            tabPageConfig = new TabPage();
            dataGridViewColumnMapping = new DataGridView();
            ColumnExcelColumn = new DataGridViewTextBoxColumn();
            ColumnDataverseColumn = new DataGridViewTextBoxColumn();
            ColumnType = new DataGridViewTextBoxColumn();
            ColumnIsPrimary = new DataGridViewTextBoxColumn();
            panelConfigTop = new Panel();
            labelActionType = new Label();
            comboBoxActionType = new ComboBox();
            buttonDeleteTask = new Button();
            buttonSaveTask = new Button();
            checkBoxPrimaryKey = new CheckBox();
            labelColumnDataType = new Label();
            comboBoxDataverseColumnType = new ComboBox();
            textBoxDataverseColumnName = new TextBox();
            textBoxExcelColumnName = new TextBox();
            textBoxEntityName = new TextBox();
            buttonRunTask = new Button();
            buttonNewTask = new Button();
            buttonAddMapping = new Button();
            buttonBrowseLocalFile = new Button();
            textBoxExcelFile = new TextBox();
            labelDataverseColumn = new Label();
            labelExcelColumn = new Label();
            labelDataverseEntity = new Label();
            labelExcelFile = new Label();
            tabPageGuide = new TabPage();
            richTextBoxGuide = new RichTextBox();
            contextMenuStripTask = new ContextMenuStrip(components);
            openFileDialog1 = new OpenFileDialog();
            panelTop.SuspendLayout();
            groupBoxTasks.SuspendLayout();
            panelBody.SuspendLayout();
            tabControlTask.SuspendLayout();
            tabPageLogs.SuspendLayout();
            tabPageConfig.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewColumnMapping).BeginInit();
            panelConfigTop.SuspendLayout();
            tabPageGuide.SuspendLayout();
            SuspendLayout();
            // 
            // panelTop
            // 
            panelTop.Controls.Add(buttonRunAll);
            panelTop.Controls.Add(buttonSave);
            panelTop.Controls.Add(textBoxClientSecret);
            panelTop.Controls.Add(textBoxClientId);
            panelTop.Controls.Add(textBoxTenantId);
            panelTop.Controls.Add(textBoxEnvUrl);
            panelTop.Controls.Add(labelClientSecret);
            panelTop.Controls.Add(labelClientId);
            panelTop.Controls.Add(labelTenantId);
            panelTop.Controls.Add(labelenvurl);
            panelTop.Dock = DockStyle.Top;
            panelTop.Location = new Point(0, 0);
            panelTop.Name = "panelTop";
            panelTop.Size = new Size(1795, 150);
            panelTop.TabIndex = 0;
            // 
            // buttonRunAll
            // 
            buttonRunAll.Location = new Point(1633, 26);
            buttonRunAll.Name = "buttonRunAll";
            buttonRunAll.Size = new Size(137, 84);
            buttonRunAll.TabIndex = 9;
            buttonRunAll.Text = "Run All";
            buttonRunAll.UseVisualStyleBackColor = true;
            buttonRunAll.Click += buttonRunAll_Click;
            // 
            // buttonSave
            // 
            buttonSave.Location = new Point(1462, 26);
            buttonSave.Name = "buttonSave";
            buttonSave.Size = new Size(137, 84);
            buttonSave.TabIndex = 8;
            buttonSave.Text = "Save";
            buttonSave.UseVisualStyleBackColor = true;
            buttonSave.Click += buttonSave_Click;
            // 
            // textBoxClientSecret
            // 
            textBoxClientSecret.Location = new Point(918, 83);
            textBoxClientSecret.Name = "textBoxClientSecret";
            textBoxClientSecret.PasswordChar = '*';
            textBoxClientSecret.Size = new Size(516, 30);
            textBoxClientSecret.TabIndex = 7;
            // 
            // textBoxClientId
            // 
            textBoxClientId.Location = new Point(206, 83);
            textBoxClientId.Name = "textBoxClientId";
            textBoxClientId.Size = new Size(516, 30);
            textBoxClientId.TabIndex = 6;
            // 
            // textBoxTenantId
            // 
            textBoxTenantId.Location = new Point(918, 23);
            textBoxTenantId.Name = "textBoxTenantId";
            textBoxTenantId.Size = new Size(516, 30);
            textBoxTenantId.TabIndex = 5;
            // 
            // textBoxEnvUrl
            // 
            textBoxEnvUrl.Location = new Point(206, 23);
            textBoxEnvUrl.Name = "textBoxEnvUrl";
            textBoxEnvUrl.Size = new Size(516, 30);
            textBoxEnvUrl.TabIndex = 4;
            // 
            // labelClientSecret
            // 
            labelClientSecret.AutoSize = true;
            labelClientSecret.Location = new Point(781, 86);
            labelClientSecret.Name = "labelClientSecret";
            labelClientSecret.Size = new Size(118, 24);
            labelClientSecret.TabIndex = 3;
            labelClientSecret.Text = "Client Secret";
            // 
            // labelClientId
            // 
            labelClientId.AutoSize = true;
            labelClientId.Location = new Point(27, 86);
            labelClientId.Name = "labelClientId";
            labelClientId.Size = new Size(84, 24);
            labelClientId.TabIndex = 2;
            labelClientId.Text = "Client ID";
            // 
            // labelTenantId
            // 
            labelTenantId.AutoSize = true;
            labelTenantId.Location = new Point(781, 26);
            labelTenantId.Name = "labelTenantId";
            labelTenantId.Size = new Size(93, 24);
            labelTenantId.TabIndex = 1;
            labelTenantId.Text = "Tenant ID";
            // 
            // labelenvurl
            // 
            labelenvurl.AutoSize = true;
            labelenvurl.Location = new Point(27, 26);
            labelenvurl.Name = "labelenvurl";
            labelenvurl.Size = new Size(158, 24);
            labelenvurl.TabIndex = 0;
            labelenvurl.Text = "Environment URL";
            // 
            // groupBoxTasks
            // 
            groupBoxTasks.Controls.Add(treeViewTasks);
            groupBoxTasks.Dock = DockStyle.Left;
            groupBoxTasks.Location = new Point(0, 150);
            groupBoxTasks.Name = "groupBoxTasks";
            groupBoxTasks.Size = new Size(426, 955);
            groupBoxTasks.TabIndex = 1;
            groupBoxTasks.TabStop = false;
            groupBoxTasks.Text = "Tasks";
            // 
            // treeViewTasks
            // 
            treeViewTasks.Dock = DockStyle.Fill;
            treeViewTasks.Location = new Point(3, 26);
            treeViewTasks.Name = "treeViewTasks";
            treeViewTasks.Size = new Size(420, 926);
            treeViewTasks.TabIndex = 0;
            treeViewTasks.NodeMouseClick += treeViewTasks_NodeMouseClick;
            // 
            // panelBody
            // 
            panelBody.Controls.Add(tabControlTask);
            panelBody.Dock = DockStyle.Fill;
            panelBody.Location = new Point(426, 150);
            panelBody.Name = "panelBody";
            panelBody.Size = new Size(1369, 955);
            panelBody.TabIndex = 2;
            // 
            // tabControlTask
            // 
            tabControlTask.Controls.Add(tabPageLogs);
            tabControlTask.Controls.Add(tabPageConfig);
            tabControlTask.Controls.Add(tabPageGuide);
            tabControlTask.Dock = DockStyle.Fill;
            tabControlTask.Location = new Point(0, 0);
            tabControlTask.Name = "tabControlTask";
            tabControlTask.SelectedIndex = 0;
            tabControlTask.Size = new Size(1369, 955);
            tabControlTask.TabIndex = 0;
            // 
            // tabPageLogs
            // 
            tabPageLogs.Controls.Add(richTextBoxLogs);
            tabPageLogs.Location = new Point(4, 33);
            tabPageLogs.Name = "tabPageLogs";
            tabPageLogs.Padding = new Padding(3);
            tabPageLogs.Size = new Size(1361, 918);
            tabPageLogs.TabIndex = 0;
            tabPageLogs.Text = "Logs";
            tabPageLogs.UseVisualStyleBackColor = true;
            // 
            // richTextBoxLogs
            // 
            richTextBoxLogs.BackColor = SystemColors.ButtonHighlight;
            richTextBoxLogs.Dock = DockStyle.Fill;
            richTextBoxLogs.Location = new Point(3, 3);
            richTextBoxLogs.Name = "richTextBoxLogs";
            richTextBoxLogs.ReadOnly = true;
            richTextBoxLogs.Size = new Size(1355, 912);
            richTextBoxLogs.TabIndex = 0;
            richTextBoxLogs.Text = "";
            // 
            // tabPageConfig
            // 
            tabPageConfig.Controls.Add(dataGridViewColumnMapping);
            tabPageConfig.Controls.Add(panelConfigTop);
            tabPageConfig.Location = new Point(4, 33);
            tabPageConfig.Name = "tabPageConfig";
            tabPageConfig.Padding = new Padding(3);
            tabPageConfig.Size = new Size(1361, 918);
            tabPageConfig.TabIndex = 1;
            tabPageConfig.Text = "Config";
            tabPageConfig.UseVisualStyleBackColor = true;
            // 
            // dataGridViewColumnMapping
            // 
            dataGridViewColumnMapping.AllowUserToAddRows = false;
            dataGridViewColumnMapping.AllowUserToDeleteRows = false;
            dataGridViewColumnMapping.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewColumnMapping.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewColumnMapping.Columns.AddRange(new DataGridViewColumn[] { ColumnExcelColumn, ColumnDataverseColumn, ColumnType, ColumnIsPrimary });
            dataGridViewColumnMapping.Dock = DockStyle.Fill;
            dataGridViewColumnMapping.EditMode = DataGridViewEditMode.EditProgrammatically;
            dataGridViewColumnMapping.Location = new Point(3, 411);
            dataGridViewColumnMapping.MultiSelect = false;
            dataGridViewColumnMapping.Name = "dataGridViewColumnMapping";
            dataGridViewColumnMapping.ReadOnly = true;
            dataGridViewColumnMapping.RowHeadersVisible = false;
            dataGridViewColumnMapping.RowHeadersWidth = 62;
            dataGridViewColumnMapping.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewColumnMapping.Size = new Size(1355, 504);
            dataGridViewColumnMapping.TabIndex = 1;
            dataGridViewColumnMapping.SelectionChanged += dataGridViewColumnMapping_SelectionChanged;
            dataGridViewColumnMapping.KeyUp += dataGridViewColumnMapping_KeyUp;
            dataGridViewColumnMapping.MouseClick += dataGridViewColumnMapping_MouseClick;
            // 
            // ColumnExcelColumn
            // 
            ColumnExcelColumn.HeaderText = "Excel Column";
            ColumnExcelColumn.MinimumWidth = 8;
            ColumnExcelColumn.Name = "ColumnExcelColumn";
            ColumnExcelColumn.ReadOnly = true;
            // 
            // ColumnDataverseColumn
            // 
            ColumnDataverseColumn.HeaderText = "Dataverse Column";
            ColumnDataverseColumn.MinimumWidth = 8;
            ColumnDataverseColumn.Name = "ColumnDataverseColumn";
            ColumnDataverseColumn.ReadOnly = true;
            // 
            // ColumnType
            // 
            ColumnType.HeaderText = "Dataverse Column Type";
            ColumnType.MinimumWidth = 8;
            ColumnType.Name = "ColumnType";
            ColumnType.ReadOnly = true;
            // 
            // ColumnIsPrimary
            // 
            ColumnIsPrimary.FillWeight = 30F;
            ColumnIsPrimary.HeaderText = "Primary?";
            ColumnIsPrimary.MinimumWidth = 8;
            ColumnIsPrimary.Name = "ColumnIsPrimary";
            ColumnIsPrimary.ReadOnly = true;
            // 
            // panelConfigTop
            // 
            panelConfigTop.Controls.Add(labelActionType);
            panelConfigTop.Controls.Add(comboBoxActionType);
            panelConfigTop.Controls.Add(buttonDeleteTask);
            panelConfigTop.Controls.Add(buttonSaveTask);
            panelConfigTop.Controls.Add(checkBoxPrimaryKey);
            panelConfigTop.Controls.Add(labelColumnDataType);
            panelConfigTop.Controls.Add(comboBoxDataverseColumnType);
            panelConfigTop.Controls.Add(textBoxDataverseColumnName);
            panelConfigTop.Controls.Add(textBoxExcelColumnName);
            panelConfigTop.Controls.Add(textBoxEntityName);
            panelConfigTop.Controls.Add(buttonRunTask);
            panelConfigTop.Controls.Add(buttonNewTask);
            panelConfigTop.Controls.Add(buttonAddMapping);
            panelConfigTop.Controls.Add(buttonBrowseLocalFile);
            panelConfigTop.Controls.Add(textBoxExcelFile);
            panelConfigTop.Controls.Add(labelDataverseColumn);
            panelConfigTop.Controls.Add(labelExcelColumn);
            panelConfigTop.Controls.Add(labelDataverseEntity);
            panelConfigTop.Controls.Add(labelExcelFile);
            panelConfigTop.Dock = DockStyle.Top;
            panelConfigTop.Location = new Point(3, 3);
            panelConfigTop.Name = "panelConfigTop";
            panelConfigTop.Size = new Size(1355, 408);
            panelConfigTop.TabIndex = 0;
            // 
            // labelActionType
            // 
            labelActionType.AutoSize = true;
            labelActionType.Location = new Point(593, 115);
            labelActionType.Name = "labelActionType";
            labelActionType.Size = new Size(113, 24);
            labelActionType.TabIndex = 25;
            labelActionType.Text = "Action Type";
            // 
            // comboBoxActionType
            // 
            comboBoxActionType.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxActionType.FormattingEnabled = true;
            comboBoxActionType.Location = new Point(718, 111);
            comboBoxActionType.Name = "comboBoxActionType";
            comboBoxActionType.Size = new Size(283, 32);
            comboBoxActionType.TabIndex = 24;
            comboBoxActionType.SelectedIndexChanged += comboBoxActionType_SelectedIndexChanged;
            // 
            // buttonDeleteTask
            // 
            buttonDeleteTask.Location = new Point(385, 341);
            buttonDeleteTask.Name = "buttonDeleteTask";
            buttonDeleteTask.Size = new Size(168, 50);
            buttonDeleteTask.TabIndex = 23;
            buttonDeleteTask.Text = "Delete Task";
            buttonDeleteTask.UseVisualStyleBackColor = true;
            buttonDeleteTask.Click += buttonDeleteTask_Click;
            // 
            // buttonSaveTask
            // 
            buttonSaveTask.Location = new Point(202, 341);
            buttonSaveTask.Name = "buttonSaveTask";
            buttonSaveTask.Size = new Size(168, 50);
            buttonSaveTask.TabIndex = 22;
            buttonSaveTask.Text = "Save Task";
            buttonSaveTask.UseVisualStyleBackColor = true;
            buttonSaveTask.Click += buttonSaveTask_Click;
            // 
            // checkBoxPrimaryKey
            // 
            checkBoxPrimaryKey.AutoSize = true;
            checkBoxPrimaryKey.Location = new Point(598, 195);
            checkBoxPrimaryKey.Name = "checkBoxPrimaryKey";
            checkBoxPrimaryKey.Size = new Size(157, 28);
            checkBoxPrimaryKey.TabIndex = 21;
            checkBoxPrimaryKey.Text = "Is Primary Key";
            checkBoxPrimaryKey.UseVisualStyleBackColor = true;
            // 
            // labelColumnDataType
            // 
            labelColumnDataType.AutoSize = true;
            labelColumnDataType.Location = new Point(593, 262);
            labelColumnDataType.Name = "labelColumnDataType";
            labelColumnDataType.Size = new Size(98, 24);
            labelColumnDataType.TabIndex = 20;
            labelColumnDataType.Text = "Data Type";
            // 
            // comboBoxDataverseColumnType
            // 
            comboBoxDataverseColumnType.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxDataverseColumnType.FormattingEnabled = true;
            comboBoxDataverseColumnType.Location = new Point(718, 262);
            comboBoxDataverseColumnType.Name = "comboBoxDataverseColumnType";
            comboBoxDataverseColumnType.Size = new Size(283, 32);
            comboBoxDataverseColumnType.TabIndex = 19;
            // 
            // textBoxDataverseColumnName
            // 
            textBoxDataverseColumnName.Location = new Point(198, 262);
            textBoxDataverseColumnName.Name = "textBoxDataverseColumnName";
            textBoxDataverseColumnName.Size = new Size(283, 30);
            textBoxDataverseColumnName.TabIndex = 18;
            textBoxDataverseColumnName.TextChanged += textBoxDataverseColumnName_TextChanged;
            textBoxDataverseColumnName.Enter += textBoxDataverseColumnName_Enter;
            // 
            // textBoxExcelColumnName
            // 
            textBoxExcelColumnName.Location = new Point(198, 190);
            textBoxExcelColumnName.Name = "textBoxExcelColumnName";
            textBoxExcelColumnName.Size = new Size(280, 30);
            textBoxExcelColumnName.TabIndex = 17;
            textBoxExcelColumnName.Enter += textBoxExcelColumnName_Enter;
            // 
            // textBoxEntityName
            // 
            textBoxEntityName.Location = new Point(198, 112);
            textBoxEntityName.Name = "textBoxEntityName";
            textBoxEntityName.Size = new Size(283, 30);
            textBoxEntityName.TabIndex = 16;
            textBoxEntityName.TextChanged += textBoxEntityName_TextChanged;
            textBoxEntityName.Enter += textBoxEntityName_Enter;
            // 
            // buttonRunTask
            // 
            buttonRunTask.Location = new Point(574, 341);
            buttonRunTask.Name = "buttonRunTask";
            buttonRunTask.Size = new Size(168, 50);
            buttonRunTask.TabIndex = 15;
            buttonRunTask.Text = "Run Task";
            buttonRunTask.UseVisualStyleBackColor = true;
            buttonRunTask.Click += buttonRunTask_Click;
            // 
            // buttonNewTask
            // 
            buttonNewTask.Location = new Point(19, 341);
            buttonNewTask.Name = "buttonNewTask";
            buttonNewTask.Size = new Size(168, 50);
            buttonNewTask.TabIndex = 13;
            buttonNewTask.Text = "New Task";
            buttonNewTask.UseVisualStyleBackColor = true;
            buttonNewTask.Click += buttonNewTask_Click;
            // 
            // buttonAddMapping
            // 
            buttonAddMapping.Location = new Point(1057, 249);
            buttonAddMapping.Name = "buttonAddMapping";
            buttonAddMapping.Size = new Size(168, 50);
            buttonAddMapping.TabIndex = 10;
            buttonAddMapping.Text = "Add Mapping";
            buttonAddMapping.UseVisualStyleBackColor = true;
            buttonAddMapping.Click += buttonAddMapping_Click;
            // 
            // buttonBrowseLocalFile
            // 
            buttonBrowseLocalFile.Location = new Point(1001, 29);
            buttonBrowseLocalFile.Name = "buttonBrowseLocalFile";
            buttonBrowseLocalFile.Size = new Size(33, 34);
            buttonBrowseLocalFile.TabIndex = 8;
            buttonBrowseLocalFile.Text = "...";
            buttonBrowseLocalFile.UseVisualStyleBackColor = true;
            buttonBrowseLocalFile.Click += buttonBrowseLocalFile_Click;
            // 
            // textBoxExcelFile
            // 
            textBoxExcelFile.Location = new Point(198, 31);
            textBoxExcelFile.Name = "textBoxExcelFile";
            textBoxExcelFile.Size = new Size(803, 30);
            textBoxExcelFile.TabIndex = 4;
            textBoxExcelFile.TextChanged += textBoxExcelFile_TextChanged;
            // 
            // labelDataverseColumn
            // 
            labelDataverseColumn.AutoSize = true;
            labelDataverseColumn.Location = new Point(15, 265);
            labelDataverseColumn.Name = "labelDataverseColumn";
            labelDataverseColumn.Size = new Size(167, 24);
            labelDataverseColumn.TabIndex = 3;
            labelDataverseColumn.Text = "Dataverse Column";
            // 
            // labelExcelColumn
            // 
            labelExcelColumn.AutoSize = true;
            labelExcelColumn.Location = new Point(15, 193);
            labelExcelColumn.Name = "labelExcelColumn";
            labelExcelColumn.Size = new Size(125, 24);
            labelExcelColumn.TabIndex = 2;
            labelExcelColumn.Text = "Excel Column";
            // 
            // labelDataverseEntity
            // 
            labelDataverseEntity.AutoSize = true;
            labelDataverseEntity.Location = new Point(15, 115);
            labelDataverseEntity.Name = "labelDataverseEntity";
            labelDataverseEntity.Size = new Size(150, 24);
            labelDataverseEntity.TabIndex = 1;
            labelDataverseEntity.Text = "Dataverse Entity";
            // 
            // labelExcelFile
            // 
            labelExcelFile.AutoSize = true;
            labelExcelFile.Location = new Point(15, 34);
            labelExcelFile.Name = "labelExcelFile";
            labelExcelFile.Size = new Size(88, 24);
            labelExcelFile.TabIndex = 0;
            labelExcelFile.Text = "Excel File";
            // 
            // tabPageGuide
            // 
            tabPageGuide.Controls.Add(richTextBoxGuide);
            tabPageGuide.Location = new Point(4, 33);
            tabPageGuide.Name = "tabPageGuide";
            tabPageGuide.Padding = new Padding(3);
            tabPageGuide.Size = new Size(1361, 918);
            tabPageGuide.TabIndex = 2;
            tabPageGuide.Text = "Guide";
            tabPageGuide.UseVisualStyleBackColor = true;
            // 
            // richTextBoxGuide
            // 
            richTextBoxGuide.AcceptsTab = true;
            richTextBoxGuide.AutoWordSelection = true;
            richTextBoxGuide.BackColor = SystemColors.Info;
            richTextBoxGuide.Dock = DockStyle.Fill;
            richTextBoxGuide.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            richTextBoxGuide.Location = new Point(3, 3);
            richTextBoxGuide.Name = "richTextBoxGuide";
            richTextBoxGuide.ReadOnly = true;
            richTextBoxGuide.Size = new Size(1355, 912);
            richTextBoxGuide.TabIndex = 0;
            richTextBoxGuide.Text = resources.GetString("richTextBoxGuide.Text");
            // 
            // contextMenuStripTask
            // 
            contextMenuStripTask.ImageScalingSize = new Size(24, 24);
            contextMenuStripTask.Name = "contextMenuStripTask";
            contextMenuStripTask.Size = new Size(61, 4);
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialogExcelFile";
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(11F, 24F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1795, 1105);
            Controls.Add(panelBody);
            Controls.Add(groupBoxTasks);
            Controls.Add(panelTop);
            Name = "FormMain";
            Text = "ExcelToDataverseTool";
            panelTop.ResumeLayout(false);
            panelTop.PerformLayout();
            groupBoxTasks.ResumeLayout(false);
            panelBody.ResumeLayout(false);
            tabControlTask.ResumeLayout(false);
            tabPageLogs.ResumeLayout(false);
            tabPageConfig.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewColumnMapping).EndInit();
            panelConfigTop.ResumeLayout(false);
            panelConfigTop.PerformLayout();
            tabPageGuide.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel panelTop;
        private GroupBox groupBoxTasks;
        private Panel panelBody;
        private TabControl tabControlTask;
        private TabPage tabPageLogs;
        private TabPage tabPageConfig;
        private RichTextBox richTextBoxLogs;
        private Panel panelConfigTop;
        private DataGridView dataGridViewColumnMapping;
        private Label labelDataverseEntity;
        private Label labelExcelFile;
        private Button buttonAddMapping;
        private Button button2;
        private Button buttonBrowseLocalFile;
        private ComboBox comboBox2;
        private TextBox textBoxDataverseEntity;
        private TextBox textBoxExcelFile;
        private Label labelDataverseColumn;
        private Label labelExcelColumn;
        private TreeView treeViewTasks;
        private ContextMenuStrip contextMenuStripTask;
        private Label labelenvurl;
        private Button buttonSave;
        private TextBox textBoxClientSecret;
        private TextBox textBoxClientId;
        private TextBox textBoxTenantId;
        private TextBox textBoxEnvUrl;
        private Label labelClientSecret;
        private Label labelClientId;
        private Label labelTenantId;
        private Button buttonRunAll;
        private Button buttonNewTask;
        private Button buttonRunTask;
        private TextBox textBoxDataverseColumnName;
        private TextBox textBoxExcelColumnName;
        private TextBox textBoxEntityName;
        private ComboBox comboBoxDataverseColumnType;
        private Label labelColumnDataType;
        private CheckBox checkBoxPrimaryKey;
        private DataGridViewTextBoxColumn ColumnExcelColumn;
        private DataGridViewTextBoxColumn ColumnDataverseColumn;
        private DataGridViewTextBoxColumn ColumnType;
        private DataGridViewTextBoxColumn ColumnIsPrimary;
        private Button buttonSaveTask;
        private OpenFileDialog openFileDialog1;
        private Button buttonDeleteTask;
        private TabPage tabPageGuide;
        private RichTextBox richTextBoxGuide;
        private Label labelActionType;
        private ComboBox comboBoxActionType;
    }
}
