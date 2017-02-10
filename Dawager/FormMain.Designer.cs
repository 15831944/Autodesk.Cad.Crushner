namespace Dawager
{
    partial class FormMain
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.ToolStripSeparator файлSeparatorToolStripMenuItem1;
            this.m_menuStripMain = new System.Windows.Forms.MenuStrip();
            this.чертежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.шаблонToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.чертежToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.выходToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.m_statusStripMain = new System.Windows.Forms.StatusStrip();
            this.m_dgvProperties = new System.Windows.Forms.DataGridView();
            this.m_treeViewBlockDefinition = new System.Windows.Forms.TreeView();
            this.m_listBoxFileSettings = new System.Windows.Forms.ListBox();
            this.m_clbBlockReferences = new System.Windows.Forms.CheckedListBox();
            this.m_buttonApply = new System.Windows.Forms.Button();
            this.buttonAcadView = new System.Windows.Forms.Button();
            this.m_buttonExit = new System.Windows.Forms.Button();
            файлSeparatorToolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.m_menuStripMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dgvProperties)).BeginInit();
            this.SuspendLayout();
            // 
            // файлSeparatorToolStripMenuItem1
            // 
            файлSeparatorToolStripMenuItem1.Name = "файлSeparatorToolStripMenuItem1";
            файлSeparatorToolStripMenuItem1.Size = new System.Drawing.Size(111, 6);
            // 
            // m_menuStripMain
            // 
            this.m_menuStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.чертежToolStripMenuItem});
            this.m_menuStripMain.Location = new System.Drawing.Point(0, 0);
            this.m_menuStripMain.Name = "m_menuStripMain";
            this.m_menuStripMain.Size = new System.Drawing.Size(580, 24);
            this.m_menuStripMain.TabIndex = 0;
            this.m_menuStripMain.Text = "menuStrip1";
            // 
            // чертежToolStripMenuItem
            // 
            this.чертежToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.шаблонToolStripMenuItem,
            this.чертежToolStripMenuItem1,
            файлSeparatorToolStripMenuItem1,
            this.выходToolStripMenuItem1});
            this.чертежToolStripMenuItem.Name = "чертежToolStripMenuItem";
            this.чертежToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.чертежToolStripMenuItem.Text = "Файл";
            // 
            // шаблонToolStripMenuItem
            // 
            this.шаблонToolStripMenuItem.Name = "шаблонToolStripMenuItem";
            this.шаблонToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.шаблонToolStripMenuItem.Text = "Шаблон";
            // 
            // чертежToolStripMenuItem1
            // 
            this.чертежToolStripMenuItem1.Name = "чертежToolStripMenuItem1";
            this.чертежToolStripMenuItem1.Size = new System.Drawing.Size(114, 22);
            this.чертежToolStripMenuItem1.Text = "Чертеж";
            // 
            // выходToolStripMenuItem1
            // 
            this.выходToolStripMenuItem1.Name = "выходToolStripMenuItem1";
            this.выходToolStripMenuItem1.Size = new System.Drawing.Size(114, 22);
            this.выходToolStripMenuItem1.Text = "Выход";
            this.выходToolStripMenuItem1.Click += new System.EventHandler(this.выходToolStripMenuItem1_Click);
            // 
            // m_statusStripMain
            // 
            this.m_statusStripMain.Location = new System.Drawing.Point(0, 469);
            this.m_statusStripMain.Name = "m_statusStripMain";
            this.m_statusStripMain.Size = new System.Drawing.Size(580, 22);
            this.m_statusStripMain.TabIndex = 1;
            this.m_statusStripMain.Text = "statusStrip1";
            // 
            // m_dgvProperties
            // 
            this.m_dgvProperties.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.m_dgvProperties.Location = new System.Drawing.Point(0, 294);
            this.m_dgvProperties.Name = "m_dgvProperties";
            this.m_dgvProperties.Size = new System.Drawing.Size(275, 143);
            this.m_dgvProperties.TabIndex = 9;
            // 
            // m_treeViewBlockDefinition
            // 
            this.m_treeViewBlockDefinition.CheckBoxes = true;
            this.m_treeViewBlockDefinition.Location = new System.Drawing.Point(0, 101);
            this.m_treeViewBlockDefinition.Name = "m_treeViewBlockDefinition";
            this.m_treeViewBlockDefinition.Size = new System.Drawing.Size(275, 187);
            this.m_treeViewBlockDefinition.TabIndex = 10;
            // 
            // m_listBoxFileSettings
            // 
            this.m_listBoxFileSettings.FormattingEnabled = true;
            this.m_listBoxFileSettings.Location = new System.Drawing.Point(0, 27);
            this.m_listBoxFileSettings.Name = "m_listBoxFileSettings";
            this.m_listBoxFileSettings.Size = new System.Drawing.Size(275, 69);
            this.m_listBoxFileSettings.TabIndex = 11;
            this.m_listBoxFileSettings.SelectedIndexChanged += new System.EventHandler(this.listBoxFileSettings_SelectedIndexChanged);
            // 
            // m_clbBlockReferences
            // 
            this.m_clbBlockReferences.FormattingEnabled = true;
            this.m_clbBlockReferences.Location = new System.Drawing.Point(281, 27);
            this.m_clbBlockReferences.Name = "m_clbBlockReferences";
            this.m_clbBlockReferences.Size = new System.Drawing.Size(299, 409);
            this.m_clbBlockReferences.TabIndex = 12;
            // 
            // m_buttonApply
            // 
            this.m_buttonApply.Location = new System.Drawing.Point(336, 442);
            this.m_buttonApply.Name = "m_buttonApply";
            this.m_buttonApply.Size = new System.Drawing.Size(75, 23);
            this.m_buttonApply.TabIndex = 13;
            this.m_buttonApply.Text = "Применить";
            this.m_buttonApply.UseVisualStyleBackColor = true;
            // 
            // buttonAcadView
            // 
            this.buttonAcadView.Location = new System.Drawing.Point(417, 442);
            this.buttonAcadView.Name = "buttonAcadView";
            this.buttonAcadView.Size = new System.Drawing.Size(75, 23);
            this.buttonAcadView.TabIndex = 14;
            this.buttonAcadView.Text = "Просмотр";
            this.buttonAcadView.UseVisualStyleBackColor = true;
            // 
            // m_buttonExit
            // 
            this.m_buttonExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.m_buttonExit.Location = new System.Drawing.Point(498, 442);
            this.m_buttonExit.Name = "m_buttonExit";
            this.m_buttonExit.Size = new System.Drawing.Size(75, 23);
            this.m_buttonExit.TabIndex = 15;
            this.m_buttonExit.Text = "Выход";
            this.m_buttonExit.UseVisualStyleBackColor = true;
            this.m_buttonExit.Click += new System.EventHandler(this.m_buttonExit_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.m_buttonExit;
            this.ClientSize = new System.Drawing.Size(580, 491);
            this.Controls.Add(this.m_buttonExit);
            this.Controls.Add(this.buttonAcadView);
            this.Controls.Add(this.m_buttonApply);
            this.Controls.Add(this.m_clbBlockReferences);
            this.Controls.Add(this.m_listBoxFileSettings);
            this.Controls.Add(this.m_treeViewBlockDefinition);
            this.Controls.Add(this.m_dgvProperties);
            this.Controls.Add(this.m_statusStripMain);
            this.Controls.Add(this.m_menuStripMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.m_menuStripMain;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Управление чертежом AutoCad";
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.m_menuStripMain.ResumeLayout(false);
            this.m_menuStripMain.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dgvProperties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip m_menuStripMain;
        private System.Windows.Forms.ToolStripMenuItem чертежToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem шаблонToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem чертежToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem выходToolStripMenuItem1;
        private System.Windows.Forms.StatusStrip m_statusStripMain;
        private System.Windows.Forms.DataGridView m_dgvProperties;
        private System.Windows.Forms.TreeView m_treeViewBlockDefinition;
        private System.Windows.Forms.ListBox m_listBoxFileSettings;
        private System.Windows.Forms.CheckedListBox m_clbBlockReferences;
        private System.Windows.Forms.Button m_buttonApply;
        private System.Windows.Forms.Button buttonAcadView;
        private System.Windows.Forms.Button m_buttonExit;
    }
}

