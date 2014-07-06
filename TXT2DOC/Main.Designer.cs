namespace TXT2DOC
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
			this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.toostripmenu1_show = new System.Windows.Forms.ToolStripMenuItem();
			this.toostripmenu2_exit = new System.Windows.Forms.ToolStripMenuItem();
			this.timer = new System.Windows.Forms.Timer(this.components);
			this.generate_button = new System.Windows.Forms.Button();
			this.TOC_groupbox = new System.Windows.Forms.GroupBox();
			this.TOC_list = new System.Windows.Forms.DataGridView();
			this.TOC_list_ChapterTitle = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.TOC_list_LineNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.TOC_label = new System.Windows.Forms.Label();
			this.statusStrip1 = new System.Windows.Forms.StatusStrip();
			this.status_label = new System.Windows.Forms.ToolStripStatusLabel();
			this.menu_menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.menu1_switchAOT = new System.Windows.Forms.ToolStripMenuItem();
			this.menu2_settings = new System.Windows.Forms.ToolStripMenuItem();
			this.menu3_export = new System.Windows.Forms.ToolStripMenuItem();
			this.menu4_import = new System.Windows.Forms.ToolStripMenuItem();
			this.menu5_clear = new System.Windows.Forms.ToolStripMenuItem();
			this.menu6_help = new System.Windows.Forms.ToolStripMenuItem();
			this.menu6_1_helpfile_menuitem = new System.Windows.Forms.ToolStripMenuItem();
			this.menu6_2_toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.menu6_3_about_menuitem = new System.Windows.Forms.ToolStripMenuItem();
			this.contextMenuStrip1.SuspendLayout();
			this.TOC_groupbox.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.TOC_list)).BeginInit();
			this.statusStrip1.SuspendLayout();
			this.menu_menuStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// notifyIcon1
			// 
			this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
			this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
			this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
			this.notifyIcon1.Text = "TXT2DOC";
			this.notifyIcon1.Visible = true;
			this.notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toostripmenu1_show,
            this.toostripmenu2_exit});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(156, 70);
			// 
			// toostripmenu1_show
			// 
			this.toostripmenu1_show.Name = "toostripmenu1_show";
			this.toostripmenu1_show.Size = new System.Drawing.Size(155, 22);
			this.toostripmenu1_show.Text = "显示 TXT2DOC";
			this.toostripmenu1_show.Click += new System.EventHandler(this.toostripmenu1_show_Click);
			// 
			// toostripmenu2_exit
			// 
			this.toostripmenu2_exit.Name = "toostripmenu2_exit";
			this.toostripmenu2_exit.Size = new System.Drawing.Size(155, 22);
			this.toostripmenu2_exit.Text = "退出";
			this.toostripmenu2_exit.Click += new System.EventHandler(this.toostripmenu2_exit_Click);
			// 
			// timer
			// 
			this.timer.Interval = 1000;
			this.timer.Tick += new System.EventHandler(this.timer_Tick);
			// 
			// generate_button
			// 
			this.generate_button.Cursor = System.Windows.Forms.Cursors.Hand;
			this.generate_button.Location = new System.Drawing.Point(122, 540);
			this.generate_button.Name = "generate_button";
			this.generate_button.Size = new System.Drawing.Size(300, 45);
			this.generate_button.TabIndex = 5;
			this.generate_button.Text = ">>> 开始生成 <<<";
			this.generate_button.UseVisualStyleBackColor = true;
			this.generate_button.Click += new System.EventHandler(this.generate_button_Click);
			// 
			// TOC_groupbox
			// 
			this.TOC_groupbox.Controls.Add(this.TOC_list);
			this.TOC_groupbox.Controls.Add(this.TOC_label);
			this.TOC_groupbox.ForeColor = System.Drawing.Color.RoyalBlue;
			this.TOC_groupbox.Location = new System.Drawing.Point(12, 35);
			this.TOC_groupbox.Name = "TOC_groupbox";
			this.TOC_groupbox.Size = new System.Drawing.Size(520, 490);
			this.TOC_groupbox.TabIndex = 4;
			this.TOC_groupbox.TabStop = false;
			this.TOC_groupbox.Text = "小提示：拖动TXT文件到下面，以下列表即为书籍目录";
			// 
			// TOC_list
			// 
			this.TOC_list.AllowDrop = true;
			this.TOC_list.AllowUserToResizeColumns = false;
			this.TOC_list.AllowUserToResizeRows = false;
			this.TOC_list.BackgroundColor = System.Drawing.Color.White;
			this.TOC_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.TOC_list.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.TOC_list_ChapterTitle,
            this.TOC_list_LineNumber});
			this.TOC_list.Location = new System.Drawing.Point(6, 40);
			this.TOC_list.Name = "TOC_list";
			this.TOC_list.RowHeadersWidth = 25;
			this.TOC_list.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.TOC_list.Size = new System.Drawing.Size(508, 444);
			this.TOC_list.TabIndex = 2;
			this.TOC_list.DragDrop += new System.Windows.Forms.DragEventHandler(this.TOC_list_DragDrop);
			this.TOC_list.DragEnter += new System.Windows.Forms.DragEventHandler(this.TOC_list_DragEnter);
			// 
			// TOC_list_ChapterTitle
			// 
			this.TOC_list_ChapterTitle.HeaderText = "章节名称";
			this.TOC_list_ChapterTitle.Name = "TOC_list_ChapterTitle";
			this.TOC_list_ChapterTitle.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.TOC_list_ChapterTitle.Width = 325;
			// 
			// TOC_list_LineNumber
			// 
			this.TOC_list_LineNumber.HeaderText = "行号";
			this.TOC_list_LineNumber.Name = "TOC_list_LineNumber";
			this.TOC_list_LineNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.TOC_list_LineNumber.Width = 125;
			// 
			// TOC_label
			// 
			this.TOC_label.AutoSize = true;
			this.TOC_label.Location = new System.Drawing.Point(59, 20);
			this.TOC_label.Name = "TOC_label";
			this.TOC_label.Size = new System.Drawing.Size(242, 17);
			this.TOC_label.TabIndex = 1;
			this.TOC_label.Text = "可以编辑单行，也可以导出后编辑再导入";
			// 
			// statusStrip1
			// 
			this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.status_label});
			this.statusStrip1.Location = new System.Drawing.Point(0, 599);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 18, 0);
			this.statusStrip1.Size = new System.Drawing.Size(544, 22);
			this.statusStrip1.SizingGrip = false;
			this.statusStrip1.TabIndex = 3;
			// 
			// status_label
			// 
			this.status_label.Name = "status_label";
			this.status_label.Size = new System.Drawing.Size(0, 17);
			this.status_label.TextChanged += new System.EventHandler(this.status_label_TextChanged);
			// 
			// menu_menuStrip1
			// 
			this.menu_menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menu1_switchAOT,
            this.menu2_settings,
            this.menu3_export,
            this.menu4_import,
            this.menu5_clear,
            this.menu6_help});
			this.menu_menuStrip1.Location = new System.Drawing.Point(0, 0);
			this.menu_menuStrip1.Name = "menu_menuStrip1";
			this.menu_menuStrip1.Padding = new System.Windows.Forms.Padding(8, 2, 0, 2);
			this.menu_menuStrip1.Size = new System.Drawing.Size(544, 24);
			this.menu_menuStrip1.TabIndex = 1;
			this.menu_menuStrip1.Text = "menuStrip1";
			// 
			// menu1_switchAOT
			// 
			this.menu1_switchAOT.Name = "menu1_switchAOT";
			this.menu1_switchAOT.ShortcutKeyDisplayString = "";
			this.menu1_switchAOT.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.T)));
			this.menu1_switchAOT.Size = new System.Drawing.Size(86, 20);
			this.menu1_switchAOT.Text = "置顶切换(&T)";
			this.menu1_switchAOT.Click += new System.EventHandler(this.menu1_switchAOT_Click);
			// 
			// menu2_settings
			// 
			this.menu2_settings.Name = "menu2_settings";
			this.menu2_settings.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.S)));
			this.menu2_settings.Size = new System.Drawing.Size(59, 20);
			this.menu2_settings.Text = "设置(&S)";
			this.menu2_settings.Click += new System.EventHandler(this.menu2_settings_Click);
			// 
			// menu3_export
			// 
			this.menu3_export.Name = "menu3_export";
			this.menu3_export.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.X)));
			this.menu3_export.Size = new System.Drawing.Size(86, 20);
			this.menu3_export.Text = "导出目录(&X)";
			this.menu3_export.Click += new System.EventHandler(this.menu3_export_Click);
			// 
			// menu4_import
			// 
			this.menu4_import.Name = "menu4_import";
			this.menu4_import.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.I)));
			this.menu4_import.Size = new System.Drawing.Size(82, 20);
			this.menu4_import.Text = "导入目录(&I)";
			this.menu4_import.Click += new System.EventHandler(this.menu4_import_Click);
			// 
			// menu5_clear
			// 
			this.menu5_clear.Name = "menu5_clear";
			this.menu5_clear.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.C)));
			this.menu5_clear.Size = new System.Drawing.Size(87, 20);
			this.menu5_clear.Text = "清空目录(&C)";
			this.menu5_clear.Click += new System.EventHandler(this.menu5_clear_Click);
			// 
			// menu6_help
			// 
			this.menu6_help.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menu6_1_helpfile_menuitem,
            this.menu6_2_toolStripSeparator2,
            this.menu6_3_about_menuitem});
			this.menu6_help.Name = "menu6_help";
			this.menu6_help.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.H)));
			this.menu6_help.Size = new System.Drawing.Size(62, 20);
			this.menu6_help.Text = "帮助(&H)";
			// 
			// menu6_1_helpfile_menuitem
			// 
			this.menu6_1_helpfile_menuitem.Name = "menu6_1_helpfile_menuitem";
			this.menu6_1_helpfile_menuitem.Size = new System.Drawing.Size(126, 22);
			this.menu6_1_helpfile_menuitem.Text = "查看帮助";
			this.menu6_1_helpfile_menuitem.Click += new System.EventHandler(this.menu6_1_helpfile_menuitem_Click);
			// 
			// menu6_2_toolStripSeparator2
			// 
			this.menu6_2_toolStripSeparator2.Name = "menu6_2_toolStripSeparator2";
			this.menu6_2_toolStripSeparator2.Size = new System.Drawing.Size(123, 6);
			// 
			// menu6_3_about_menuitem
			// 
			this.menu6_3_about_menuitem.Name = "menu6_3_about_menuitem";
			this.menu6_3_about_menuitem.Size = new System.Drawing.Size(126, 22);
			this.menu6_3_about_menuitem.Text = "关于…";
			this.menu6_3_about_menuitem.Click += new System.EventHandler(this.menu6_3_about_menuitem_Click);
			// 
			// Main
			// 
			this.AllowDrop = true;
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(544, 621);
			this.Controls.Add(this.generate_button);
			this.Controls.Add(this.TOC_groupbox);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.menu_menuStrip1);
			this.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MainMenuStrip = this.menu_menuStrip1;
			this.MaximizeBox = false;
			this.Name = "Main";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "TXT2DOC";
			this.Load += new System.EventHandler(this.Main_Load);
			this.contextMenuStrip1.ResumeLayout(false);
			this.TOC_groupbox.ResumeLayout(false);
			this.TOC_groupbox.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.TOC_list)).EndInit();
			this.statusStrip1.ResumeLayout(false);
			this.statusStrip1.PerformLayout();
			this.menu_menuStrip1.ResumeLayout(false);
			this.menu_menuStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menu_menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menu1_switchAOT;
        private System.Windows.Forms.ToolStripMenuItem menu2_settings;
        private System.Windows.Forms.ToolStripMenuItem menu3_export;
        private System.Windows.Forms.ToolStripMenuItem menu4_import;
        private System.Windows.Forms.ToolStripMenuItem menu5_clear;
        private System.Windows.Forms.ToolStripMenuItem menu6_help;
        private System.Windows.Forms.ToolStripMenuItem menu6_1_helpfile_menuitem;
        private System.Windows.Forms.ToolStripSeparator menu6_2_toolStripSeparator2;
		private System.Windows.Forms.ToolStripMenuItem menu6_3_about_menuitem;
        private System.Windows.Forms.GroupBox TOC_groupbox;
        private System.Windows.Forms.Label TOC_label;
        private System.Windows.Forms.Button generate_button;
		private System.Windows.Forms.DataGridView TOC_list;
		private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
		private System.Windows.Forms.ToolStripMenuItem toostripmenu1_show;
		private System.Windows.Forms.ToolStripMenuItem toostripmenu2_exit;
		private System.Windows.Forms.NotifyIcon notifyIcon1;
		private System.Windows.Forms.DataGridViewTextBoxColumn TOC_list_ChapterTitle;
		private System.Windows.Forms.DataGridViewTextBoxColumn TOC_list_LineNumber;
		private System.Windows.Forms.StatusStrip statusStrip1;
		private System.Windows.Forms.ToolStripStatusLabel status_label;
		private System.Windows.Forms.Timer timer;
    }
}

