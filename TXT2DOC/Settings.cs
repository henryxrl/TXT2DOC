using Ini;
using System;
using System.IO;
using System.Windows.Forms;
/*
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using Ini;
*/

namespace TXT2DOC
{
	public partial class Settings : Tools
    {
		String iniPath = Path.Combine(Path.GetTempPath(), "TXT2DOC") + "\\settings.ini";

        public Settings()
        {
            InitializeComponent();
        }

		private void Settings_Load(object sender, EventArgs e)
		{
			IniFile ini = loadINI(iniPath);
			try
			{
				loadCurSettings(this, ini);
			}
			catch
			{
				MessageBox.Show("加载设置文件出错，即将导入默认设置！");
				writeDefaultSettings(ini);
				loadCurSettings(this, ini);
			}
		}

		private void settings4_3_reset_button_Click(object sender, EventArgs e)
		{
			DialogResult dialogResult = MessageBox.Show("确认还原默认设置吗？", "确认", MessageBoxButtons.YesNo);
			if (dialogResult == DialogResult.Yes)
			{
				saveSettings(0, this, iniPath);
			}
		}

		private void settings_done_button_Click(object sender, EventArgs e)
		{
			saveSettings(1, this, iniPath);
		}

		private void settings4_1_filelocation_button_Click(object sender, EventArgs e)
		{
			if (settings4_1_filelocation_dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				settings4_1_filelocation_textbox.Text = settings4_1_filelocation_dialog.SelectedPath;
			}
		}

		private void settings1_1_TOC_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings1_1_TOCload_checkbox.Enabled = this.settings1_1_TOC_checkbox.Checked;

			if (!this.settings1_1_TOCload_checkbox.Enabled)
			{
				this.settings1_1_TOCload_checkbox.Checked = false;
				this.settings1_1_TOCload_label.Enabled = false;
			}
			else
			{
				this.settings1_1_TOCload_checkbox.Checked = false;
				this.settings1_1_TOCload_label.Enabled = true;
			}
		}

		private void settings5_1_header_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings5_1_headersize_label.Enabled = this.settings5_1_header_checkbox.Checked;
			this.settings5_1_headersize_textbox.Enabled = this.settings5_1_header_checkbox.Checked;
			this.settings5_1_headeralign_label.Enabled = this.settings5_1_header_checkbox.Checked;
			this.settings5_1_headeralign_combobox.Enabled = this.settings5_1_header_checkbox.Checked;
			this.settings5_1_headerboder_checkbox.Enabled = this.settings5_1_header_checkbox.Checked;

			if (!this.settings5_1_headerboder_checkbox.Enabled)
			{
				this.settings5_1_headerboder_checkbox.Checked = false;
			}
			else
			{
				this.settings5_1_headerboder_checkbox.Checked = true;
			}
		}

		private void settings5_2_footer_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings5_2_footersize_label.Enabled = this.settings5_2_footer_checkbox.Checked;
			this.settings5_2_footersize_textbox.Enabled = this.settings5_2_footer_checkbox.Checked;
			this.settings5_2_footeralign_label.Enabled = this.settings5_2_footer_checkbox.Checked;
			this.settings5_2_footeralign_combobox.Enabled = this.settings5_2_footer_checkbox.Checked;
			this.settings5_2_footerstyle_label.Enabled = this.settings5_2_footer_checkbox.Checked;
			this.settings5_2_footerstyle_combobox.Enabled = this.settings5_2_footer_checkbox.Checked;
			this.settings5_2_footerboder_checkbox.Enabled = this.settings5_2_footer_checkbox.Checked;

			if (!this.settings5_2_footerboder_checkbox.Enabled)
			{
				this.settings5_2_footerboder_checkbox.Checked = false;
			}
			else
			{
				this.settings5_2_footerboder_checkbox.Checked = true;
			}
		}

		private void settings4_2_savetimemode_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings4_2_pdf_checkbox.Enabled = !this.settings4_2_savetimemode_checkbox.Checked;

			if (!this.settings4_2_pdf_checkbox.Enabled)
			{
				this.settings4_2_pdf_checkbox.Checked = false;
				this.settings4_2_pdf_label.Enabled = false;
			}
			else
			{
				this.settings4_2_pdf_checkbox.Checked = false;
				this.settings4_2_pdf_label.Enabled = true;
			}
		}

		private void settings4_2_pdf_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings4_2_savetimemode_checkbox.Enabled = !this.settings4_2_pdf_checkbox.Checked;

			if (!this.settings4_2_savetimemode_checkbox.Enabled)
			{
				this.settings4_2_savetimemode_checkbox.Checked = false;
				this.settings4_2_savetimemode_checkbox.Enabled = false;
			}
			else
			{
				this.settings4_2_savetimemode_checkbox.Checked = false;
				this.settings4_2_savetimemode_checkbox.Enabled = true;
			}
		}

		private void settings1_3_StT_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings1_3_TtS_checkbox.Enabled = !this.settings1_3_StT_checkbox.Checked;

			if (!this.settings1_3_TtS_checkbox.Enabled)
			{
				this.settings1_3_TtS_checkbox.Checked = false;
				this.settings1_3_TtS_checkbox.Enabled = false;
			}
			else
			{
				this.settings1_3_TtS_checkbox.Checked = false;
				this.settings1_3_TtS_checkbox.Enabled = true;
			}
		}

		private void settings1_3_TtS_checkbox_CheckedChanged(object sender, EventArgs e)
		{
			this.settings1_3_StT_checkbox.Enabled = !this.settings1_3_TtS_checkbox.Checked;

			if (!this.settings1_3_StT_checkbox.Enabled)
			{
				this.settings1_3_StT_checkbox.Checked = false;
				this.settings1_3_StT_checkbox.Enabled = false;
			}
			else
			{
				this.settings1_3_StT_checkbox.Checked = false;
				this.settings1_3_StT_checkbox.Enabled = true;
			}
		}
	}
}
