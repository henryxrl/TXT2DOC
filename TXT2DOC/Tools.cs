using Ini;
using System;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using Ini;
*/

namespace TXT2DOC
{
	public class Tools : Form
	{
		public IniFile loadINI(String iniPath)
		{
			IniFile ini = new IniFile(iniPath);
			if (!File.Exists(iniPath))
			{
				writeDefaultSettings(ini);
			}
			return ini;
		}

		public void saveSettings(int flag, Settings settingsForm, String iniPath)		// flag == 0: restore default; flag == 1: save current
		{
			IniFile ini = new IniFile(iniPath);
			FileInfo iniInfo = new FileInfo(iniPath);
			if (!File.Exists(iniPath) || !iniInfo.IsReadOnly)
			{
				if (flag == 0) writeDefaultSettings(ini);
				else writeCurSettings(settingsForm, ini);
				loadCurSettings(settingsForm, ini);
				this.Close();
			}
			else
			{
				MessageBox.Show("写入设置文件出错，可能是设置文件被设为只读。\n请取消其只读状态或删除设置文件，并点击“确认”键重试！");
			}
		}

		public void loadCurSettings(Settings settingsForm, IniFile ini)
		{
			settingsForm.settings1_1_cover_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Cover")));
			settingsForm.settings1_1_TOC_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "TOC")));
			settingsForm.settings1_1_TOCload_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "TOC_Load")));
			if (!settingsForm.settings1_1_TOC_checkbox.Checked)
			{
				settingsForm.settings1_1_TOCload_checkbox.Checked = false;
				settingsForm.settings1_1_TOCload_checkbox.Enabled = false;
				settingsForm.settings1_1_TOCload_label.Enabled = false;
			}
			settingsForm.settings1_3_vertical_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Vertical")));
			settingsForm.settings1_3_replace_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Replace")));
			settingsForm.settings1_3_StT_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "StT")));
			settingsForm.settings1_3_TtS_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "TtS")));
			if (settingsForm.settings1_3_StT_checkbox.Checked)
			{
				settingsForm.settings1_3_TtS_checkbox.Checked = false;
				settingsForm.settings1_3_TtS_checkbox.Enabled = false;
			}
			if (settingsForm.settings1_3_TtS_checkbox.Checked)
			{
				settingsForm.settings1_3_StT_checkbox.Checked = false;
				settingsForm.settings1_3_StT_checkbox.Enabled = false;
			}
			//settingsForm.settings1_3_dropcap_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Drop_Cap")));
			settingsForm.settings2_1_pw_textbox.Text = ini.IniReadValue("Tab_2", "Page_Width");
			settingsForm.settings2_1_ph_textbox.Text = ini.IniReadValue("Tab_2", "Page_Height");
			settingsForm.settings2_1_pc_combobox.Text = convertCodeToColor(ini.IniReadValue("Tab_2", "Page_Color"));
			settingsForm.settings2_2_pmU_textbox.Text = ini.IniReadValue("Tab_2", "Page_Margin_Up");
			settingsForm.settings2_2_pmD_textbox.Text = ini.IniReadValue("Tab_2", "Page_Margin_Down");
			settingsForm.settings2_2_pmL_textbox.Text = ini.IniReadValue("Tab_2", "Page_Margin_Left");
			settingsForm.settings2_2_pmR_textbox.Text = ini.IniReadValue("Tab_2", "Page_Margin_Right");
			settingsForm.settings2_3_header_textbox.Text = ini.IniReadValue("Tab_2", "Header_Margin");
			settingsForm.settings2_3_footer_textbox.Text = ini.IniReadValue("Tab_2", "Footer_Margin");
			settingsForm.settings3_1_tfont_combobox.Text = ini.IniReadValue("Tab_3", "Title_Font");
			settingsForm.settings3_1_tsize_textbox.Text = ini.IniReadValue("Tab_3", "Title_Size");
			settingsForm.settings3_1_tcolor_combobox.Text = convertCodeToColor(ini.IniReadValue("Tab_3", "Title_Color"));
			settingsForm.settings3_2_bfont_combobox.Text = ini.IniReadValue("Tab_3", "Body_Font");
			settingsForm.settings3_2_bsize_textbox.Text = ini.IniReadValue("Tab_3", "Body_Size");
			settingsForm.settings3_2_bcolor_combobox.Text = convertCodeToColor(ini.IniReadValue("Tab_3", "Body_Color"));
			settingsForm.settings3_3_linespacing_textbox.Text = ini.IniReadValue("Tab_3", "Line_Spacing");
			settingsForm.settings3_3_addparagraphspacing_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_3", "Add_Paragraph_Spacing")));
			settingsForm.settings4_1_filelocation_textbox.Text = ini.IniReadValue("Tab_4", "Generated_File_Location");
			settingsForm.settings4_2_dragclearlist_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "Drag_Clear_List")));
			settingsForm.settings4_2_deletetempfiles_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "Delete_Temp_Files")));
			settingsForm.settings4_2_pdf_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "PDF")));
			settingsForm.settings4_2_savetimemode_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "Save_Time_Mode")));
			if (settingsForm.settings4_2_savetimemode_checkbox.Checked)
			{
				settingsForm.settings4_2_pdf_checkbox.Checked = false;
				settingsForm.settings4_2_pdf_checkbox.Enabled = false;
				settingsForm.settings4_2_pdf_label.Enabled = false;
			}
			settingsForm.settings5_1_header_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Header")));
			settingsForm.settings5_1_headersize_textbox.Text = ini.IniReadValue("Tab_5", "Header_Size");
			settingsForm.settings5_1_headeralign_combobox.Text = convertNumToHan_Align(ini.IniReadValue("Tab_5", "Header_Align"));
			settingsForm.settings5_1_headerboder_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Header_Border")));
			if (!settingsForm.settings5_1_header_checkbox.Checked)
			{
				settingsForm.settings5_1_headersize_textbox.Enabled = false;
				settingsForm.settings5_1_headeralign_combobox.Enabled = false;
				settingsForm.settings5_1_headerboder_checkbox.Enabled = false;
			}
			settingsForm.settings5_2_footer_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Footer")));
			settingsForm.settings5_2_footersize_textbox.Text = ini.IniReadValue("Tab_5", "Footer_Size");
			settingsForm.settings5_2_footeralign_combobox.Text = convertNumToHan_Align(ini.IniReadValue("Tab_5", "Footer_Align"));
			settingsForm.settings5_2_footerstyle_combobox.Text = convertNumToHan_Style(ini.IniReadValue("Tab_5", "Footer_Style"));
			settingsForm.settings5_2_footerboder_checkbox.Checked = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Footer_Border")));
			if (!settingsForm.settings5_2_footer_checkbox.Checked)
			{
				settingsForm.settings5_2_footersize_textbox.Enabled = false;
				settingsForm.settings5_2_footeralign_combobox.Enabled = false;
				settingsForm.settings5_2_footerstyle_combobox.Enabled = false;
				settingsForm.settings5_2_footerboder_checkbox.Enabled = false;
			}
		}

		public void writeCurSettings(Settings settingsForm, IniFile ini)
		{
			ini.IniWriteValue("Tab_1", "Cover", (Convert.ToInt32(settingsForm.settings1_1_cover_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_1", "TOC", (Convert.ToInt32(settingsForm.settings1_1_TOC_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_1", "TOC_Load", (Convert.ToInt32(settingsForm.settings1_1_TOCload_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_1", "Vertical", (Convert.ToInt32(settingsForm.settings1_3_vertical_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_1", "Replace", (Convert.ToInt32(settingsForm.settings1_3_replace_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_1", "StT", (Convert.ToInt32(settingsForm.settings1_3_StT_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_1", "TtS", (Convert.ToInt32(settingsForm.settings1_3_TtS_checkbox.Checked)).ToString());
			//ini.IniWriteValue("Tab_1", "Drop_Cap", (Convert.ToInt32(settingsForm.settings1_3_dropcap_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_2", "Page_Width", settingsForm.settings2_1_pw_textbox.Text);
			ini.IniWriteValue("Tab_2", "Page_Height", settingsForm.settings2_1_ph_textbox.Text);
			ini.IniWriteValue("Tab_2", "Page_Color", convertColorToCode(settingsForm.settings2_1_pc_combobox.Text, 0));
			ini.IniWriteValue("Tab_2", "Page_Margin_Up", settingsForm.settings2_2_pmU_textbox.Text);
			ini.IniWriteValue("Tab_2", "Page_Margin_Down", settingsForm.settings2_2_pmD_textbox.Text);
			ini.IniWriteValue("Tab_2", "Page_Margin_Left", settingsForm.settings2_2_pmL_textbox.Text);
			ini.IniWriteValue("Tab_2", "Page_Margin_Right", settingsForm.settings2_2_pmR_textbox.Text);
			ini.IniWriteValue("Tab_2", "Header_Margin", settingsForm.settings2_3_header_textbox.Text);
			ini.IniWriteValue("Tab_2", "Footer_Margin", settingsForm.settings2_3_footer_textbox.Text);
			ini.IniWriteValue("Tab_3", "Title_Font", settingsForm.settings3_1_tfont_combobox.Text);
			ini.IniWriteValue("Tab_3", "Title_Size", settingsForm.settings3_1_tsize_textbox.Text);
			ini.IniWriteValue("Tab_3", "Title_Color", convertColorToCode(settingsForm.settings3_1_tcolor_combobox.Text, 1));
			ini.IniWriteValue("Tab_3", "Body_Font", settingsForm.settings3_2_bfont_combobox.Text);
			ini.IniWriteValue("Tab_3", "Body_Size", settingsForm.settings3_2_bsize_textbox.Text);
			ini.IniWriteValue("Tab_3", "Body_Color", convertColorToCode(settingsForm.settings3_2_bcolor_combobox.Text, 1));
			ini.IniWriteValue("Tab_3", "Line_Spacing", settingsForm.settings3_3_linespacing_textbox.Text);
			ini.IniWriteValue("Tab_3", "Add_Paragraph_Spacing", (Convert.ToInt32(settingsForm.settings3_3_addparagraphspacing_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_4", "Generated_File_Location", settingsForm.settings4_1_filelocation_textbox.Text);
			ini.IniWriteValue("Tab_4", "Drag_Clear_List", (Convert.ToInt32(settingsForm.settings4_2_dragclearlist_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_4", "Delete_Temp_Files", (Convert.ToInt32(settingsForm.settings4_2_deletetempfiles_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_4", "PDF", (Convert.ToInt32(settingsForm.settings4_2_pdf_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_4", "Save_Time_Mode", (Convert.ToInt32(settingsForm.settings4_2_savetimemode_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_5", "Header", (Convert.ToInt32(settingsForm.settings5_1_header_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_5", "Header_Size", settingsForm.settings5_1_headersize_textbox.Text);
			ini.IniWriteValue("Tab_5", "Header_Align", convertHanToNum_Align(settingsForm.settings5_1_headeralign_combobox.Text));
			ini.IniWriteValue("Tab_5", "Header_Border", (Convert.ToInt32(settingsForm.settings5_1_headerboder_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_5", "Footer", (Convert.ToInt32(settingsForm.settings5_2_footer_checkbox.Checked)).ToString());
			ini.IniWriteValue("Tab_5", "Footer_Size", settingsForm.settings5_2_footersize_textbox.Text);
			ini.IniWriteValue("Tab_5", "Footer_Align", convertHanToNum_Align(settingsForm.settings5_2_footeralign_combobox.Text));
			ini.IniWriteValue("Tab_5", "Footer_Style", convertHanToNum_Style(settingsForm.settings5_2_footerstyle_combobox.Text));
			ini.IniWriteValue("Tab_5", "Footer_Border", (Convert.ToInt32(settingsForm.settings5_2_footerboder_checkbox.Checked)).ToString());
		}

		public void writeDefaultSettings(IniFile ini)
		{
			/*
			settings1_1_cover_checkbox = True;
			settings1_1_TOC_checkbox = True;
			settings1_1_TOCload_checkbox = True;
			settings1_3_vertical_checkbox = False;
			settings1_3_replace_checkbox = False;
			settings1_3_StT_checkbox = False;
			settings1_3_TtS_checkbox = False;
			settings1_3_dropcap_checkbox = False;
			settings2_1_pw_textbox = 4.8(4.89);
			settings2_1_ph_textbox = 6.5(6.5);
			settings2_1_pc_combobox = 白色;
			settings2_2_pmU_textbox = 0.4;
			settings2_2_pmD_textbox = 0.4;
			settings2_2_pmL_textbox = 0.2;
			settings2_2_pmR_textbox = 0.2;
			settings2_3_header_textbox = 0.1;
			settings2_3_footer_textbox = 0.1;
			settings3_1_tfont_combobox = 微软雅黑;
			settings3_1_tsize_textbox = 22;
			settings3_1_tcolor_combobox = 深蓝;
			settings3_2_bfont_combobox = 微软雅黑;
			settings3_2_bsize_textbox = 14;
			settings3_2_bcolor_combobox = 黑色;
			settings3_3_linespacing_textbox = 1.15;
			settings3_3_addparagraphspacing_checkbox = False;
			settings4_1_filelocation_textbox = Application.StartupPath;
			settings4_2_dragclearlist_checkbox = True;
			settings4_2_deletetempfiles_checkbox = True;
			settings4_2_pdf_checkbox = False;
			settings4_2_savetimemode_checkbox = False;
			settings5_1_header_checkbox = True;
			settings5_1_headersize_textbox = 11;
			settings5_1_headeralign_combobox = 1;
			settings5_1_headerboder_checkbox = 1;
			settings5_2_footer_checkbox = True;
			settings5_2_footersize_textbox = 11;
			settings5_2_footeralign_combobox = 1;
			settings5_2_footerstyle_combobox = 0;
			settings5_2_footerboder_checkbox = 1;
			*/
			ini.IniWriteValue("Tab_1", "Cover", "1");
			ini.IniWriteValue("Tab_1", "TOC", "1");
			ini.IniWriteValue("Tab_1", "TOC_Load", "1");
			ini.IniWriteValue("Tab_1", "Vertical", "0");
			ini.IniWriteValue("Tab_1", "Replace", "0");
			ini.IniWriteValue("Tab_1", "StT", "0");
			ini.IniWriteValue("Tab_1", "TtS", "0");
			//ini.IniWriteValue("Tab_1", "Drop_Cap", "0");
			ini.IniWriteValue("Tab_2", "Page_Width", "4.8");
			ini.IniWriteValue("Tab_2", "Page_Height", "6.5");
			ini.IniWriteValue("Tab_2", "Page_Color", "white");
			ini.IniWriteValue("Tab_2", "Page_Margin_Up", "0.4");
			ini.IniWriteValue("Tab_2", "Page_Margin_Down", "0.4");
			ini.IniWriteValue("Tab_2", "Page_Margin_Left", "0.2");
			ini.IniWriteValue("Tab_2", "Page_Margin_Right", "0.2");
			ini.IniWriteValue("Tab_2", "Header_Margin", "0.1");
			ini.IniWriteValue("Tab_2", "Footer_Margin", "0.1");
			ini.IniWriteValue("Tab_3", "Title_Font", "微软雅黑");
			ini.IniWriteValue("Tab_3", "Title_Size", "22");
			ini.IniWriteValue("Tab_3", "Title_Color", "#365F91");
			ini.IniWriteValue("Tab_3", "Body_Font", "微软雅黑");
			ini.IniWriteValue("Tab_3", "Body_Size", "14");
			ini.IniWriteValue("Tab_3", "Body_Color", "black");
			ini.IniWriteValue("Tab_3", "Line_Spacing", "1.15");
			ini.IniWriteValue("Tab_3", "Add_Paragraph_Spacing", "0");
			ini.IniWriteValue("Tab_4", "Generated_File_Location", Application.StartupPath);
			ini.IniWriteValue("Tab_4", "Drag_Clear_List", "1");
			ini.IniWriteValue("Tab_4", "Delete_Temp_Files", "1");
			ini.IniWriteValue("Tab_4", "PDF", "0");
			ini.IniWriteValue("Tab_4", "Save_Time_Mode", "0");
			ini.IniWriteValue("Tab_5", "Header", "1");
			ini.IniWriteValue("Tab_5", "Header_Size", "11");
			ini.IniWriteValue("Tab_5", "Header_Align", "2");
			ini.IniWriteValue("Tab_5", "Header_Border", "1");
			ini.IniWriteValue("Tab_5", "Footer", "1");
			ini.IniWriteValue("Tab_5", "Footer_Size", "11");
			ini.IniWriteValue("Tab_5", "Footer_Align", "1");
			ini.IniWriteValue("Tab_5", "Footer_Style", "0");
			ini.IniWriteValue("Tab_5", "Footer_Border", "1");
		}

		public String convertColorToCode(String input, int flag)		// flag == 0: from settingsForm.settings2_1_pc_combobox.Text; else otherwise
		{
			String text = input.Trim();
			if (String.Compare(text, "白色") == 0)
			{
				return "white";
			}
			else if (String.Compare(text, "黑色") == 0)
			{
				return "black";
			}
			else if (String.Compare(text, "红色") == 0)
			{
				return "red";
			}
			else if (String.Compare(text, "黄色") == 0)
			{
				return "yellow";
			}
			else if (String.Compare(text, "蓝色") == 0)
			{
				return "blue";
			}
			else if (String.Compare(text, "绿色") == 0)
			{
				return "green";
			}
			else if (String.Compare(text, "紫色") == 0)
			{
				return "purple";
			}
			else if (String.Compare(text, "橙色") == 0)
			{
				return "orange";
			}
			else if (String.Compare(text, "护眼绿") == 0)
			{
				return "#B2DEC2";
			}
			else if (String.Compare(text, "深蓝") == 0)
			{
				return "#365F91";
			}
			else if (String.Compare(text, "浅蓝") == 0)
			{
				return "#4F81BD";
			}
			else if (String.Compare(text, "羊皮纸") == 0)
			{
				return "羊皮纸";
			}
			else
			{
				if (containChineseChar(text))
				{
					if (flag == 0) return "white";
					else return "black";
				}
				else
				{
					// if input doesn't contain Chinese character
					// treat it as HTML code
					return text;
				}
			}
		}

		private bool containChineseChar(String text)
		{
			bool containUnicode = false;
			for (int x = 0; x < text.Length; x++)
			{
				if (char.GetUnicodeCategory(text[x]) == UnicodeCategory.OtherLetter)
				{
					containUnicode = true;
					break;
				}
			}
			return containUnicode;
		}

		public String convertCodeToColor(String input)
		{
			String text = input.Trim();
			if (String.Compare(text, "white") == 0)
			{
				return "白色";
			}
			else if (String.Compare(text, "black") == 0)
			{
				return "黑色";
			}
			else if (String.Compare(text, "red") == 0)
			{
				return "红色";
			}
			else if (String.Compare(text, "yellow") == 0)
			{
				return "黄色";
			}
			else if (String.Compare(text, "blue") == 0)
			{
				return "蓝色";
			}
			else if (String.Compare(text, "green") == 0)
			{
				return "绿色";
			}
			else if (String.Compare(text, "purple") == 0)
			{
				return "紫色";
			}
			else if (String.Compare(text, "orange") == 0)
			{
				return "橙色";
			}
			else if (String.Compare(text, "#B2DEC2") == 0)
			{
				return "护眼绿";
			}
			else if (String.Compare(text, "#365F91") == 0)
			{
				return "深蓝";
			}
			else if (String.Compare(text, "#4F81BD") == 0)
			{
				return "浅蓝";
			}
			else return text;
		}

		public String convertHanToNum_Align(String input)
		{
			switch (input)
			{
				case "居左":
					return "0";
				case "居中":
					return "1";
				case "居右":
					return "2";
				default:
					return "1";
			}
		}

		public String convertNumToHan_Align(String input)
		{
			switch (input)
			{
				case "0":
					return "居左";
				case "1":
					return "居中";
				case "2":
					return "居右";
				default:
					return "居中";
			}
		}

		public String convertHanToNum_Style(String input)
		{
			/*
			1, 2, 3, ...
			１, ２, ３, ...
			一, 二, 三 (简)...
			壹, 貳, 叁 (简)...
			一, 二, 三 (繁)...
			壹, 貳, 叁 (繁)...
			a, b, c, ...
			A, B, C, ...
			i, ii, iii, ...
			I, II, III, ...
			*/
			switch (input)
			{
				case "1, 2, 3, ...":
					return "0";
				case "１, ２, ３, ...":
					return "1";
				case "一, 二, 三 (简)...":
					return "2";
				case "壹, 貳, 叁 (简)...":
					return "3";
				case "一, 二, 三 (繁)...":
					return "4";
				case "壹, 貳, 叁 (繁)...":
					return "5";
				case "a, b, c, ...":
					return "6";
				case "A, B, C, ...":
					return "7";
				case "i, ii, iii, ...":
					return "8";
				case "I, II, III, ...":
					return "9";
				default:
					return "0";
			}
		}

		public String convertNumToHan_Style(String input)
		{
			switch (input)
			{
				case "0":
					return "1, 2, 3, ...";
				case "1":
					return "１, ２, ３, ...";
				case "2":
					return "一, 二, 三 (简)...";
				case "3":
					return "壹, 貳, 叁 (简)...";
				case "4":
					return "一, 二, 三 (繁)...";
				case "5":
					return "壹, 貳, 叁 (繁)...";
				case "6":
					return "a, b, c, ...";
				case "7":
					return "A, B, C, ...";
				case "8":
					return "i, ii, iii, ...";
				case "9":
					return "I, II, III, ...";
				default:
					return "1, 2, 3, ...";
			}
		}

	}
}
