using Ini;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using VB = Microsoft.VisualBasic;
using Word = Microsoft.Office.Interop.Word;
/*
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Diagnostics;
using System.Threading;
using Ini;
using Word = Microsoft.Office.Interop.Word;
using VB = Microsoft.VisualBasic;
*/

namespace TXT2DOC
{
    public partial class Main : Tools
    {
		private Settings settingsForm = new Settings();		// Able to get data from Settings form

		List<String> bookAndAuthor;
		List<int> titleLineNumbers = new List<int>();
		String TXTPath;
		String HTMLPath;
		String HTMLFolderPath;
		String tempPath = Path.Combine(Path.GetTempPath(), "TXT2DOC");
		String iniPath = Path.Combine(Path.GetTempPath(), "TXT2DOC") + "\\settings.ini";		// Application.StartupPath + "\\settings.ini"

		String defaultStatusText = "TXT2DOC 版本: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

		Stopwatch stopWatch;
		int timerCount;

		Word.Application objWord;
		Word.Document objDoc;

        public Main()
        {
            InitializeComponent();
        }

		private void Main_Load(object sender, EventArgs e)
		{
			this.status_label.Text = defaultStatusText;
			Directory.CreateDirectory(tempPath);
			IniFile ini = loadINI(iniPath);
			try
			{
				loadCurSettings(settingsForm, ini);
			}
			catch
			{
				MessageBox.Show("加载设置文件出错，即将导入默认设置！");
				writeDefaultSettings(ini);
				loadCurSettings(settingsForm, ini);
			}
		}

        private void menu1_switchAOT_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
				this.status_label.Text = "已取消置顶";
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
				showBalloonTip("温馨提示", "已取消置顶");
            }
            else
            {
                this.TopMost = true;
				this.status_label.Text = "已置顶";
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
				showBalloonTip("温馨提示", "已置顶");
            }
        }

		private void menu2_settings_Click(object sender, EventArgs e)
		{
			//Settings settingsForm = new Settings();
			settingsForm = new Settings();
			settingsForm.ShowDialog();
		}

		private void menu3_export_Click(object sender, EventArgs e)
		{
			if (TOC_list.Rows[0].Cells[0].Value != null && TOC_list.Rows[0].Cells[1].Value != null)
			{
				this.status_label.Text = "正在导出目录...";

				String TOCPath = getTOCPath(iniPath);
				StreamWriter sw = new StreamWriter(TOCPath, false, Encoding.GetEncoding("GB2312"));

				foreach (DataGridViewRow row in TOC_list.Rows)
				{
					if (row.Cells[0].Value != null && row.Cells[1].Value != null)
					{
						String toPrint = row.Cells[0].Value.ToString() + ">" + row.Cells[1].Value.ToString();
						sw.WriteLine(toPrint);
					}
				}
				sw.Close();

				this.status_label.Text = "目录导出完毕！";
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
				showBalloonTip("温馨提示", "目录导出完毕！\n目录位置：" + TOCPath);
			}
			else
			{
				MessageBox.Show("目录是空的，无法导出！");
			}
		}

		private void menu4_import_Click(object sender, EventArgs e)
		{
			String TOCPath = getTOCPath(iniPath);
			if (File.Exists(TOCPath))
			{
				StreamReader sr = new StreamReader(TOCPath, Encoding.GetEncoding("GB2312"));

				if (TOC_list.Rows[0].Cells[0].Value != null && TOC_list.Rows[0].Cells[1].Value != null)
				{
					DialogResult dialogResult = MessageBox.Show("目录已存在！\n是否先清空目录？", "确认", MessageBoxButtons.YesNo);
					if (dialogResult == DialogResult.Yes)
					{
						TOC_list.Rows.Clear();
					}
				}

				this.status_label.Text = "正在导入目录...";
				stopWatch = new Stopwatch();
				stopWatch.Start();

				// Set cell font colors
				setCellFontColor(Color.Black, Color.RoyalBlue);

				String nextLine;
				String[] separators = { ">" };
				while ((nextLine = sr.ReadLine()) != null)
				{
					String[] data = nextLine.Split(separators, StringSplitOptions.RemoveEmptyEntries);
					DataGridViewRow row = (DataGridViewRow)TOC_list.Rows[0].Clone();
					row.Cells[0].Value = data[0];
					row.Cells[1].Value = data[1];
					TOC_list.Rows.Add(row);
				}
				sr.Close();

				this.status_label.Text = "目录导入完毕！耗时：" + getProcessTime().ToString() + " 秒";
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
				showBalloonTip("温馨提示", "目录导入完毕！");
			}
			else
			{
				MessageBox.Show("目录文件不存在，无法导入目录！");
			}
		}

		private void menu5_clear_Click(object sender, EventArgs e)
		{
			if (TOC_list.Rows[0].Cells[0].Value != null && TOC_list.Rows[0].Cells[1].Value != null)
			{
				DialogResult dialogResult = MessageBox.Show("建议你按导出目录菜单以保存列表框中的数据\n你确定要清空列表框中的数据？", "确认", MessageBoxButtons.YesNo);
				if (dialogResult == DialogResult.Yes)
				{
					this.status_label.Text = "正在清空目录...";

					TOC_list.Rows.Clear();

					this.status_label.Text = "目录清空完毕！";
					notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
					showBalloonTip("温馨提示", "目录清空完毕！");
				}
			}
			else
			{
				MessageBox.Show("目录是空的，无法清空！");
			}
		}

		private void menu6_1_helpfile_menuitem_Click(object sender, EventArgs e)
		{
			try
			{
				System.Diagnostics.Process.Start(@"help.pdf");
			}
			catch
			{
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
				showBalloonTip("错误！", "帮助文件不存在！");
			}
		}

		private void menu6_3_about_menuitem_Click(object sender, EventArgs e)
		{
			About aboutForm = new About();
			aboutForm.ShowDialog();
		}

		private void toostripmenu1_show_Click(object sender, EventArgs e)
		{
			this.WindowState = FormWindowState.Normal;
		}

		private void toostripmenu2_exit_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}

		private void notifyIcon1_DoubleClick(object sender, EventArgs e)
		{
			this.WindowState = FormWindowState.Normal;
		}

		private void TOC_list_DragEnter(object sender, DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
				e.Effect = DragDropEffects.All;
			else
				e.Effect = DragDropEffects.None;
		}

		private void TOC_list_DragDrop(object sender, DragEventArgs e)
		{
			this.status_label.Text = "正在提取目录...";

			String[] test = (String[])e.Data.GetData(DataFormats.FileDrop);
			if (test.Length != 1)
			{
				this.status_label.Text = defaultStatusText;
				MessageBox.Show("只能拖拽一个文件！");
				e.Effect = DragDropEffects.None;
			}
			else if (!test[0].ToLower().EndsWith(".txt"))
			{
				this.status_label.Text = defaultStatusText;
				MessageBox.Show("拖拽进来的不是TXT文件！");
				e.Effect = DragDropEffects.None;
			}
			else
			{
				stopWatch = new Stopwatch();
				stopWatch.Start();

				IniFile ini = loadINI(iniPath);
				if (Convert.ToInt32(ini.IniReadValue("Tab_4", "Drag_Clear_List")) == 1)		//拖入文件时清空列表
				{
					if (TOC_list.Rows[0].Cells[0].Value != null && TOC_list.Rows[0].Cells[1].Value != null)
					{
						TOC_list.Rows.Clear();
					}
				}

				// Save txt file path
				TXTPath = test[0];

				// Set cell font colors
				setCellFontColor(Color.Black, Color.RoyalBlue);

				// Get file name
				String filename = Path.GetFileNameWithoutExtension(TXTPath);

				// Get book name and author info to fill the first two rows of TOC_list
				bookAndAuthor = getBookNameAndAuthorInfo(TXTPath, filename);
				DataGridViewRow tempRow1 = (DataGridViewRow)TOC_list.Rows[0].Clone();
				tempRow1.Cells[0].Value = bookAndAuthor[0];
				tempRow1.Cells[1].Value = "★书名，勿删此行★";
				TOC_list.Rows.Add(tempRow1);
				DataGridViewRow tempRow2 = (DataGridViewRow)TOC_list.Rows[0].Clone();
				tempRow2.Cells[0].Value = bookAndAuthor[1];
				tempRow2.Cells[1].Value = "★作者，勿删此行★";
				TOC_list.Rows.Add(tempRow2);

				// Prepare a list of title line numbers
				StreamReader sr = new StreamReader(TXTPath, Encoding.GetEncoding("GB2312"));
				String nextLine;
				int lineNumber = 1;
				int rowNumber = 2;
				while ((nextLine = sr.ReadLine()) != null)
				{
					//Match title = Regex.Match(nextLine, "^([\\s\t　]{0,20}(正文[\\s\t　]{0,4})?[第【]([——-——一二两三四五六七八九十○零百千0-9０-９]{1,12}).*[章节節回集卷部】].*?$)|(Ui)(第.{1,5}章)|(Ui)(第.{1,5}节)");
					//Match title = Regex.Match(nextLine, "^([\\s\t　]*(正文[\\s\t　]*)?[第【][\\s\t　]*([——-——一二两三四五六七八九十○零百千壹贰叁肆伍陆柒捌玖拾佰仟0-9０-９]*)[\\s\t　]*[章节節回集卷部】][\\s\t　]*.{0,40}?$)|(Ui)(第.{1,5}章)|(Ui)(第.{1,5}节)");
					Match title = Regex.Match(nextLine, "^([\\s\t　]*([【])?(正文[\\s\t　]*)?[第【][\\s\t　]*([——-——一二两三四五六七八九十○零百千壹贰叁肆伍陆柒捌玖拾佰仟0-9０-９]*)[\\s\t　]*[章节節回集卷部】][\\s\t　]*.{0,40}?$)");

					// Chapter title (with its line number) found!
					if (title.Success)
					{
						DataGridViewRow row = (DataGridViewRow)TOC_list.Rows[rowNumber].Clone();
						row.Cells[0].Value = title.ToString().Trim();
						row.Cells[1].Value = lineNumber;
						TOC_list.Rows.Add(row);
						rowNumber++;

						titleLineNumbers.Add(lineNumber);
					}
					lineNumber++;
				}
				sr.Close();

				this.status_label.Text = "目录提取完毕！耗时：" + getProcessTime().ToString() + " 秒";
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
				showBalloonTip("温馨提示", "目录提取完毕！");
			}
		}

		private void generate_button_Click(object sender, EventArgs e)
		{
			if (TXTPath == null)
			{
				MessageBox.Show("没有源文件！请拖入TXT文件后重试！");
				return;
			}

			stopWatch = new Stopwatch();
			stopWatch.Start();
			
			/*** Load new TOC ***/
			loadNewTOC();

			/*** Load settings ***/
			IniFile ini = loadINI(iniPath);
			bool cover = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Cover")));
			bool TOC = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "TOC")));
			bool TOCLoad = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "TOC_Load")));
			bool vertical = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Vertical")));
			bool replace = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Replace")));
			bool StT = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "StT")));
			bool TtS = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "TtS")));
			if (StT && TtS)
			{
				StT = false;
				TtS = false;
			}
			int translation;		// 0: 不转; 1: 简转繁; 2: 繁转简
			if (StT && TtS)
			{
				translation = 0;
			}
			else if (StT && !TtS)
			{
				translation = 1;
				// 书名作者名简转繁
				String temp1 = bookAndAuthor[0];
				String temp2 = bookAndAuthor[1];
				bookAndAuthor.Clear();
				bookAndAuthor.Add(VB.Strings.StrConv(temp1, VB.VbStrConv.TraditionalChinese, 0));
				bookAndAuthor.Add(VB.Strings.StrConv(temp2, VB.VbStrConv.TraditionalChinese, 0));
			}
			else if (!StT && TtS)
			{
				translation = 2;
				// 书名作者名繁转简
				String temp1 = bookAndAuthor[0];
				String temp2 = bookAndAuthor[1];
				bookAndAuthor.Clear();
				bookAndAuthor.Add(VB.Strings.StrConv(temp1, VB.VbStrConv.SimplifiedChinese, 0));
				bookAndAuthor.Add(VB.Strings.StrConv(temp2, VB.VbStrConv.SimplifiedChinese, 0));
			}
			else translation = 0;
			//bool dropCap = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_1", "Drop_Cap")));
			float pageWidth = float.Parse(ini.IniReadValue("Tab_2", "Page_Width"));
			float pageHeight = float.Parse(ini.IniReadValue("Tab_2", "Page_Height")); ;
			String pageColor = ini.IniReadValue("Tab_2", "Page_Color"); ;
			float pageMarginU = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Up")); ;
			float pageMarginD = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Down")); ;
			float pageMarginL = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Left")); ;
			float pageMarginR = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Right")); ;
			float headerMargin = float.Parse(ini.IniReadValue("Tab_2", "Header_Margin")); ;
			float footerMargin = float.Parse(ini.IniReadValue("Tab_2", "Footer_Margin")); ;
			String titleFont = ini.IniReadValue("Tab_3", "Title_Font"); ;
			float titleSize = float.Parse(ini.IniReadValue("Tab_3", "Title_Size")); ;
			String titleColor = ini.IniReadValue("Tab_3", "Title_Color"); ;
			String bodyFont = ini.IniReadValue("Tab_3", "Body_Font"); ;
			float bodySize = float.Parse(ini.IniReadValue("Tab_3", "Body_Size")); ;
			String bodyColor = ini.IniReadValue("Tab_3", "Body_Color"); ;
			float lineSpacing = float.Parse(ini.IniReadValue("Tab_3", "Line_Spacing")); ;
			bool addParagraphSpacing = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_3", "Add_Paragraph_Spacing")));
			String fileLocation = ini.IniReadValue("Tab_4", "Generated_File_Location"); ;
			bool deleteTempFiles = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "Delete_Temp_Files")));
			bool pdf = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "PDF")));
			bool STMode = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_4", "Save_Time_Mode")));
			if (pdf && STMode)
			{
				pdf = false;
			}
			bool header = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Header")));
			float headerSize = 0.0F;
			String headerAlign = "";
			bool headerBorder = false;
			if (header)
			{
				headerSize = float.Parse(ini.IniReadValue("Tab_5", "Header_Size"));
				headerAlign = ini.IniReadValue("Tab_5", "Header_Align");
				headerBorder = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Header_Border")));
			}
			if (!header && headerBorder)
			{
				headerBorder = false;
			}
			bool footer = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Footer")));
			float footerSize = 0.0F;
			String footerAlign = "";
			String footerStyle = "";
			bool footerBorder = false;
			if (footer)
			{
				footerSize = float.Parse(ini.IniReadValue("Tab_5", "Footer_Size"));
				footerAlign = ini.IniReadValue("Tab_5", "Footer_Align");
				footerStyle = ini.IniReadValue("Tab_5", "Footer_Style");
				footerBorder = Convert.ToBoolean(Convert.ToInt32(ini.IniReadValue("Tab_5", "Footer_Border")));
			}
			if (!footer && footerBorder)
			{
				footerBorder = false;
			}
			HTMLPath = getHTMLPath(TXTPath);
			HTMLFolderPath = getHTMLFolderPath(TXTPath);

			/*** Generate temp HTML ***/
			generateHTML(translation, vertical, cover, TOC, replace, pageWidth, pageHeight, pageMarginL, pageMarginR, pageMarginU, pageMarginD, headerMargin, footerMargin, lineSpacing, titleFont, titleSize, titleColor, bodyFont, bodySize, bodyColor);

			/*** Open HTML via Word ***/
			objWord = new Word.Application();
			objDoc = objWord.Documents.OpenNoRepairDialog(HTMLPath);
			processHTMLInWord(translation, vertical, cover, TOC, TOCLoad, pageColor, titleFont, titleSize, titleColor, bodyFont, bodySize, bodyColor, lineSpacing, addParagraphSpacing, header, headerSize, headerAlign, headerBorder, footer, footerSize, footerAlign, footerStyle, footerBorder);
			//objWord.ActiveDocument.Range(0, 0).TCSCConverter(Word.WdTCSCConverterDirection.wdTCSCConverterDirectionAuto);
			//Word.Range rng = objWord.ActiveDocument.Sections[3].Range;
			//rng.Font.Name = "微软雅黑";

			/*** Save Time Mode? ***/
			if (STMode)
			{
				this.status_label.Text = "生成完毕！耗时：" + getProcessTime().ToString() + " 秒";
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
				showBalloonTip("温馨提示", "生成完毕！");

				objWord.Visible = true;
				return;
			}

			/*** Save HTML as DOCX or PDF ***/
			String DocName = "《" + bookAndAuthor[0] + "》作者：" + bookAndAuthor[1];
			object filename = fileLocation + "\\" + DocName;
			String ext = wordSaveHTML(DocName, filename, pdf);
			objWord.Quit();

			this.status_label.Text = "生成完毕！文件：" + DocName + " ；耗时：" + getProcessTime().ToString() + " 秒";
			notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
			showBalloonTip("温馨提示", filename.ToString() + ext + "\n已生成完毕！");

			DialogResult dialogResult = MessageBox.Show("生成完毕！\n是否调用Word打开\n\"" + filename.ToString() + ext + "\"", "确认", MessageBoxButtons.YesNo);
			if (dialogResult == DialogResult.Yes)
			{
				try
				{
					System.Diagnostics.Process.Start(filename.ToString() + ext);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
			}

			/*** Delete temp files ***/
			if (deleteTempFiles)
			{
				deletTempHTMLFile();
			}
		}

		private void status_label_TextChanged(object sender, EventArgs e)
		{
			if (!this.status_label.Text.Contains("正在"))
			{
				timer.Enabled = true;
				timerCount = 15;
			}
		}

		private void timer_Tick(object sender, EventArgs e)
		{
			timerCount--;
			if (timerCount == 0) 
			{
				this.status_label.Text = defaultStatusText;
				timer.Stop();
			}
		}

		/*** Helper functions ***/
		private void loadNewTOC()
		{
			this.status_label.Text = "正在生成... 正在加载设置...";

			if (TOC_list.Rows[0].Cells[0].Value != null && TOC_list.Rows[1].Cells[0].Value != null)
			{
				// load new book name and author
				bookAndAuthor.Clear();
				bookAndAuthor.Add(TOC_list.Rows[0].Cells[0].Value.ToString());
				bookAndAuthor.Add(TOC_list.Rows[1].Cells[0].Value.ToString());

				// load new title line number
				if (TOC_list.Rows[2].Cells[0].Value != null && TOC_list.Rows[2].Cells[0].Value != null)
				{
					titleLineNumbers.Clear();
					for (int i = 2; i < TOC_list.Rows.Count; i++)
					{
						if (TOC_list.Rows[i].Cells[0].Value != null && TOC_list.Rows[i].Cells[1].Value != null)
						{
							titleLineNumbers.Add(Convert.ToInt32(TOC_list.Rows[i].Cells[1].Value));
						}
					}
				}
			}
		}

		/*
		private void loadNewSettings(IniFile ini)
		{

			// Tab_1
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "Cover")) == 1) cover = true;
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "TOC")) == 1) TOC = true;
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "TOC_Load")) == 1) TOCLoad = true;
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "Vertical")) == 1) vertical = true;
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "Replace")) == 1) replace = true;
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "PDF")) == 1) pdf = true;
			if (Convert.ToInt32(ini.IniReadValue("Tab_1", "Save_Time_Mode")) == 1) STMode = true;
			MessageBox.Show(STMode.ToString());
			if (pdf && STMode)
			{
				pdf = false;
			}

			// Tab_2
			pageWidth = float.Parse(ini.IniReadValue("Tab_2", "Page_Width"));
			pageHeight = float.Parse(ini.IniReadValue("Tab_2", "Page_Height"));
			pageColor = ini.IniReadValue("Tab_2", "Page_Color");
			pageMarginU = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Up"));
			pageMarginD = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Down"));
			pageMarginL = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Left"));
			pageMarginR = float.Parse(ini.IniReadValue("Tab_2", "Page_Margin_Right"));
			headerMargin = float.Parse(ini.IniReadValue("Tab_2", "Header_Margin"));
			footerMargin = float.Parse(ini.IniReadValue("Tab_2", "Footer_Margin"));

			// Tab_3
			titleFont = ini.IniReadValue("Tab_3", "Title_Font");
			titleSize = float.Parse(ini.IniReadValue("Tab_3", "Title_Size"));
			titleColor = ini.IniReadValue("Tab_3", "Title_Color");
			bodyFont = ini.IniReadValue("Tab_3", "Body_Font");
			bodySize = float.Parse(ini.IniReadValue("Tab_3", "Body_Size"));
			bodyColor = ini.IniReadValue("Tab_3", "Body_Color");
			lineSpacing = float.Parse(ini.IniReadValue("Tab_3", "Line_Spacing"));
			if (Convert.ToInt32(ini.IniReadValue("Tab_3", "Add_Paragraph_Spacing")) == 1) addParagraphSpacing = true;

			// Tab_4
			fileLocation = ini.IniReadValue("Tab_4", "Generated_File_Location");
			if (Convert.ToInt32(ini.IniReadValue("Tab_4", "Delete_Temp_Files")) == 1) deleteTempFiles = true;

			// Tab_5
			if (Convert.ToInt32(ini.IniReadValue("Tab_5", "Header")) == 1) header = true;
			if (header)
			{
				headerSize = float.Parse(ini.IniReadValue("Tab_5", "Header_Size"));
				headerAlign = ini.IniReadValue("Tab_5", "Header_Align");
				if (Convert.ToInt32(ini.IniReadValue("Tab_5", "Header_Border")) == 1) headerBorder = true;
			}
			if (!header && headerBorder)
			{
				headerBorder = false;
			}
			if (Convert.ToInt32(ini.IniReadValue("Tab_5", "Footer")) == 1) footer = true;
			if (footer)
			{
				footerSize = float.Parse(ini.IniReadValue("Tab_5", "Footer_Size"));
				footerAlign = ini.IniReadValue("Tab_5", "Footer_Align");
				footerStyle = ini.IniReadValue("Tab_5", "Footer_Style");
				if (Convert.ToInt32(ini.IniReadValue("Tab_5", "Footer_Border")) == 1) footerBorder = true;
			}
			if (!footer && footerBorder)
			{
				footerBorder = false;
			}

			HTMLPath = getHTMLPath(TXTPath);
			HTMLFolderPath = getHTMLFolderPath(TXTPath);
		}
		*/

		private void generateHTML(int translation, bool vertical, bool cover, bool TOC, bool replace, float pageWidth, float pageHeight, float pageMarginL, float pageMarginR, float pageMarginU, float pageMarginD, float headerMargin, float footerMargin, float lineSpacing, String titleFont, float titleSize, String titleColor, String bodyFont, float bodySize, String bodyColor)
		{
			this.status_label.Text = "正在生成... 正在生成临时HTML文件...";

			StreamReader sr = new StreamReader(TXTPath, Encoding.GetEncoding("GB2312"));
			StreamWriter sw = new StreamWriter(HTMLPath, false, Encoding.GetEncoding("GB2312"));

			// Generate CSS style
			String CSS_vertical = "";
			if (vertical)
			{
				CSS_vertical = "layout-flow:vertical-ideographic;";
			}

			String pageSettings = "size:" + pageWidth + "in " + pageHeight + "in;margin:" + pageMarginU + "in " + pageMarginR + "in " + pageMarginD + "in " + pageMarginL + "in;mso-header-margin:" + headerMargin + "in;mso-footer-margin:" + footerMargin + "in;";
			sw.WriteLine("<html>\n<style>\np{text-indent:" + bodySize * 2 + "pt;margin-top:0em;margin-bottom:0em}\n@page WordSection1{" + pageSettings + CSS_vertical + "}div.WordSection1{page:WordSection1;}\n@page WordSection2{" + pageSettings + CSS_vertical + "}div.WordSection2{page:WordSection2;}\n@page WordSection3{" + pageSettings + CSS_vertical + "}div.WordSection3{page:WordSection3;}\n</style>\n<body>");

			// Generate cover
			if (cover)
			{
				sw.WriteLine("<div class=WordSection1>");
				String coverHTML = "";
				if (vertical)
				{
					float tableWidth = (pageWidth - pageMarginL - pageMarginR) * 72 - 8;
					float tableHeight = (pageHeight - pageMarginU - pageMarginD) * 72;
					coverHTML = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=\"" + tableWidth + "pt\" style='width:" + tableWidth + "pt;mso-table-anchor-vertical:margin;mso-table-anchor-horizontal:column;mso-table-left:left;mso-table-top:top'>\n	<td width=\"38%\" style='width:38.0%;mso-border-right-alt:solid #4F81BD .5pt;layout-flow:vertical-ideographic;height:" + tableHeight + "pt'>\n	<p class=MsoNoSpacing align=center style='text-align:center;text-indent:0pt;mso-element-frame-hspace:2.25pt;mso-element-wrap:around;mso-element-anchor-horizontal:column;mso-element-left:left;mso-element-top:top;mso-height-rule:exactly'><span style='font-size:26.0pt;font-family:" + bodyFont + ";color:" + bodyColor + "'>作者：" + bookAndAuthor[1] + "</span></p>\n	</td>\n	<td width=\"62%\" style='width:62.0%;layout-flow:vertical-ideographic;height:" + tableHeight + "pt'>\n	<p class=MsoNoSpacing align=center style='text-align:center;text-indent:0pt;mso-element-frame-hspace:2.25pt;mso-element-wrap:around;mso-element-anchor-horizontal:column;mso-element-left:left;mso-element-top:top;mso-height-rule:exactly'><span style='font-size:72.0pt;font-family:" + titleFont + ";color:" + titleColor + "'><b>" + bookAndAuthor[0] + "</b></span></p>\n	</td>\n</table>";
				}
				else
				{
					float tableTop = (pageHeight - pageMarginU - pageMarginD) * 72 / 4;
					coverHTML = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=\"100%\" style='width:100.0%;mso-table-top:" + tableTop + "pt'>\n	<td width=\"100%\" style='width:100.0%;mso-border-bottom-alt:solid #4F81BD .5pt'>\n	<p class=MsoNoSpacing align=center style='text-align:center;text-indent:0pt;mso-element-top:" + tableTop + "pt;text-indent:0pt'><span style='font-size:64.0pt;font-family:" + titleFont + ";color:" + titleColor + "'><b>" + bookAndAuthor[0] + "</b></span></p>\n	</td>\n</tr>\n	<td width=\"100%\" style='width:100.0%'>\n	<p class=MsoNoSpacing align=right style='text-align:right;text-indent:0pt;mso-element-top:" + tableTop + "pt'><span style='font-size:26.0pt;font-family:" + bodyFont + ";color:" + bodyColor + "'>作者：" + bookAndAuthor[1] + "</span></p>\n	</td>\n</tr>\n</table>";
				}
				sw.WriteLine(coverHTML);
				sw.WriteLine("</div>\n<br clear=all style='page-break-before:always;mso-break-type:section-break'>");
			}

			// Generate TOC
			if (TOC)
			{
				String TOCText = "目录";
				if (translation == 1) TOCText = "目錄";
				sw.WriteLine("<div class=WordSection2>");
				sw.WriteLine("<p class=MsoNormal style=\"font-size:" + titleSize + "pt;font-family:" + titleFont + ";color:" + titleColor + ";text-align:center;line-height:" + lineSpacing * 100 + "%;font-weight:bold;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto\">" + TOCText + "</p>\n<w:Sdt SdtDocPart=\"t\" DocPartType=\"Table of Contents\" DocPartUnique=\"t\">\n<p class=MsoNormal>\n<!--[if supportFields]>\n<span style='mso-element:field-begin'></span>\n<span style='mso-spacerun:yes'>&nbsp;</span>\nTOC \\o &quot;1-3&quot; \\h \\z \\u \n<span style='mso-element:field-separator'></span>\n<![endif]-->\n<span style='font-family:" + bodyFont + ";font-size:" + bodySize + "pt;color:" + bodyColor + ";font-weight:normal;mso-no-proof:yes'>此处为目录，请按右键并更新域。</span>\n<!--[if supportFields]>\n<span style='mso-element:field-end'></span>\n<![endif]-->\n</p>\n</w:Sdt>");
				sw.WriteLine("</div>\n<br clear=all style='page-break-before:always;mso-break-type:section-break'>");
			}

			sw.WriteLine("<div class=WordSection3>");
			String nextLine;
			int TXTlineNumber = 1;
			int TLN_idx = 0;
			int TLN_size = titleLineNumbers.Count;

			var vbTranslation = VB.VbStrConv.None;
			switch (translation)
			{
				case 0:
					vbTranslation = VB.VbStrConv.None;
					break;
				case 1:
					vbTranslation = VB.VbStrConv.TraditionalChinese;
					break;
				case 2:
					vbTranslation = VB.VbStrConv.SimplifiedChinese;
					break;
				default:
					vbTranslation = VB.VbStrConv.None;
					break;
			}

			while ((nextLine = sr.ReadLine()) != null)
			{
				Match emptyLine = Regex.Match(nextLine, "^\\s*$");
				if (!emptyLine.Success)		// Remove empty lines
				{
					nextLine = nextLine.Trim();
					nextLine = VB.Strings.StrConv(nextLine, vbTranslation, 0);		// 简繁转换

					if (vertical)		// 半角字符转全角
					{
						nextLine = ToSBC(nextLine);
						//nextLine = VB.Strings.StrConv(nextLine, VB.VbStrConv.Wide);
						nextLine = nextLine.Replace("—", "<span lang=EN-US style='font-family:\"Times New Roman\"'>—</span>");
					}

					if (TLN_idx < TLN_size && TXTlineNumber == titleLineNumbers[TLN_idx])		// Chapter titles!
					{
						if (replace)		// 替换标题中的数字为汉字
						{
							nextLine = numberToHan(nextLine);
						}
						sw.WriteLine("<h1>" + nextLine + "</h1>");
						TLN_idx++;
					}
					else
					{
						sw.WriteLine("<p>" + nextLine + "</p>");
					}
				}
				TXTlineNumber++;
			}
			sw.WriteLine("</div>\n</body>\n</html>");

			sr.Close();
			sw.Close();
		}

		private void processHTMLInWord(int translation, bool vertical, bool cover, bool TOC, bool TOCLoad, String pageColor, String titleFont, float titleSize, String titleColor, String bodyFont, float bodySize, String bodyColor, float lineSpacing, bool addParagraphSpacing, bool header, float headerSize, String headerAlign, bool headerBorder, bool footer, float footerSize, String footerAlign, String footerStyle, bool footerBorder)
		{
			this.status_label.Text = "正在生成... 正在处理HTML文件...";

			objWord.Visible = false;
			objWord.ScreenUpdating = false;

			// 简繁转换 （moved to HTML generation section)
			//objWord.ActiveDocument.Range(0, 0).TCSCConverter(Word.WdTCSCConverterDirection.wdTCSCConverterDirectionAuto);
			
			// Set page (moved to HTML word section settings)
			//wordSetPage(objWord, objDoc, ini);

			// Set Heading1 format
			wordSetHeading1(titleFont, titleSize, titleColor);

			// Set body format
			wordSetNormal(bodyFont, bodySize, bodyColor, lineSpacing, addParagraphSpacing, vertical);

			// Set page color
			wordSetPageColor(pageColor);

			// Generate Header
			if (header)
			{
				generateHeader(translation, cover, TOC, headerSize, headerAlign, headerBorder);
			}

			// Generate Footer
			if (footer)
			{
				generateFooter(cover, TOC, footerSize, footerAlign, footerStyle, footerBorder);
			}

			// Update cover format
			if (cover)
			{
				objWord.ActiveDocument.Sections[1].Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
			}

			// Update TOC
			if (TOCLoad)
			{
				this.status_label.Text = "正在生成... 正在加载目录...";
				objWord.ActiveDocument.TablesOfContents[1].UpdatePageNumbers();
			}

			/*
			objWord.ActiveDocument.Sections[3].Range.DetectLanguage();
			if (objWord.ActiveDocument.Sections[3].Range.LanguageID == Word.WdLanguageID.wdSimplifiedChinese) MessageBox.Show("简体中文！");
			else if (objWord.ActiveDocument.Sections[3].Range.LanguageID == Word.WdLanguageID.wdTraditionalChinese) MessageBox.Show("繁体中文！");
			else MessageBox.Show("靠！");
			*/

		}

		/*
		private void wordSetPage(float pageWidth, float pageHeight, float pageMarginL, float pageMarginR, float pageMarginU, float pageMarginD, float headerMargin, float footerMargin)
		{
			objDoc.PageSetup.PageWidth = objWord.InchesToPoints(pageWidth);
			objDoc.PageSetup.PageHeight = objWord.InchesToPoints(pageHeight);
			objDoc.PageSetup.LeftMargin = objWord.InchesToPoints(pageMarginL);
			objDoc.PageSetup.RightMargin = objWord.InchesToPoints(pageMarginR);
			objDoc.PageSetup.TopMargin = objWord.InchesToPoints(pageMarginU);
			objDoc.PageSetup.BottomMargin = objWord.InchesToPoints(pageMarginD);
			objDoc.PageSetup.HeaderDistance = objWord.InchesToPoints(headerMargin);
			objDoc.PageSetup.FooterDistance = objWord.InchesToPoints(footerMargin);
		}
		*/

		private void wordSetHeading1(String titleFont, float titleSize, String titleColor)
		{
			object heading1;
			if (objWord.Language == Microsoft.Office.Core.MsoLanguageID.msoLanguageIDSimplifiedChinese)
				heading1 = "标题 1";
			else if (objWord.Language == Microsoft.Office.Core.MsoLanguageID.msoLanguageIDTraditionalChinese)
				heading1 = "標題 1";
			else heading1 = "Heading 1";
			objWord.ActiveDocument.Styles[-2].ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
			objWord.ActiveDocument.Styles[-2].Font.Name = titleFont;
			objWord.ActiveDocument.Styles[-2].Font.Size = titleSize;
			objWord.ActiveDocument.Styles[-2].Font.Color = convertHTMLColorToWdColor(titleColor, 1);
			objWord.ActiveDocument.Styles[-2].ParagraphFormat.PageBreakBefore = -1;		// Define page break INSIDE heading to prevent unnecessary blank pages!
			//objWord.ActiveDocument.Styles[-2].ParagraphFormat.IndentFirstLineCharWidth(0);
			//objWord.ActiveDocument.Styles[-2].NoSpaceBetweenParagraphsOfSameStyle = true;
		}

		private void wordSetNormal(String bodyFont, float bodySize, String bodyColor, float lineSpacing, bool addParagraphSpacing, bool vertical)
		{
			object normal;
			if (objWord.Language == Microsoft.Office.Core.MsoLanguageID.msoLanguageIDSimplifiedChinese)
				normal = "正文";
			else if (objWord.Language == Microsoft.Office.Core.MsoLanguageID.msoLanguageIDTraditionalChinese)
				normal = "正文";
			else normal = "Normal";
			objWord.ActiveDocument.Styles[-1].Font.Name = bodyFont;
			objWord.ActiveDocument.Styles[-1].Font.Size = bodySize;
			objWord.ActiveDocument.Styles[-1].Font.Color = convertHTMLColorToWdColor(bodyColor, 2);
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.LineSpacing = objWord.LinesToPoints(lineSpacing);
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.SpaceBefore = 0.0F;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.SpaceAfter = 0.0F;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.WidowControl = 0;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.KeepTogether = 0;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.KeepWithNext = 0;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.Hyphenation = 0;
			objWord.ActiveDocument.Styles[-1].ParagraphFormat.HangingPunctuation = 0;
			if (addParagraphSpacing)
			{
				objWord.ActiveDocument.Styles[-1].ParagraphFormat.SpaceBeforeAuto = -1;		// True is -1; False is 0
				objWord.ActiveDocument.Styles[-1].ParagraphFormat.SpaceAfterAuto = -1;
			}
			if (vertical)
			{
				objWord.Selection.Orientation = Word.WdTextOrientation.wdTextOrientationVerticalFarEast;
				//objWord.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
			}
			//objWord.ActiveDocument.Styles[-1].ParagraphFormat.IndentFirstLineCharWidth(2);
			//objWord.ActiveDocument.Styles[-1].NoSpaceBetweenParagraphsOfSameStyle = true;
		}

		private void wordSetPageColor(String pageColor)
		{
			if (String.Compare(pageColor, "羊皮纸") == 0)
			{
				objWord.ActiveDocument.Background.Fill.PresetTextured(Microsoft.Office.Core.MsoPresetTexture.msoTextureParchment);
			}
			/*
			else if (String.Compare(pageColor, "浅蓝网格") == 0)
			{
				objWord.ActiveDocument.Background.Fill.UserPicture(Application.StartupPath + "\\qianlanwangge.png");
			}
			*/
			else
			{
				objWord.ActiveDocument.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(convertHTMLColorToDrawColor(pageColor, 0));
				objWord.ActiveDocument.Background.Fill.Solid();
			}
		}

		private void generateHeader(int translation, bool cover, bool TOC, float headerSize, String headerAlign, bool headerBorder)
		{
			this.status_label.Text = "正在生成... 正在生成页眉...";

			if (cover && TOC)
			{
				setHeader(translation, 2, headerSize, headerAlign, headerBorder, 2);
				setHeader(translation, 3, headerSize, headerAlign, headerBorder, 2);
			}
			else if (!cover && TOC)
			{
				setHeader(translation, 1, headerSize, headerAlign, headerBorder, 1);
				setHeader(translation, 2, headerSize, headerAlign, headerBorder, 1);
			}
			else if (cover && !TOC)
			{
				setHeader(translation, 2, headerSize, headerAlign, headerBorder, 0);
			}
			else
			{
				setHeader(translation, 1, headerSize, headerAlign, headerBorder, 0);
			}
		}

		private void setHeader(int translation, int sectionNumber, float headerSize, String headerAlign, bool headerBorder, int TOC_sectionNumber)
		{
			Word.HeaderFooter header = objWord.ActiveDocument.Sections[sectionNumber].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
			header.LinkToPrevious = false;
			Word.Range headerRange = header.Range;

			if (sectionNumber == TOC_sectionNumber)
			{
				headerRange.Text = "目录";
				if (translation == 1) headerRange.Text = "目錄";
			}
			else
			{
				String header_text;
				if (objWord.Language == Microsoft.Office.Core.MsoLanguageID.msoLanguageIDSimplifiedChinese)
					header_text = "标题 1";
				else if (objWord.Language == Microsoft.Office.Core.MsoLanguageID.msoLanguageIDTraditionalChinese)
					header_text = "標題 1";
				else header_text = "Heading 1";
				headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldStyleRef, "\"" + header_text + "\"");
			}
			headerRange.Font.Size = headerSize;
			headerRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

			if (headerBorder)
			{
				headerRange.ParagraphFormat.Borders.Enable = -1;
				headerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
				headerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth025pt;
				headerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].Visible = true;
				headerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
				headerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].Visible = false;
				headerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].Visible = false;
			}

			switch (headerAlign)
			{
				case "0":
					headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
					break;
				case "1":
					headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
					break;
				case "2":
					headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
					break;
				default:
					headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
					break;
			}
		}

		private void generateFooter(bool cover, bool TOC, float footerSize, String footerAlign, String footerStyle, bool footerBorder)
		{
			this.status_label.Text = "正在生成... 正在生成页脚...";

			if (cover && TOC)
			{
				setPageNumber(2, footerSize, footerAlign, footerStyle, footerBorder);
				setPageNumber(3, footerSize, footerAlign, footerStyle, footerBorder);
			}
			else if (!cover && TOC)
			{
				setPageNumber(1, footerSize, footerAlign, footerStyle, footerBorder);
				setPageNumber(2, footerSize, footerAlign, footerStyle, footerBorder);
			}
			else if (cover && !TOC)
			{
				setPageNumber(2, footerSize, footerAlign, footerStyle, footerBorder);
			}
			else
			{
				setPageNumber(1, footerSize, footerAlign, footerStyle, footerBorder);
			}
		}

		private void setPageNumber(int sectionNumber, float footerSize, String footerAlign, String footerStyle, bool footerBorder)
		{
			/*
			objWord.ActiveDocument.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = HFSize;
			objWord.ActiveDocument.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
			objWord.ActiveDocument.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
			objWord.ActiveDocument.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = 1;
			objWord.ActiveDocument.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter, true);
			*/
			//objWord.ActiveDocument.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleSimpChinNum1;

			Word.HeaderFooter PN = objWord.ActiveDocument.Sections[sectionNumber].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
			PN.LinkToPrevious = false;
			PN.PageNumbers.RestartNumberingAtSection = true;
			PN.PageNumbers.StartingNumber = 1;
			Word.Range footerRange = PN.Range;
			footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
			footerRange.Font.Size = footerSize;
			footerRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

			if (footerBorder)
			{
				footerRange.ParagraphFormat.Borders.Enable = -1;
				footerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
				footerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth025pt;
				footerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].Visible = true;
				footerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
				footerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].Visible = false;
				footerRange.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].Visible = false;
			}

			switch (footerAlign)
			{
				case "0":
					footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
					break;
				case "1":
					footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
					break;
				case "2":
					footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
					break;
				default:
					footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
					break;
			}

			switch (footerStyle)
			{
				case "0":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleArabic;
					break;
				case "1":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleArabicFullWidth;
					break;
				case "2":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleSimpChinNum1;
					break;
				case "3":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleSimpChinNum2;
					break;
				case "4":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleTradChinNum1;
					break;
				case "5":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleTradChinNum2;
					break;
				case "6":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleLowercaseLetter;
					break;
				case "7":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleUppercaseLetter;
					break;
				case "8":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleLowercaseRoman;
					break;
				case "9":
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
					break;
				default:
					PN.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleArabic;
					break;
			}
		}

		private String wordSaveHTML(String DocName, object filename, bool pdf)
		{
			String ext = "";

			try
			{
				this.status_label.Text = "正在生成... 正在保存DOC文件...";

				objWord.ActiveDocument.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyTitle] = bookAndAuthor[0];
				objWord.ActiveDocument.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyAuthor] = bookAndAuthor[1];

				//objDoc.SaveAs2(filename, 16, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
				//objDoc.SaveAs2(filename, 0, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

				if (pdf)
				{
					ext = ".pdf";
					objDoc.SaveAs(filename, 0);
					this.status_label.Text = "正在生成... 正在保存PDF文件...";
					objDoc.SaveAs(filename, Word.WdSaveFormat.wdFormatPDF);
					objDoc.Save();
				}
				else
				{
					ext = ".doc";
					objDoc.SaveAs2(filename, 0);		// save as doc instead of docx is because docx automatically suppresses space after a hard page break
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

			return ext;
		}

		private void deletTempHTMLFile()
		{
			if (File.Exists(HTMLPath))
			{
				File.Delete(HTMLPath);
			}
			if (Directory.Exists(HTMLFolderPath))
			{
				Directory.Delete(HTMLFolderPath, true);
			}
		}

		private double getProcessTime()
		{
			stopWatch.Stop();
			long duration = stopWatch.ElapsedMilliseconds;
			return (duration / 1000d);
		}

		private void showBalloonTip(String title, String text)
		{
			notifyIcon1.BalloonTipTitle = title;
			notifyIcon1.BalloonTipText = text;
			notifyIcon1.ShowBalloonTip(1000);
		}

		private String getTOCPath(String iniPath)
		{
			IniFile ini = loadINI(iniPath);
			return ini.IniReadValue("Tab_4", "Generated_File_Location") + "\\目录.txt";
		}

		private String getHTMLPath(String TXTPath)
		{
			IniFile ini = loadINI(iniPath);
			return ini.IniReadValue("Tab_4", "Generated_File_Location") + "\\" + Path.GetFileNameWithoutExtension(TXTPath) + ".html";
		}

		private String getHTMLFolderPath(String TXTPath)
		{
			IniFile ini = loadINI(iniPath);
			return ini.IniReadValue("Tab_4", "Generated_File_Location") + "\\" + Path.GetFileNameWithoutExtension(TXTPath) + "_files";
		}

		private void setCellFontColor(Color a, Color b)
		{
			TOC_list.Columns[0].DefaultCellStyle.ForeColor = a;
			TOC_list.Columns[1].DefaultCellStyle.ForeColor = b;
		}

		private List<String> getBookNameAndAuthorInfo(String path, String filename)
		{
			List<String> result = new List<String>();

			filename = filename.Trim();
			// 1. get info from file name
			int pos = filename.IndexOf("作者：");
			if (pos == -1)
				pos = filename.IndexOf("作者:");
			if (pos != -1)
			{
				String bookname = filename.Substring(0, pos);
				char[] charsToTrim = { '《', '》' };
				bookname = bookname.Trim(charsToTrim);
				bookname = bookname.Replace("书名：", "");
				bookname = bookname.Replace("书名:", "");
				bookname = bookname.Trim();
				String author = filename.Substring(filename.IndexOf("作者：") + 3, filename.Length - filename.IndexOf("作者：") - 3);
				author = author.Trim();

				result.Add(bookname);
				result.Add(author);

				return result;
			}
			else
			{
				// No complete book name and author info
				notifyIcon1.BalloonTipIcon = ToolTipIcon.Warning;
				showBalloonTip("文件名中无完整书名和作者信息！", "完整信息范例：“《书名》作者：作者名”`n（文件名中一定要有“作者”二字，书名号可以省略）`n`n开始检测文件第一、二行……`n其中第一行为书名，第二行为作者名。！");

				// Treat first line as book name and second line as author
				StreamReader sr = new StreamReader(path, Encoding.GetEncoding("GB2312"));
				String bookname = sr.ReadLine();
				if (bookname == null) bookname = filename;
				else
				{
					char[] charsToTrim = { '《', '》' };
					bookname = bookname.Trim(charsToTrim);
					bookname = bookname.Replace("书名：", "");
					bookname = bookname.Replace("书名:", "");
					bookname = bookname.Trim();
				}

				String author = sr.ReadLine();
				if (author == null) author = "东皇";
				else
				{
					author = author.Replace("作者：", "");
					author = author.Replace("作者:", "");
					author = author.Trim();
				}

				result.Add(bookname);
				result.Add(author);

				return result;
			}
		}

		private static String numberToHan(String nextLine)
		{
			nextLine = nextLine.Replace("0", "零");
			nextLine = nextLine.Replace("1", "一");
			nextLine = nextLine.Replace("2", "二");
			nextLine = nextLine.Replace("3", "三");
			nextLine = nextLine.Replace("4", "四");
			nextLine = nextLine.Replace("5", "五");
			nextLine = nextLine.Replace("6", "六");
			nextLine = nextLine.Replace("7", "七");
			nextLine = nextLine.Replace("8", "八");
			nextLine = nextLine.Replace("9", "九");
			nextLine = nextLine.Replace("０", "零");
			nextLine = nextLine.Replace("１", "一");
			nextLine = nextLine.Replace("２", "二");
			nextLine = nextLine.Replace("３", "三");
			nextLine = nextLine.Replace("４", "四");
			nextLine = nextLine.Replace("５", "五");
			nextLine = nextLine.Replace("６", "六");
			nextLine = nextLine.Replace("７", "七");
			nextLine = nextLine.Replace("８", "八");
			nextLine = nextLine.Replace("９", "九");
			return nextLine;
		}

		private System.Drawing.Color convertHTMLColorToDrawColor(String HTMLColor, int flag)
		{
			System.Drawing.Color color = new Color();
			try
			{
				color = ColorTranslator.FromHtml(HTMLColor);
			}
			catch
			{

				IniFile ini = loadINI(iniPath);
				if (flag == 0)
				{
					color = ColorTranslator.FromHtml("white");
					ini.IniWriteValue("Tab_2", "Page_Color", "white");
				}
				else if (flag == 1)
				{
					color = ColorTranslator.FromHtml("black");
					ini.IniWriteValue("Tab_3", "Title_Color", "black");
				}
				else if (flag == 2)
				{
					color = ColorTranslator.FromHtml("black");
					ini.IniWriteValue("Tab_3", "Body_Color", "black");
				}
				else
					MessageBox.Show("Wrong position flag!");
			}
			return color;
		}

		private Word.WdColor convertHTMLColorToWdColor(String HTMLColor, int flag)		// flag == 0: page color; flag == 1: title color; flag == 2: body color
		{
			System.Drawing.Color color = convertHTMLColorToDrawColor(HTMLColor, flag);
			int rgbColor = VB.Information.RGB(color.R, color.G, color.B);
			Word.WdColor wdColor = (Word.WdColor)rgbColor;
			return wdColor;
		}

		private static String ToSBC(String input)
		{
			// 半角转全角：
			char[] c = input.ToCharArray();
			for (int i = 0; i < c.Length; i++)
			{
				if (c[i] == 32)
				{
					c[i] = (char)12288;
					continue;
				}
				if (c[i] < 127)
					c[i] = (char)(c[i] + 65248);
			}
			return new String(c);
		}

		/*
		private static String ToDBC(String input)
		{
			// 全角转半角：
			char[] c = input.ToCharArray();
			for (int i = 0; i < c.Length; i++)
			{
				if (c[i] == 12288)
				{
					c[i] = (char)32;
					continue;
				}
				if (c[i] > 65280 && c[i] < 65375)
					c[i] = (char)(c[i] - 65248);
			}
			return new String(c);
		}
		*/

    }
}
