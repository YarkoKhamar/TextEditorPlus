using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TextEditorPlus
{

    public partial class TextEditorPlus : Form
    {
        private int TabCount = 0;

        public TextEditorPlus()
        {
            InitializeComponent();   
        }


        #region Methods
        #region Tabs
        private void AddTab()
        {
            RichTextBox Body = new RichTextBox();

            Body.Name = "Body";
            Body.Dock = DockStyle.Fill;
            Body.ContextMenuStrip = contextMenuStrip1;

            TabPage NewPage = new TabPage();
            TabCount += 1;

            string DocumentText = "Document " + TabCount;
            NewPage.Name = DocumentText;
            NewPage.Text = DocumentText;
            NewPage.Controls.Add(Body);

            tabControl1.TabPages.Add(NewPage);
        }
        private void RemoveTab()
        {
            if(tabControl1.TabPages.Count != 1)
            {
                tabControl1.TabPages.Remove(tabControl1.SelectedTab);
            }
            else
            {
                tabControl1.TabPages.Remove(tabControl1.SelectedTab);
                AddTab();
            }
        }
        private void RemoveAllTabs()
        {
            foreach(TabPage page in tabControl1.TabPages)
            {
                tabControl1.TabPages.Remove(page);
            }
        }
        private void RemoveAllTabsButThis()
        {
            foreach (TabPage page in tabControl1.TabPages)
            {
                if (page.Name != tabControl1.SelectedTab.Name)
                {
                    tabControl1.TabPages.Remove(page);
                }
            }
        }
        #endregion
        #region SaveAndOpen
        private void Save()
        {
            saveFileDialog1.FileName = tabControl1.SelectedTab.Name;
            //tabControl1.SelectedTab.Name = saveFileDialog1.FileName;
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveFileDialog1.Filter = "RTF|*.rtf";
            saveFileDialog1.Title = "Save";
            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if(saveFileDialog1.FileName.Length > 0)
                {
                    GetCurrentDocument.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    tabControl1.SelectedTab.Text = Path.GetFileName(saveFileDialog1.FileName);
                }
            }
        }
        private void SaveAs()
        {
            saveFileDialog1.FileName = tabControl1.SelectedTab.Name;
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveFileDialog1.Filter = "Text Files|*.txt|VB Files|*.vb|C# Files|*.cs|All Files|*.*";
            saveFileDialog1.Title = "Save As";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog1.FileName.Length > 0)
                {
                    GetCurrentDocument.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.PlainText);
                }
            }
        }
        private void Open()
        {
            //openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openFileDialog1.Filter = "RTF|*.rtf|Text Files|*.txt|VB Files|*.vb|C# Files|*.cs|All Files|*.*";

            //if((openFileDialog1.ShowDialog() == DialogResult.OK) && (openFileDialog1.Filter == "RTF Files|*.rtf"))
            //{
            //    if(openFileDialog1.FileName.Length > 0)
            //    {
            //        GetCurrentDocument.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.RichText);
            //    }
            //}
            //else
            //{
            //    if (openFileDialog1.FileName.Length > 0)
            //    {
            //        GetCurrentDocument.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.PlainText);
            //    }
            //}
            openFileDialog1.DefaultExt = "*.rtf";
            openFileDialog1.Filter = "RTF Files|*.rtf|All Files|*.*";
            string ext = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ext = Path.GetExtension(openFileDialog1.FileName);
                if (ext == ".rtf")
                {
                    GetCurrentDocument.Rtf = File.ReadAllText(openFileDialog1.FileName);
                }
                else
                {
                    GetCurrentDocument.Text = File.ReadAllText(openFileDialog1.FileName);
                }
            }
        }
        #endregion
        #region GeneralFonts
        private void GetFontCollection()
        {
            InstalledFontCollection InsFonts = new InstalledFontCollection();

            foreach (FontFamily font in InsFonts.Families)
            {
                if(font.Name != "")
                {
                    toolStripComboBox1.Items.Add(font.Name);
                }
            }

            toolStripComboBox1.SelectedIndex = 0;
        }
        private void ChangeFontSizes()
        {
            for(int i = 1; i < 75; i++)
            {
                toolStripComboBox2.Items.Add(i);
            }

            toolStripComboBox2.SelectedIndex = 11;
        }
        #endregion
        #region ToolStripUpButtons

        private FontStyle GetCurrentFontStyle(FontStyle style)
        {
            FontStyle font = FontStyle.Regular;

            if(GetCurrentDocument.SelectionFont.Bold && style != FontStyle.Bold)
            {
                font = font | FontStyle.Bold;
            }

            if (GetCurrentDocument.SelectionFont.Italic && style != FontStyle.Italic)
            {
                font = font | FontStyle.Italic;
            }

            if (GetCurrentDocument.SelectionFont.Strikeout && style != FontStyle.Strikeout)
            {
                font = font | FontStyle.Strikeout;
            }

            if (GetCurrentDocument.SelectionFont.Underline && style != FontStyle.Underline)
            {
                font = font | FontStyle.Underline;
            }

            return font;
        }

        private void BoldFont()
        {
            FontStyle font = GetCurrentFontStyle(FontStyle.Bold);

            Font boldFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                     GetCurrentDocument.SelectionFont.SizeInPoints, 
                                     GetCurrentDocument.SelectionFont.Style|FontStyle.Bold);
            Font regFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                     GetCurrentDocument.SelectionFont.SizeInPoints,
                                     font);
            if(GetCurrentDocument.SelectionFont.Bold)
            {
                GetCurrentDocument.SelectionFont = regFont;
            }
            else
            {
                GetCurrentDocument.SelectionFont = boldFont;
            }
        }
        private void ItalicFont()
        {
            FontStyle font = GetCurrentFontStyle(FontStyle.Italic);

            Font italicFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                    GetCurrentDocument.SelectionFont.SizeInPoints,
                                    GetCurrentDocument.SelectionFont.Style | FontStyle.Italic);
            Font regFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                     GetCurrentDocument.SelectionFont.SizeInPoints,
                                     font);
            if (GetCurrentDocument.SelectionFont.Italic)
            {
                GetCurrentDocument.SelectionFont = regFont;
            }
            else
            {
                GetCurrentDocument.SelectionFont = italicFont;
            }
        }
        private void UnderlinedFont()
        {
            FontStyle font = GetCurrentFontStyle(FontStyle.Underline);

            Font underlinedFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                    GetCurrentDocument.SelectionFont.SizeInPoints,
                                    GetCurrentDocument.SelectionFont.Style | FontStyle.Underline);
            Font regFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                     GetCurrentDocument.SelectionFont.SizeInPoints,
                                     font);
            if (GetCurrentDocument.SelectionFont.Underline)
            {
                GetCurrentDocument.SelectionFont = regFont;
            }
            else
            {
                GetCurrentDocument.SelectionFont = underlinedFont;
            }
        }
        private void StrikeoutFont()
        {
            FontStyle font = GetCurrentFontStyle(FontStyle.Strikeout);

            Font strikedoutFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                    GetCurrentDocument.SelectionFont.SizeInPoints,
                                    GetCurrentDocument.SelectionFont.Style | FontStyle.Strikeout);
            Font regFont = new Font(GetCurrentDocument.SelectionFont.FontFamily,
                                     GetCurrentDocument.SelectionFont.SizeInPoints,
                                     font);
            if (GetCurrentDocument.SelectionFont.Strikeout)
            {
                GetCurrentDocument.SelectionFont = regFont;
            }
            else
            {
                GetCurrentDocument.SelectionFont = strikedoutFont;
            }
        }
        private void ToUpperFont()
        {
            GetCurrentDocument.SelectedText = GetCurrentDocument.SelectedText.ToUpper();
        }
        private void ToLowerFont()
        {
            GetCurrentDocument.SelectedText = GetCurrentDocument.SelectedText.ToLower();
        }
        private void FontIncrease()
        {
            float newFontSize = 0;
            if (GetCurrentDocument.SelectionFont.SizeInPoints <= 73)
            {
                newFontSize = GetCurrentDocument.SelectionFont.SizeInPoints + 2;
            }
            Font newSize = new Font(GetCurrentDocument.SelectionFont.Name,
                                    newFontSize,
                                    GetCurrentDocument.SelectionFont.Style);
            GetCurrentDocument.SelectionFont = newSize;
        }
        private void FontDecrease()
        {
            float newFontSize = 0;
            if (GetCurrentDocument.SelectionFont.SizeInPoints >= 3)
            {
                newFontSize = GetCurrentDocument.SelectionFont.SizeInPoints - 2;
            }
            Font newSize = new Font(GetCurrentDocument.SelectionFont.Name,
                                    newFontSize,
                                    GetCurrentDocument.SelectionFont.Style);
            GetCurrentDocument.SelectionFont = newSize;
        }
        private void FontForecolor()
        {
            if(colorDialog1.ShowDialog() == DialogResult.OK)
            {
                GetCurrentDocument.SelectionColor = colorDialog1.Color;
            }
        }
        private void HighlightGreen()
        {
            GetCurrentDocument.SelectionBackColor = Color.LawnGreen;
        }
        private void HighlightYellow()
        {
            GetCurrentDocument.SelectionBackColor = Color.Yellow;
        }
        private void HighlightOrange()
        {
            GetCurrentDocument.SelectionBackColor = Color.Orange;
        }
        private void FontFam()
        {
            Font newFont = new Font(toolStripComboBox1.SelectedItem.ToString(),
                                    GetCurrentDocument.SelectionFont.Size,
                                    GetCurrentDocument.SelectionFont.Style);

            GetCurrentDocument.SelectionFont = newFont;
        }
        private void FontSize()
        {
            float newSize = 0;
            float.TryParse(toolStripComboBox2.SelectedItem.ToString(), out newSize);

            Font newFont = new Font(GetCurrentDocument.SelectionFont.Name,
                                    newSize,
                                    GetCurrentDocument.SelectionFont.Style);

            GetCurrentDocument.SelectionFont = newFont;
        }
        #endregion
        #region TextFunctions
        private String Justify(String s, Int32 count)
        {
            if (count <= 0)
                return s;

            Int32 middle = s.Length / 2;
            IDictionary<Int32, Int32> spaceOffsetsToParts = new Dictionary<Int32, Int32>();
            String[] parts = s.Split(' ');
            for (Int32 partIndex = 0, offset = 0; partIndex < parts.Length; partIndex++)
            {
                spaceOffsetsToParts.Add(offset, partIndex);
                offset += parts[partIndex].Length + 1;
            }
            foreach (var pair in spaceOffsetsToParts.OrderBy(entry => Math.Abs(middle - entry.Key)))
            {
                count--;
                if (count < 0)
                    break;
                parts[pair.Value] += ' ';
            }
            return String.Join(" ", parts);
        }
        private void Undo()
        {
            GetCurrentDocument.Undo();
        }
        private void Redo()
        {
            GetCurrentDocument.Redo();
        }
        private void Cut()
        {
            GetCurrentDocument.Cut();
        }
        private void Copy()
        {
            GetCurrentDocument.Copy();
        }
        private void Paste()
        {
            GetCurrentDocument.Paste();
        }
        private void SelectAll()
        {
            GetCurrentDocument.SelectAll();
        }
        #endregion
        #region StatusStrip
        private int CountWords()
        {
            var text = GetCurrentDocument.Text.Trim();
            int wordCount = 0, index = 0;

            while (index < text.Length)
            {
                while (index < text.Length && !char.IsWhiteSpace(text[index]))
                {
                    index++;
                }

                wordCount++;

                while (index < text.Length && char.IsWhiteSpace(text[index]))
                {
                    index++;
                }
            }

            return wordCount;
        }
        #endregion
        #region ToolStripDownButtons
        private void CreateList()
        {
            
        }
        private void Search()
        {
            if (GetCurrentDocument.Text != string.Empty)
            {
                int index = 0;
                var temp = GetCurrentDocument.Text;
                
                GetCurrentDocument.Text = "";
                GetCurrentDocument.Text = temp;
                while (index < GetCurrentDocument.Text.LastIndexOf(toolStripTextBox1.Text))
                {
                    GetCurrentDocument.Find(toolStripTextBox1.Text, index, GetCurrentDocument.TextLength, RichTextBoxFinds.None);
                    GetCurrentDocument.SelectionBackColor = Color.Yellow;
                    index = GetCurrentDocument.Text.IndexOf(toolStripTextBox1.Text, index) + 1;
                    GetCurrentDocument.Select();
                }
            }
        }
        private void Replace()
        {
            int i = 0;
            int n = 0;
            int a = toolStripTextBox2.Text.Length - toolStripTextBox1.Text.Length;
            foreach (Match m in Regex.Matches(GetCurrentDocument.Text.ToString(), toolStripTextBox1.Text.ToString()))
            {
                GetCurrentDocument.Select(m.Index + i, toolStripTextBox1.Text.Length);
                i += a;
                GetCurrentDocument.SelectedText = toolStripTextBox2.Text.ToString();
                n++;
            }
            MessageBox.Show("Replaced " + n + " matches!");
        }
        private void InsertTable()
        {
            StringBuilder rtf = new StringBuilder();

            rtf.Append(@"{\rtf1");
            rtf.Append(@"\trowd");
            rtf.Append(@"\cellx1000");
            rtf.Append(@"\cellx2000");
            rtf.Append(@"\intbl \cell \cell \row");
            rtf.Append(@"\pard");
            rtf.Append(@"}");
            GetCurrentDocument.Rtf = rtf.ToString();
        }
        private void AllowDragAndDrop()
        {
            GetCurrentDocument.EnableAutoDragDrop = true;
        }
        private void DisallowDragAndDrop()
        {
            GetCurrentDocument.EnableAutoDragDrop = false;
        }
        #endregion
        #endregion
        
        #region Properties
        public string GetSearchText
        {
            get { return toolStripTextBox1.Text.ToString();}
        }
        private RichTextBox GetCurrentDocument
        {
            get { return (RichTextBox)tabControl1.SelectedTab.Controls["Body"]; }
        }
        #endregion
        
        #region Events

        private void ChangeElementCheckedState(ToolStripButton button)
        {
            if(button.Checked == true)
            {
                button.Checked = false;
            }
            else
            {
                button.Checked = true;
            }
        }

        private void TextEditorPlus_Load(object sender, EventArgs e)
        {
            AddTab();
            GetFontCollection();
            ChangeFontSizes();
            cutToolStripMenuItem1.Enabled = false;
            copyToolStripMenuItem1.Enabled = false;
            pasteToolStripMenuItem1.Enabled = false;
        }
        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            //
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Redo();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Paste();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectAll();
        }

        private void cutToolStripButton_Click(object sender, EventArgs e)
        {
            Cut();
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {
            Copy();
        }

        private void pasteToolStripButton_Click(object sender, EventArgs e)
        {
            Paste();
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddTab();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Open();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAs();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {
            Open();
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void undoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Undo();
        }

        private void redoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Redo();
        }

        private void cutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Cut();
        }

        private void copyToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Copy();
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Paste();
        }
        
        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RemoveTab();
        }

        private void closeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RemoveAllTabs();
        }

        private void closeAllButThisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RemoveAllTabsButThis();
        }

        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            AddTab();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (GetCurrentDocument.Text.Length > 0)
            {
                toolStripStatusLabel6.Text = CountWords().ToString();
                toolStripStatusLabel7.Text = GetCurrentDocument.Text.Length.ToString();
                toolStripStatusLabel2.Text = GetCurrentDocument.Lines.Count().ToString();
            }
            if(Clipboard.ContainsText() | Clipboard.ContainsImage())
            {
                pasteToolStripMenuItem1.Enabled = true;
            }
            if(GetCurrentDocument.SelectedText != string.Empty)
            {
                cutToolStripMenuItem1.Enabled = true;
                copyToolStripMenuItem1.Enabled = true;
            }
            else
            {
                cutToolStripMenuItem1.Enabled = false;
                copyToolStripMenuItem1.Enabled = false;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            BoldFont();
            ChangeElementCheckedState(toolStripButton1);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ItalicFont();
            ChangeElementCheckedState(toolStripButton2);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            UnderlinedFont();
            ChangeElementCheckedState(toolStripButton3);
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            StrikeoutFont();
            ChangeElementCheckedState(toolStripButton4);
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            ToUpperFont();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            ToLowerFont();
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            FontIncrease();
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            FontDecrease();
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            FontForecolor();
        }

        private void RemoveTabToolStripButton_Click(object sender, EventArgs e)
        {
            RemoveTab();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //
        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            GetCurrentDocument.SelectionAlignment = HorizontalAlignment.Left;

            ChangeElementCheckedState(toolStripButton10);
            toolStripButton11.Checked = false;
            toolStripButton12.Checked = false;
        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            GetCurrentDocument.SelectionAlignment = HorizontalAlignment.Center;

            ChangeElementCheckedState(toolStripButton11);
            toolStripButton10.Checked = false;
            toolStripButton12.Checked = false;
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            GetCurrentDocument.SelectionAlignment = HorizontalAlignment.Right;

            ChangeElementCheckedState(toolStripButton12);
            toolStripButton10.Checked = false;
            toolStripButton11.Checked = false;
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            //
        }

        private void HighGreen_Click(object sender, EventArgs e)
        {
            HighlightGreen();
        }

        private void toolStripButton13_ButtonClick(object sender, EventArgs e)
        {
            GetCurrentDocument.SelectionBackColor = Color.Transparent;
        }

        private void HighYellow_Click(object sender, EventArgs e)
        {
            HighlightYellow();
        }

        private void HighOrange_Click(object sender, EventArgs e)
        {
            HighlightOrange();
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            GetCurrentDocument.Text = Justify(GetCurrentDocument.SelectedText, 100);
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            Undo();
        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            Redo();
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            FontFam();
        }

        private void toolStripComboBox2_Click(object sender, EventArgs e)
        {
            //
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            FontSize();
        }
        
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            MessageBox.Show("Text Editor Plus.\n \u00A9 Yaroslav Khamar, 2016","About");
        }

        private void tabControl1_CursorChanged(object sender, EventArgs e)
        {
            //int line = GetCurrentDocument.CurrentLine;
            //int col = GetCurrentDocument.CurrentColumn;
            //int pos = GetCurrentDocument.CurrentPosition;

            //toolStripStatusLabel1.Text = "Line " + line + ", Col " + col +
            //                 ", Position " + pos;
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            SelectAll();
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            GetCurrentDocument.Clear();
        }

        private void toolStripButton19_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {
            Replace();
        }

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            InsertTable();
        }

        private void toolStripButton22_Click(object sender, EventArgs e)
        {
            AllowDragAndDrop();
        }

        private void toolStripButton23_Click(object sender, EventArgs e)
        {
            DisallowDragAndDrop();
        }

        private void contentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help helpForm = new Help();
            helpForm.Show();
        }
        #endregion

        private void darkStyleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackColor = Color.DarkGray;
        }

        private void lightStyleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackColor = Color.LightSkyBlue;
        }

        private void menuStripToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void lightThemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            menuStrip1.BackColor = Color.LightSkyBlue;
            menuStrip1.ForeColor = Color.Black;
        }

        private void darkThemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            menuStrip1.BackColor = Color.DarkCyan;
            menuStrip1.ForeColor = Color.Yellow;
        }

        private void lightStyleToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            toolStrip2.BackColor = Color.LightSkyBlue;
        }

        private void darkStyleToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            toolStrip2.BackColor = Color.DarkCyan;
        }

        private void lightStyleToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            toolStrip3.BackColor = Color.LightSkyBlue;
        }

        private void darkStyleToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            toolStrip3.BackColor = Color.DarkCyan;
        }

        private void lightStyleToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            toolStrip1.BackColor = Color.LightSkyBlue;
        }

        private void darkStyleToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            toolStrip1.BackColor = Color.DarkCyan;
        }

        private void lightStyleToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            statusStrip1.BackColor = Color.LightSkyBlue;
            statusStrip1.ForeColor = Color.Black;
        }

        private void darkStyleToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            statusStrip1.BackColor = Color.DarkCyan;
            statusStrip1.ForeColor = Color.Yellow;
        }

        private void toolStripDownToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void lightStyleToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            toolStripContainer1.LeftToolStripPanel.BackColor = Color.LightSkyBlue;
            toolStripContainer1.TopToolStripPanel.BackColor = Color.LightSkyBlue;
        }

        private void darkStyleToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            toolStripContainer1.LeftToolStripPanel.BackColor = Color.DarkCyan;
            toolStripContainer1.TopToolStripPanel.BackColor = Color.DarkCyan;
        }

        private void toolStripUpToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void languageToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ukrainianToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fileToolStripMenuItem.Text = "&Файл";
            editToolStripMenuItem.Text = "&Редагувати";
            toolsToolStripMenuItem.Text = "&Інструменти";
            helpToolStripMenuItem.Text = "&Допомога";
            newToolStripButton.Text = "&Новий файл";
            openToolStripButton.Text = "&Відкрити";
            saveToolStripButton.Text = "&Зберегти";
            saveAsToolStripMenuItem.Text = "&Зберегти як";
            exitToolStripMenuItem.Text = "&Вихід";
            undoToolStripMenuItem.Text = "&Назад";
            redoToolStripMenuItem.Text = "&Вперед";
            copyToolStripMenuItem.Text = "&Копіювати";
            pasteToolStripMenuItem.Text = "&Вставити";
            cutToolStripMenuItem.Text = "&Вирізати";
            selectAllToolStripMenuItem.Text = "&Виділити все";
            customizeToolStripMenuItem.Text = "&Персоналізація";
            languageToolStripMenuItem.Text = "&Мова";
            ukrainianToolStripMenuItem.Text = "&Українська";
            englishToolStripMenuItem.Text = "&Англійська";
            menuStripToolStripMenuItem.Text = "&Меню";
            toolStripUpToolStripMenuItem.Text = "&Інструменти (зверху)";
            toolStripDownToolStripMenuItem.Text = "&Інструменти (знизу)";
            leftTooStripToolStripMenuItem.Text = "&Інструменти зліва (вертикально)";
            statusBarToolStripMenuItem.Text = "&Рядок статистики";
            edgesToolStripMenuItem.Text = "&Рамки";
            lightThemeToolStripMenuItem.Text = "&Світлий стиль";
            lightStyleToolStripMenuItem.Text = "&Світлий стиль";
            lightStyleToolStripMenuItem1.Text = "&Світлий стиль";
            lightStyleToolStripMenuItem2.Text = "&Світлий стиль";
            lightStyleToolStripMenuItem3.Text = "&Світлий стиль";
            lightStyleToolStripMenuItem4.Text = "&Світлий стиль";
            darkThemeToolStripMenuItem.Text = "&Темний стиль";
            darkStyleToolStripMenuItem.Text = "&Темний стиль";
            darkStyleToolStripMenuItem1.Text = "&Темний стиль";
            darkStyleToolStripMenuItem2.Text = "&Темний стиль";
            darkStyleToolStripMenuItem3.Text = "&Темний стиль";
            darkStyleToolStripMenuItem4.Text = "&Темний стиль";
            contentsToolStripMenuItem.Text = "&Інструкція";
            aboutToolStripMenuItem.Text = "&Про програму";
            newToolStripMenuItem.Text = "&Новий файл";
            openToolStripMenuItem.Text = "&Відкрити";
            saveToolStripMenuItem.Text = "&Зберегти";
            copyToolStripButton.Text = "&Копіювати";
            cutToolStripButton.Text = "&Вирізати";
            pasteToolStripButton.Text = "&Вставити";
            RemoveTabToolStripButton.Text = "&Видалити сторінку";
            toolStripStatusLabel3.Text = "&Символи:";
            toolStripStatusLabel4.Text = "&Рядки:";
            toolStripStatusLabel5.Text = "&Слова:";
        }

        private void defaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            menuStrip1.BackColor = SystemColors.Control;
            menuStrip1.ForeColor = Color.Black;
        }

        private void defaultToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            toolStrip2.BackColor = SystemColors.Control;
        }

        private void defaultToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            toolStrip3.BackColor = SystemColors.Control;
        }

        private void defaultToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            toolStrip1.BackColor = SystemColors.Control;
        }

        private void defaultToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            toolStripContainer1.LeftToolStripPanel.BackColor = SystemColors.Control;
            toolStripContainer1.TopToolStripPanel.BackColor = SystemColors.Control;
        }

        private void defaultToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            statusStrip1.BackColor = SystemColors.Control;
            statusStrip1.ForeColor = Color.Black;
        }

        private void englishToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fileToolStripMenuItem.Text = "&File";
            editToolStripMenuItem.Text = "&Edit";
            toolsToolStripMenuItem.Text = "&Tools";
            helpToolStripMenuItem.Text = "&Help";
            newToolStripButton.Text = "&New";
            openToolStripButton.Text = "&Open";
            saveToolStripButton.Text = "&Save";
            saveAsToolStripMenuItem.Text = "&Save As";
            exitToolStripMenuItem.Text = "&Exit";
            undoToolStripMenuItem.Text = "&Undo";
            redoToolStripMenuItem.Text = "&Redo";
            copyToolStripMenuItem.Text = "&Copy";
            pasteToolStripMenuItem.Text = "&Paste";
            cutToolStripMenuItem.Text = "&Cut";
            selectAllToolStripMenuItem.Text = "&Select all";
            customizeToolStripMenuItem.Text = "&Customize";
            languageToolStripMenuItem.Text = "&Language";
            ukrainianToolStripMenuItem.Text = "&Ukrainian";
            englishToolStripMenuItem.Text = "&English";
            menuStripToolStripMenuItem.Text = "&Menu Strip";
            toolStripUpToolStripMenuItem.Text = "&Tool Strip Up";
            toolStripDownToolStripMenuItem.Text = "&Tool Strip Down";
            leftTooStripToolStripMenuItem.Text = "&Left Tool Strip";
            statusBarToolStripMenuItem.Text = "&Status Bar";
            edgesToolStripMenuItem.Text = "&Edges";
            lightThemeToolStripMenuItem.Text = "&Light Style";
            lightStyleToolStripMenuItem.Text = "&Light Style";
            lightStyleToolStripMenuItem1.Text = "&Light Style";
            lightStyleToolStripMenuItem2.Text = "&Light Style";
            lightStyleToolStripMenuItem3.Text = "&Light Style";
            lightStyleToolStripMenuItem4.Text = "&Light Style";
            darkThemeToolStripMenuItem.Text = "&Dark Style";
            darkStyleToolStripMenuItem.Text = "&Dark Style";
            darkStyleToolStripMenuItem1.Text = "&Dark Style";
            darkStyleToolStripMenuItem2.Text = "&Dark Style";
            darkStyleToolStripMenuItem3.Text = "&Dark Style";
            darkStyleToolStripMenuItem4.Text = "&Dark Style";
            contentsToolStripMenuItem.Text = "&Instructions";
            aboutToolStripMenuItem.Text = "&About";
            newToolStripMenuItem.Text = "&New";
            openToolStripMenuItem.Text = "&Open";
            saveToolStripMenuItem.Text = "&Save";
            copyToolStripButton.Text = "&Copy";
            cutToolStripButton.Text = "&Cut";
            pasteToolStripButton.Text = "&Paste";
            RemoveTabToolStripButton.Text = "&Remove Tab";
            toolStripStatusLabel3.Text = "&Characters:";
            toolStripStatusLabel4.Text = "&Lines:";
            toolStripStatusLabel5.Text = "&Words:";
        }

        private void toolStripButton24_Click(object sender, EventArgs e)
        {
            OpenFileDialog op1 = new OpenFileDialog();
            op1.Filter = "All files|*.*";
            if(op1.ShowDialog() == DialogResult.OK)
            {
                var image = new Bitmap(op1.FileName);
                Clipboard.SetImage(image);
                GetCurrentDocument.Paste();
            }
        }
        
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripButton25_Click(object sender, EventArgs e)
        {
            SendKeys.Send("^+L^+L");
        }
    }
}
