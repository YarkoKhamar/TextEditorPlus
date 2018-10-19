using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TextEditorPlus
{
    public partial class Help : Form
    {
        public Help()
        {
            InitializeComponent();
            tabPage1.Text = "Огляд Інтерфейсу";
            tabPage2.Text = "Опис Функціоналу";
            richTextBox2.LoadFile(@"../../Instructions.rtf", RichTextBoxStreamType.RichText);
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
    }
}
