using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TextEditorPlus
{
    public partial class findReplaceForm : Form
    {
        public findReplaceForm()
        {
            
            InitializeComponent();
            textBox1.Text = ((TextEditorPlus)ActiveForm).GetSearchText;

        }
        public string GetText1()
        {
            return textBox1.Text;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }
    }
}
