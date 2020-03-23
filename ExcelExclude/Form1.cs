using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelExclude
{
    public partial class fileButtonA : Form
    {
        AppFlow flow;

        public fileButtonA()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            flow = new AppFlow(fileAButton, fileBButton, pathLabelA, pathLabelB, comboSheetA, comboSheetB, comboColumnA, comboColumnB, infoLabel, logTextBox1, compareButton);
        }

        private void exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void fileAButton_Click(object sender, EventArgs e)
        {
            flow.loadFile('A');
        }

        private void fileBButton_Click(object sender, EventArgs e)
        {
            flow.loadFile('B');
        }
    }
}
