using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExclude
{
    class AppFlow
    {
        private Button fileAButton, fileBButton, compareButton;
        private Label pathALabel, pathBLabel, infoLabel;
        private ComboBox comboSheetA, comboSheetB, comboColumnA, comboColumnB;
        private TextBox logBox;
        private String pathA, pathB;

        public AppFlow(Button fileAButton, Button fileBButton, Label pathALabel, Label pathBLabel, ComboBox comboSheetA, ComboBox comboSheetB, ComboBox comboColumnA, ComboBox comboColumnB, Label infoLabel, TextBox logBox, Button compareButton)
        {
            this.fileAButton = fileAButton;
            this.fileBButton = fileBButton;
            this.compareButton = compareButton;
            this.pathALabel = pathALabel;
            this.pathBLabel = pathBLabel;
            this.infoLabel = infoLabel;
            this.comboSheetA = comboSheetA;
            this.comboSheetB = comboSheetB;
            this.comboColumnA = comboColumnA;
            this.comboColumnB = comboColumnB;
            this.logBox = logBox;
        }

        public void setFileAPath(String path)
        {
            this.pathA = path;
        }

        public void setFileBPath(String path)
        {
            this.pathB = path;
        }

        public String getFileAPath()
        {
            return pathA;
        }

        public String getFileBPath()
        {
            return pathB;
        }
    }
}
