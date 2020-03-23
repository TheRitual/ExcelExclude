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
        private String pathA = "", pathB = "";

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
            check();
        }

        public void loadFile(char aorb)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "Old Microsoft Excel (*.xls)|*.xls|Microsoft Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                    if (aorb == 'A')
                    {
                        setFileAPath(filePath);
                        check();
                    } else
                    {
                        setFileBPath(filePath);
                        check();
                    }
                    addLog("Wybrano ścieżkę pliku " + aorb + ":");
                    addLog(filePath);
                }
            }
        }

        private void check()
        {
            if (pathA == "")
            {
                information("Wybierz plik A");
                lockControls(true, false, false, false, false, false, false);
            } else if (pathB == "")
            {
                information("Wybierz plik B");
                lockControls(true, true, false, false, false, false, false);
            } else if (pathA == pathB)
            {
                information("Proszę wybrać 2 różne pliki");
                lockControls(true, true, false, false, false, false, false);
            } else
            {
                information("Wybierz Arkusze plików A oraz B");
                lockControls(true, true, true, true, false, false, false);
            }
        }

        private void lockControls(bool fileAButton, bool fileBButton, bool comboSheetA, bool comboSheetB, bool comboColumnA, bool comboColumnB, bool compareButton)
        {
            this.fileAButton.Enabled = fileAButton;
            this.fileBButton.Enabled = fileBButton;
            this.comboSheetA.Enabled = comboSheetA;
            this.comboSheetB.Enabled = comboSheetB;
            this.comboColumnA.Enabled = comboColumnA;
            this.comboColumnB.Enabled = comboColumnB;
            this.compareButton.Enabled = compareButton;
        }
        
        private void information(String info)
        {
            this.infoLabel.Text = info;
        }

        private void addLog(String log)
        {
            logBox.Select(logBox.TextLength + 1, 0);
            logBox.SelectedText = "\r\n" + getTimeStamp() + " ::: " + log;
        }

        private void setFileAPath(String path)
        {
            pathALabel.Text = path;
            this.pathA = path;
        }

        private void setFileBPath(String path)
        {
            pathBLabel.Text = path;
            this.pathB = path;
        }

        private String getFileAPath()
        {
            return pathA;
        }

        private String getFileBPath()
        {
            return pathB;
        }

        private String getTimeStamp()
        {
            DateTime date = DateTime.Now;
            return date.ToString("[HH:mm:ss:ffff] ");
        }
    }
}
