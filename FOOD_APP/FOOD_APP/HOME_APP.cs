using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace FOOD_APP
{
    public partial class HOME_APP : Form
    {
        public HOME_APP()
        {
            InitializeComponent();
        }
        
        private void createEventCB()
        {
            for(int i = 1; i < 23; i++)
            {
                CheckBox cb = this.Controls.Find("cb" + i.ToString(), true).First() as CheckBox;
                cb.CheckedChanged += new EventHandler(this.CbCheckedChanged);

                CheckBox c = this.Controls.Find("c" + i.ToString(), true).First() as CheckBox;
                c.CheckedChanged += new EventHandler(this.CCheckedChanged);

            }
        }

        private string run_weka(String cmd)
        {
            string output;
            
                StreamWriter sw;
                StreamReader sr;
                StreamReader err;
                Process dir = new Process();

                dir.StartInfo.FileName = "CMD.EXE";
                dir.StartInfo.UseShellExecute = false;
                dir.StartInfo.CreateNoWindow = true;
                dir.StartInfo.RedirectStandardInput = true;
                dir.StartInfo.RedirectStandardError = true;
                dir.StartInfo.RedirectStandardOutput = true;

                dir.Start();


                sw = dir.StandardInput;
                sr = dir.StandardOutput;
                err = dir.StandardError;
                sw.AutoFlush = true;

                sw.WriteLine(cmd);
                
                sw.Close();
                err.Close();
            output = sr.ReadToEnd().ToString();
                sr.Close();
                return output;
        }
        
        public void create_model()
        {
            String weka_path = @"C:\Program Files\Weka-3-9\";
            String cmd = "java -cp \"" + weka_path +
                "weka.jar\" weka.classifiers.trees.J48 -t \"" + Application.StartupPath + "\\trainingData\\AI_project.arff\"" +
                " -d \"" + Application.StartupPath + "\\model\\AI_project.model\"";
            String output = run_weka(cmd);
            txtMonitor.Text = output.Replace("\n", Environment.NewLine);
                        
        }

        public void OnRun()
        {
            SaveValueToExcel();
        }

        private void HOME_APP_Load(object sender, EventArgs e)
        {
            try
            {
                create_model();

                createEventCB();

                

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                OnRun();
                test_model();

                for (int i = 1; i < 23; i++)
                {
                    CheckBox cb = this.Controls.Find("cb" + i.ToString(), true).First() as CheckBox;
                    cb.Checked = true;
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        void CbCheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;
            if (cb.Checked)
            {
                CheckBox t = this.Controls.Find(cb.Name.Replace("b", ""), true).First() as CheckBox;
                t.Checked = false;
            }
            else
            {
                CheckBox t = this.Controls.Find(cb.Name.Replace("b", ""), true).First() as CheckBox;
                t.Checked = true;
            }
        }

        void CCheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;
            if (cb.Checked)
            {
                CheckBox t = this.Controls.Find(cb.Name.Replace("c", "cb"), true).First() as CheckBox;
                t.Checked = false;
            }
            else
            {
                CheckBox t = this.Controls.Find(cb.Name.Replace("c", "cb"), true).First() as CheckBox;
                t.Checked = true;
            }
        }

        private void SaveValueToExcel()
        {

            CheckFile();


                Excel.Application app = new Excel.Application();

                Excel.Workbook book = app.Workbooks.Open(Application.StartupPath + @"\Template\Work.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            try
            {

                int N = 22;
                string[] position = new string[N];
                char cc = 'A';
                for (int i = 0; i < N; i++)
                {
                    position[i] = cc + "2";
                    cc++;

                }
                for (int i = 1; i < 23; i++)
                {
                    
                        app.Cells[2, i] = "NO";
                    
                        app.Cells[3, i] = "YES";
                        app.Cells[4, i] = "YES";

                }
                app.Cells[2, 23] = "N";
                app.Cells[3, 23] = "S";
                app.Cells[4, 23] = "W";

                for (int i = 1; i < 23; i++)
                {
                    CheckBox c = this.Controls.Find("c" + i.ToString(), true).First() as CheckBox;
                    if (c.Checked)
                    {
                        app.Cells[5, i] = "YES";
                    }
                    else
                    {
                        app.Cells[5, i] = "NO";
                    }

                }




                book.SaveAs(Application.StartupPath + @"/TestFile/FileTestData.csv"
                    , Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows
                    , Type.Missing
                    , Type.Missing
                    , false
                    , false
                    , Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive
                    , Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges
                    , false
                    , Type.Missing
                    , Type.Missing
                    , Type.Missing);

                book.Close(false, Type.Missing, Type.Missing);

                app.Quit();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            


        }

        private void CheckFile()
        {
            string curFile = Application.StartupPath + @"/TestFile/FileTestData.csv";
            if (File.Exists(curFile))
            {
                File.Delete(curFile);
            }
        }

        private void convertCSVtoARFF()
        {
            String weka_path = @"C:\Program Files\Weka-3-9\";
            string curFile = Application.StartupPath + @"\TestFile\FileTestData.csv";
            string toFile = Application.StartupPath + @"\TestFile\FileTestData.arff";

            String cmd = "java -cp \"" + weka_path + "weka.jar\" weka.core.converters.CSVLoader " + curFile + " > " + toFile;

            String output = run_weka(cmd);
        }

        private void test_model()
        {
            convertCSVtoARFF();
            String weka_path = @"C:\Program Files\Weka-3-9\";
            string curFile = Application.StartupPath + @"\TestFile\FileTestData.arff";
            string model = Application.StartupPath + @"\model\modelTestt.model";

            String cmd = "java -cp \"" + weka_path + "weka.jar\" weka.classifiers.trees.J48 -T \"" + curFile + "\"" + " -l \"" + model + "\" -p 0";

            String output = run_weka(cmd);
            

            string temp = string.Empty;
            bool status = false;
            string gg = "";
            for (int i = 0; i < output.Length; i++)
            {
                if (output[i] == '?') {
                    status = true;
                }
                if (status)
                {
                        gg += output[i];
                    if (output[i] == 'S' || output[i] == 'W' || output[i] == 'N')
                    {
                        temp = output[i].ToString();
                        status = false;
                        break;
                    }
                }
            }
           
            string showMeg = string.Empty;
            if (temp == "N") showMeg = "ภาคเหนือ";
            else if (temp == "W") showMeg = "ภาคอีสาน";
            else if (temp == "S") showMeg = "ภาคใต้";

            txtMonitor.Text = output.Replace("\n", Environment.NewLine);


            MessageBox.Show("คำนวณเสร็จแล้ว !! คำตอบคือ " + showMeg);
            gg = "";
        }

        
    }
}
