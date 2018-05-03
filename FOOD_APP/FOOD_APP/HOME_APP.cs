using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FOOD_APP
{
    public partial class HOME_APP : Form
    {
        public HOME_APP()
        {
            InitializeComponent();
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

        }

        private void HOME_APP_Load(object sender, EventArgs e)
        {
            try
            {
                create_model();

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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
