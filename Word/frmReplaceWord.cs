using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Novacode;
using System.IO;
using System.Diagnostics;

namespace Word
{
    public partial class frmReplaceWord : Form
    {
        public frmReplaceWord()
        {
            InitializeComponent();
        }
        private List<DataWord> sData = new List<DataWord>();
        String name_file = "VPB-TBLAN1.docx";
        private void btnIn_Click(object sender, EventArgs e)
        {
            CreateData();
            word.Application wordApp = new word.Application();
            //String name_file = "VPB-TBLAN1.docx";
            String StartupPath = System.Windows.Forms.Application.StartupPath + "\\" + name_file;
            Document newDocument = wordApp.Documents.Open(StartupPath, ReadOnly: true);
            wordApp.Visible = false;
            if (!File.Exists(StartupPath))
            {
                MessageBox.Show("Không tìm thấy file Template", "Thông Báo");
                return;
            }
            else
            {
                try
                {
                    foreach (DataWord dt in sData)
                    {
                        while (newDocument.Content.Find.Execute(FindText: dt.K_Name))
                        {
                            if (newDocument.Content.Find.Execute(FindText: dt.K_Name))
                            {
                                newDocument.Content.Find.Execute(FindText: dt.K_Name, ReplaceWith: dt.K_Replace, Replace: WdReplace.wdReplaceOne);
                            }
                        }
                        richTextBox1.Text = newDocument.Content.Text.ToString();
                    }

                    newDocument.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Temp.docx");
                    newDocument.Close();
                }
                catch (Exception)
                {
                    newDocument.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Temp " + DateTime.Now.ToString("HHmmss") + ".docx");
                    newDocument.Close();
                }
            }
        }
        private void CreateData()
        {
            sData.Add(new DataWord("[STT]", "NHAC_"));
            sData.Add(new DataWord("[ChiNhanh]", "NHAC_"));
            sData.Add(new DataWord("[BankName]", "NHAC_"));
            sData.Add(new DataWord("[CustomerName]", "NHAC_"));
            sData.Add(new DataWord("[CustomerAddress]", "NHAC_"));
            sData.Add(new DataWord("[BankNo]", "NHAC_"));
            sData.Add(new DataWord("[Date]", "NHAC_"));
            sData.Add(new DataWord("[TotalAmt]", "NHAC_"));
            sData.Add(new DataWord("[PayReqDate]", "NHAC_"));
            sData.Add(new DataWord("[LoanNo_OverDueAmt_OverDueDate]", "NHAC_"));
            sData.Add(new DataWord("[EmployessName]", "NHAC_"));
            sData.Add(new DataWord("[Position]", "NHAC_"));
            sData.Add(new DataWord("[Phone]", "NHAC_"));
        }
        public List<Processes> GetRunningProcesses()
        {
            List<Processes> ProcessList = new List<Processes>();
            //here we're going to get a list of all running processes on
            //the computer
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("WINWORD"))
                {
                    ProcessList.Add(new Processes(clsProcess.Id, clsProcess.MainWindowTitle.ToString()));
                }
            }
            return ProcessList;
        }
        private void KillProcesses(List<Processes> processesbeforegen, List<Processes> processesaftergen)
        {
            foreach (Processes pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (Processes pidbefore in processesbeforegen)
                {
                    if (pidafter.K_ID == pidbefore.K_ID)
                    {
                        processfound = true;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter.K_ID);
                    clsProcess.Kill();
                }
            }
        }
        private void KillProcesses(List<Processes> LstProcess, String name_file)
        {

            foreach (Processes proc in LstProcess)
            {
                if (proc.K_NAME.Contains(name_file))
                {
                    Process clsProcess = Process.GetProcessById(proc.K_ID);
                    clsProcess.Kill();
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //List<Processes> processesbeforegen = GetRunningProcesses(); // lấy danh sách trước khi code 
            //// APP CREATION/ DOCUMENT CREATION HERE...
            //List<Processes> processesaftergen = GetRunningProcesses(); // lấy danh sách sau khi code
            // KillProcesses(processesbeforegen, processesaftergen); // kill process phát sinh
            List<Processes> LstProcess = GetRunningProcesses();
            KillProcesses(LstProcess, name_file);
        }
    }
    public class Processes
    {
        private int ID;
        private String NAME;

        public Processes(int K_ID, String K_NAME)
        {
            ID = K_ID;
            NAME = K_NAME;
        }

        public int K_ID
        {
            get { return ID; }
            set { ID = value; }
        }
        public String K_NAME
        {
            get { return NAME; }
            set { NAME = value; }
        }
    }
    public class DataWord
    {
        private String Name;
        private String Replace;

        public DataWord(String K_Name, String K_Replace)
        {
            Name = K_Name;
            Replace = K_Replace;
        }

        public String K_Name
        {
            get { return Name; }
            set { Name = value; }
        }
        public String K_Replace
        {
            get { return Replace; }
            set { Replace = value; }
        }
    }
}
