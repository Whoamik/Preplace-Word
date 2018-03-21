using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Novacode;
using System.IO;

namespace Word
{
    public partial class frmReplaceWord : Form
    {
        public frmReplaceWord()
        {
            InitializeComponent();
        }
        private void btnIn_Click(object sender, EventArgs e)
        {
            List<DataWord> sData = new List<DataWord>();
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
            word.Application wordApp = new word.Application();
            String StartupPath = System.Windows.Forms.Application.StartupPath + "\\VPB-TBLAN1.docx";
            if (!File.Exists(StartupPath))
            {
                MessageBox.Show("Không tìm thấy file Template", "Thông Báo");
                return;
            }
            else
            {
                Document newDocument = wordApp.Documents.Open(StartupPath);
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
                    newDocument.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Temp (1).docx");
                    newDocument.Close();
                }
            }
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
