using CSAY_ContractManagementSoftware;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSAY_DipGen
{
    public partial class FrmDipGen : Form
    {
        public string[] Date_Value = new string[50];
        public string[,] CC_Value = new string[10,50];
        string Project_Name, FY, Work_Completion_date, Final_Bill_GT;
        string Lvl1, Lvl2, Lvl3, Division, Lvl1_Context, Lvl2_Context;
        public FrmDipGen()
        {
            InitializeComponent();
        }

        

        private void FrmDipGen_Load(object sender, EventArgs e)
        {
            Generate_Date_Datagridview();
            Generate_CC_Datagridview();
        }

        private void newToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            TxtProjectName.Text = "";
            TxtFY.Text = "";
            TxtWorkCompletion.Text = "";
            TxtFinalBill_GT.Text = "";

            for (int i = 0; i < 22; i++)
            {
                dataGridViewDate.Rows[i].Cells[2].Value = "";
            }

            int[] idx_cc = { 2, 4, 6, 8 };

            for (int c = 0; c < 4; c++)
            {
                for (int r = 0; r < 7; r++)
                {
                    CC_Value[c, r] = "";
                }

            }

            for (int i = 0; i < 7; i++)
            {
                dataGridViewCC.Rows[i].Cells[2].Value = "";
                dataGridViewCC.Rows[i].Cells[4].Value = "";
                dataGridViewCC.Rows[i].Cells[6].Value = "";
                dataGridViewCC.Rows[i].Cells[8].Value = "";

            }

            TxtLvl1.Text = "";
            TxtLvl2.Text = "";
            TxtLvl3.Text = "";
            TxtContext1.Text = "";
            TxtContext2.Text = "";
            TxtDivision.Text = "";
        }

        private void TxtFinalBill_GT_TextChanged(object sender, EventArgs e)
        {
            double num;
            string num2words;
            CSAYNumToWord cnw = new CSAYNumToWord();
            if(TxtFinalBill_GT.Text != "")
            {
                num = Convert.ToDouble(TxtFinalBill_GT.Text);
                num2words = cnw.ConvertNumberToNepaliWord(num);
                TxtFinalBillNepaliWords.Text = num2words;
                num2words = cnw.ConvertNumberToEnglishWord(num);
                TxtFinalBillEnglishWords.Text = num2words;
            }
            else
            {
                num2words = "";
                TxtFinalBillNepaliWords.Text = num2words;
                TxtFinalBillEnglishWords.Text = num2words;
            }
            
        }

        public void Generate_CC_Datagridview()
        {
            //initialize and declared variables
            string[] Description = new string[]
            {
                "Contractor's Name and Address", "Subtotal", "VAT 13% of Subtotal",
                "GrandTotal = Subtotal + VAT ", "Amount below/above",
                "Percentage below/above", "Rank"
            };

            //generate rows in contract and bill datagrid
            for (int i = 0; i < 7; i++) //0 to 7
            {
                dataGridViewCC.Rows.Add();
                dataGridViewCC.Rows[i].Cells[0].Value = (i + 1).ToString();//Description of Estimate
                dataGridViewCC.Rows[i].Cells[1].Value = Description[i];//Description of Estimate
                dataGridViewCC.Columns[4].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridViewCC.Columns[6].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridViewCC.Columns[8].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridViewCC.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            }
        }
        
        public void Generate_Date_Datagridview()
        {

            //initialize and declared variables
            string[] Description = new string[]
            {
                "Budget Head no.", 
                "Request letter, if any", 
                "Estimate",
                "Tippani for Estimate preparation ", 
                "Tippani for Estimate checking",
                "दररेट पेश गर्ने सम्बन्धमा - 1", 
                "दररेट पेश गर्ने सम्बन्धमा - 2", 
                "दररेट पेश गर्ने सम्बन्धमा - 3",
                "दररेट पेश गरेको सम्बन्धमा - 1",
                "दररेट पेश गरेको सम्बन्धमा - 2",
                "दररेट पेश गरेको सम्बन्धमा - 3",
                "Comparative chart",
                "Tippani for comparative chart preparation",
                "Tippani for comparative chart checking",
                "सम्झौता गर्न आउने सम्बन्धमा",
                "सम्झौता गर्न आएको बारे",
                "सम्झौता पत्र",
                "कार्यादेश पत्र",
                "अन्तिम विल भुक्तानि सम्बन्धमा",
                "Bill",
                "Tippani for bill preparation",
                "Tippani for bill checking"
            };

            //generate rows in contract and bill datagrid
            for (int i = 0; i < 22; i++) //0 to 7
            {
                
                dataGridViewDate.Rows.Add();
                dataGridViewDate.Rows[i].Cells[0].Value = (i+1).ToString();//Description of Estimate
                dataGridViewDate.Rows[i].Cells[1].Value = Description[i];//Description of Estimate
                dataGridViewDate.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridViewDate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            }
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewCC.SelectedCells.Count < 1) return;

                string[] lines;

                int row = dataGridViewCC.SelectedCells[0].RowIndex;
                int col =dataGridViewCC.SelectedCells[0].ColumnIndex;

                //get copied values
                lines = Clipboard.GetText().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                string[] values;
                for (int i = 0; i < lines.Length; i++)
                {
                    values = lines[i].Split('\t');

                    if (row >= dataGridViewCC.Rows.Count || dataGridViewCC.Rows[row].IsNewRow) continue;
                    //if (row >= dataGridView1.Rows.Count || dataGridView1.Rows[row].IsNewRow) dataGridView1.Rows.Add();
                    for (int j = 0; j < values.Length; j++)
                    {
                        if (col + j >= dataGridViewCC.Columns.Count) continue;
                        dataGridViewCC.Rows[row].Cells[col + j].Value = values[j];
                    }

                    row++;
                }

            }
            catch
            {

            }
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewDate.SelectedCells.Count < 1) return;

                string[] lines;

                int row = dataGridViewDate.SelectedCells[0].RowIndex;
                int col = dataGridViewDate.SelectedCells[0].ColumnIndex;

                //get copied values
                lines = Clipboard.GetText().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                string[] values;
                for (int i = 0; i < lines.Length; i++)
                {
                    values = lines[i].Split('\t');

                    if (row >= dataGridViewDate.Rows.Count || dataGridViewDate.Rows[row].IsNewRow) continue;
                    //if (row >= dataGridView1.Rows.Count || dataGridView1.Rows[row].IsNewRow) dataGridView1.Rows.Add();
                    for (int j = 0; j < values.Length; j++)
                    {
                        if (col + j >= dataGridViewDate.Columns.Count) continue;
                        dataGridViewDate.Rows[row].Cells[col + j].Value = values[j];
                    }

                    row++;
                }

            }
            catch
            {

            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        public void Input_Data()
        {
            Project_Name = TxtProjectName.Text;
            FY = TxtFY.Text;
            Work_Completion_date = TxtWorkCompletion.Text;
            Final_Bill_GT = TxtFinalBill_GT.Text;

            for (int i = 0; i < 22; i++)
            {
                Date_Value[i] = dataGridViewDate.Rows[i].Cells[2].Value.ToString();
            }

            int[] idx_cc = {2, 4, 6, 8};

            for (int c = 0; c < 4; c++)
            {
                for(int r = 0; r<7; r++)
                {
                    CC_Value[c,r] = dataGridViewCC.Rows[r].Cells[idx_cc[c]].Value.ToString();
                }
                
            }

            Lvl1 = TxtLvl1.Text;
            Lvl2 = TxtLvl2.Text;
            Lvl3 = TxtLvl3.Text;
            Division = TxtDivision.Text;
            Lvl1_Context = TxtContext1.Text;
            Lvl2_Context = TxtContext2.Text;

        }

        private void savedipToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string[] cc_office = new string[10];
            string[] cc_contractor1 = new string[10];
            string[] cc_contractor2 = new string[10];
            string[] cc_contractor3 = new string[10]; 

            Input_Data();

            for(int i =0; i<7;i++)
            {
                cc_office[i] = CC_Value[0, i];
                cc_contractor1[i] = CC_Value[1, i];
                cc_contractor2[i] = CC_Value[2, i];
                cc_contractor3[i] = CC_Value[3, i];

            }
            DipGenInputClass.DipGenValueInfo DipGenIn = new DipGenInputClass.DipGenValueInfo
            {
                Project_Name_ser = Project_Name,
                FY_ser = FY,
                Work_Completion_date_ser = Work_Completion_date,
                Final_Bill_GT_ser = Final_Bill_GT,
                Date_value_ser = Date_Value,
                CC_Office_value_ser = cc_office,
                CC_Contractor1_value_ser = cc_contractor1,
                CC_Contractor2_value_ser = cc_contractor2,
                CC_Contractor3_value_ser = cc_contractor3,
                Lvl1_ser = Lvl1,
                Lvl2_ser = Lvl2,
                Lvl3_ser = Lvl3,
                Lvl1_Context_ser = Lvl1_Context,
                Lvl2_Context_ser = Lvl2_Context,
                Division_ser = Division

            };
               
        
            BinaryFormatter bf = new BinaryFormatter();

            string path = "";
            SaveFileDialog savefiledialog1 = new SaveFileDialog();
            savefiledialog1.Filter = "Direct Purchase (*.dip)|*.dip";//"Text File(*.txt)|*.txt|Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*";
            savefiledialog1.FilterIndex = 1;

            if (savefiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = savefiledialog1.FileName;
                //LoadTxtToDatagridview(dataGridView1, path, 1, 3);
            }
            else if (savefiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel) return;


            FileStream fsout = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            try
            {
                using (fsout)
                {
                    bf.Serialize(fsout, DipGenIn);
                    MessageBox.Show("File saved to\n" + path);
                }
            }
            catch
            {
                MessageBox.Show("Error saving (*.dip)...");
            }
        }



        private void loaddipToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DipGenInputClass.DipGenValueInfo DipGenIn = new DipGenInputClass.DipGenValueInfo();
            BinaryFormatter bf = new BinaryFormatter();

            string path = "";
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.Filter = "Direct Purchase (*.dip)|*.dip";//"Text File(*.txt)|*.txt|Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*";
            openfiledialog1.FilterIndex = 1;

            if (openfiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = openfiledialog1.FileName;
            }
            else if (openfiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel) return;


            FileStream fsin = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            try
            {
                using (fsin)
                {
                    DipGenIn = (DipGenInputClass.DipGenValueInfo)bf.Deserialize(fsin);

                    TxtProjectName.Text = DipGenIn.Project_Name_ser;
                    TxtFY.Text = DipGenIn.FY_ser;
                    TxtWorkCompletion.Text = DipGenIn.Work_Completion_date_ser;
                    TxtFinalBill_GT.Text = DipGenIn.Final_Bill_GT_ser;

                    for (int i = 0; i < 22; i++)
                    {
                        dataGridViewDate.Rows[i].Cells[2].Value = DipGenIn.Date_value_ser[i];
                    }

                    for (int i = 0; i < 7; i++)
                    {
                        dataGridViewCC.Rows[i].Cells[2].Value =DipGenIn.CC_Office_value_ser[i];
                        dataGridViewCC.Rows[i].Cells[4].Value = DipGenIn.CC_Contractor1_value_ser[i];
                        dataGridViewCC.Rows[i].Cells[6].Value = DipGenIn.CC_Contractor2_value_ser[i];
                        dataGridViewCC.Rows[i].Cells[8].Value = DipGenIn.CC_Contractor3_value_ser[i];

                    }

                    TxtLvl1.Text = DipGenIn.Lvl1_ser;
                    TxtLvl2.Text = DipGenIn.Lvl2_ser;
                    TxtLvl3.Text = DipGenIn.Lvl3_ser;
                    TxtContext1.Text = DipGenIn.Lvl1_Context_ser;
                    TxtContext2.Text = DipGenIn.Lvl2_Context_ser;
                    TxtDivision.Text = DipGenIn.Division_ser;

                    
                    MessageBox.Show("File loaded from \n" + path);
                }
            }
            catch
            {
                MessageBox.Show("Error loading...");
            }
        }








    }
}
