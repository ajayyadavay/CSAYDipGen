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
// Interop
using Word = Microsoft.Office.Interop.Word;

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
            TxtInWorkCompletion.Text = "";
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

            try
            {
                TxtFinalBill_GT_Nepali.Text = TxtFinalBill_GT.Text;
            }
            catch
            {

            }
            
        }

        private void TxtFY_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TxtFYNepali.Text = TxtFY.Text;
            }
            catch
            {

            }
        }

        private void TxtWorkCompletion_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TxtInWorkcompletionNepali.Text = TxtInWorkCompletion.Text;
            }
            catch
            {

            }
        }

        public (string Name, string Address) ParseContractor(string contractorFull)
        {
            if (string.IsNullOrWhiteSpace(contractorFull))
                return (string.Empty, string.Empty);

            int commaIndex = contractorFull.IndexOf(',');
            if (commaIndex >= 0)
            {
                string name = contractorFull.Substring(0, commaIndex).Trim();
                string address = contractorFull.Substring(commaIndex + 1).Trim();
                return (name, address);
            }
            else
            {
                // No comma -> only name
                return (contractorFull.Trim(), string.Empty);
            }
        }


        private Dictionary<string, string> GetReplacementDictionary()
        {
            double num;
            string num2words, contractorFull, Rank1ContractorColName;
            //string Rank1_Name, Rank1_Address, Rank1_Amount, Rank1_diff, Rank1_percent;
            var dict = new Dictionary<string, string>();
            CSAYNumToWord cnw = new CSAYNumToWord();
            
            
            //Estimate Preparation tippani
            dict["???Est_Date???"] = dataGridViewDate.Rows[3].Cells["ColDate"].Value?.ToString();
            dict["???Context1???"] = TxtContext1.Text;
            dict["???आ.व.???"] = TxtFYNepali.Text;
            dict["???Budget_Subhead???"] = dataGridViewDate.Rows[0].Cells["ColDate"].Value?.ToString();
            dict["<<PROJECT_NAME>>"] = TxtProjectName.Text;
            dict["???Est_Amount???"] = dataGridViewCC.Rows[3].Cells["ColOfficeEst"].Value?.ToString();

            num = Convert.ToDouble(dataGridViewCC.Rows[3].Cells["ColOfficeEst"].Value?.ToString());
            num2words = cnw.ConvertNumberToNepaliWord(num);
            dict["<<EST_NEPALI_WORDS>>"] = num2words;

            dict["<<LVL1_NAME_POSITION>>"] = TxtLvl1.Text;

            //Estimate checking tippani
            dict["???Est_Chk_Date???"] = dataGridViewDate.Rows[4].Cells["ColDate"].Value?.ToString();
            dict["???Context2???"] = TxtContext2.Text;
            dict["<<LVL2_NAME_POSITION>>"] = TxtLvl2.Text;

            //Dar rate letter_office
            contractorFull = dataGridViewCC.Rows[0].Cells["ColContractor1"].Value?.ToString();
            var (Contractor1_Name, Contractor1_Address) = ParseContractor(contractorFull);
            dict["???Date_of_DarRate1???"] = dataGridViewDate.Rows[5].Cells["ColDate"].Value?.ToString();
            dict["???Contractor1_Name???"] = Contractor1_Name;
            dict["???Contractor1_Address???"] = Contractor1_Address;
            dict["<<LVL3_NAME_POSITION>>"] = TxtLvl3.Text;

            contractorFull = dataGridViewCC.Rows[0].Cells["ColContractor2"].Value?.ToString();
            var (Contractor2_Name, Contractor2_Address) = ParseContractor(contractorFull);
            dict["???Date_of_DarRate2???"] = dataGridViewDate.Rows[6].Cells["ColDate"].Value?.ToString();
            dict["???Contractor2_Name???"] = Contractor2_Name;
            dict["???Contractor2_Address???"] = Contractor2_Address;

            contractorFull = dataGridViewCC.Rows[0].Cells["ColContractor3"].Value?.ToString();
            var (Contractor3_Name, Contractor3_Address) = ParseContractor(contractorFull);
            dict["???Date_of_DarRate3???"] = dataGridViewDate.Rows[7].Cells["ColDate"].Value?.ToString();
            dict["???Contractor3_Name???"] = Contractor3_Name;
            dict["???Contractor3_Address???"] = Contractor3_Address;

            //Dar rate letter_Contractor
            dict["???Date_of_DarRate11???"] = dataGridViewDate.Rows[8].Cells["ColDate"].Value?.ToString();
            dict["???Date_of_DarRate22???"] = dataGridViewDate.Rows[9].Cells["ColDate"].Value?.ToString();
            dict["???Date_of_DarRate33???"] = dataGridViewDate.Rows[10].Cells["ColDate"].Value?.ToString();

            //Comparative chart preparation
            dict["???Date_of_CCP???"] = dataGridViewDate.Rows[12].Cells["ColDate"].Value?.ToString();
            dict["???C1_Amount???"] = dataGridViewCC.Rows[3].Cells["ColContractor1"].Value?.ToString();
            dict["???C1_diff???"] = dataGridViewCC.Rows[4].Cells["ColContractor1"].Value?.ToString();
            dict["???C1_per???"] = dataGridViewCC.Rows[5].Cells["ColContractor1"].Value?.ToString();

            dict["???C2_Amount???"] = dataGridViewCC.Rows[3].Cells["ColContractor2"].Value?.ToString();
            dict["???C2_diff???"] = dataGridViewCC.Rows[4].Cells["ColContractor2"].Value?.ToString();
            dict["???C2_per???"] = dataGridViewCC.Rows[5].Cells["ColContractor2"].Value?.ToString();

            dict["???C3_Amount???"] = dataGridViewCC.Rows[3].Cells["ColContractor3"].Value?.ToString();
            dict["???C3_diff???"] = dataGridViewCC.Rows[4].Cells["ColContractor3"].Value?.ToString();
            dict["???C3_per???"] = dataGridViewCC.Rows[5].Cells["ColContractor3"].Value?.ToString();

            Rank1ContractorColName = RankOneContractorColumnName();
            contractorFull = dataGridViewCC.Rows[0].Cells[Rank1ContractorColName].Value?.ToString();
            var (R1_Name, R1_Address) = ParseContractor(contractorFull);
            dict["???R1_Name???"] = R1_Name;
            dict["???R1_Address???"] = R1_Address;

            dict["???R1_Amount???"] = dataGridViewCC.Rows[3].Cells[Rank1ContractorColName].Value?.ToString();
            dict["???R1_diff???"] = dataGridViewCC.Rows[4].Cells[Rank1ContractorColName].Value?.ToString();
            dict["???R1_per???"] = dataGridViewCC.Rows[5].Cells[Rank1ContractorColName].Value?.ToString();

            double temp_amount = Convert.ToDouble(dataGridViewCC.Rows[3].Cells[Rank1ContractorColName].Value?.ToString());
            string temp_words = cnw.ConvertNumberToNepaliWord(temp_amount);
            dict["???R1_words???"] = temp_words;


            //Comparative chart preparation
            dict["???Date_of_CCC???"] = dataGridViewDate.Rows[13].Cells["ColDate"].Value?.ToString();

            //letter requesting contractor to sign contract
            dict["???Date_of_ReqCtr???"] = dataGridViewDate.Rows[14].Cells["ColDate"].Value?.ToString();

            //letter from contractor to sign contract
            dict["???Date_of_CtrPrsnt???"] = dataGridViewDate.Rows[15].Cells["ColDate"].Value?.ToString();

            //Contract agreement
            dict["???Date_of_CtrSign???"] = dataGridViewDate.Rows[16].Cells["ColDate"].Value?.ToString();
            dict["???Date_of_WorkStart???"] = dataGridViewDate.Rows[17].Cells["ColDate"].Value?.ToString();
            dict["???Date_of_IntendWorkComplete???"] = TxtInWorkcompletionNepali.Text;

            //Work completion letter submitted by contractor
            dict["???Date_of_WorkComplete???"] = dataGridViewDate.Rows[18].Cells["ColDate"].Value?.ToString();

            //Tippani of bill preparation
            dict["???Bill_Date???"] = dataGridViewDate.Rows[20].Cells["ColDate"].Value?.ToString();
            dict["???Bill_Amount???"] = TxtFinalBill_GT_Nepali.Text;
            dict["???BILL_NEPALI_WORDS???"] = TxtFinalBillNepaliWords.Text;

            //Tippani of bill checking
            dict["???Bill_Chk_Date???"] = dataGridViewDate.Rows[21].Cells["ColDate"].Value?.ToString();


            return dict;
        }

        public string RankOneContractorColumnName()
        {
            int Rank1, Rank2;
            Rank1 = Convert.ToInt32(dataGridViewCC.Rows[6].Cells["ColContractor1"].Value?.ToString());
            Rank2 = Convert.ToInt32(dataGridViewCC.Rows[6].Cells["ColContractor2"].Value?.ToString());
            //Rank3 = Convert.ToInt32(dataGridViewCC.Rows[6].Cells["ColContractor3"].Value?.ToString());

            if (Rank1 == 1) return "ColContractor1";
            else if (Rank2 == 2) return "ColContractor2";
            else return "ColContractor3";

        }


        private void generateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Cur_Dir = Environment.CurrentDirectory;
            
            string templatePath = Cur_Dir + "\\DipGenFileFormat\\" + "DipGenFileFormat.docx";
            //string outputPath = Cur_Dir + "\\DipGenOutput\\" + "DipGenOutputFile.docx";

            string outputPath = "";
            SaveFileDialog savefiledialog1 = new SaveFileDialog();
            savefiledialog1.Filter = "Document (*.docx)|*.docx";//"Text File(*.txt)|*.txt|Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*";
            savefiledialog1.FilterIndex = 1;

            if (savefiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputPath = savefiledialog1.FileName;
            }
            else if (savefiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel) return;


            var replacements = GetReplacementDictionary();

            var wordApp = new Word.Application();
            Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(templatePath);

                foreach (var pair in replacements)
                {
                    Word.Find findObject = wordApp.Selection.Find;
                    findObject.ClearFormatting();
                    findObject.Text = pair.Key;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = pair.Value ?? "";

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    findObject.Execute(Replace: ref replaceAll);
                }

                doc.SaveAs2(outputPath);
                MessageBox.Show("Document created !");
            }
            finally
            {
                doc?.Close();
                wordApp.Quit();
            }

        }

        private void ShowAboutDialog()
        {
            Form aboutForm = new Form();
            aboutForm.Text = "About This Software";
            aboutForm.Size = new Size(600, 450);
            aboutForm.StartPosition = FormStartPosition.CenterScreen;
            aboutForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            aboutForm.MaximizeBox = false;
            aboutForm.MinimizeBox = false;
            aboutForm.BackColor = Color.WhiteSmoke;

            // Title
            Label titleLabel = new Label();
            titleLabel.Text = "CSAY DipGen Software Information";
            titleLabel.Font = new Font("Segoe UI", 16, FontStyle.Bold);
            titleLabel.ForeColor = Color.FromArgb(0, 120, 215);
            titleLabel.AutoSize = true;
            titleLabel.Location = new Point(120, 30);

            // Created by
            Label createdBy = new Label();
            createdBy.Text = "Created by: Ajay Yadav";
            createdBy.Font = new Font("Segoe UI", 11);
            createdBy.Location = new Point(40, 90);
            createdBy.AutoSize = true;

            // Version
            Label version = new Label();
            version.Text = "Version: 1.0.0 (2025)";
            version.Font = new Font("Segoe UI", 11);
            version.Location = new Point(40, 120);
            version.AutoSize = true;

            // Download link
            LinkLabel download = new LinkLabel();
            download.Text = "Download Software: https://github.com/ajayyadavay";
            download.Font = new Font("Segoe UI", 10);
            download.LinkColor = Color.Blue;
            download.Location = new Point(40, 150);
            download.AutoSize = true;
            download.LinkClicked += (s, e) => Process.Start(new ProcessStartInfo
            {
                FileName = "https://github.com/ajayyadavay",
                UseShellExecute = true
            });

            // Email
            LinkLabel email = new LinkLabel();
            email.Text = "Email: civil.ajayyadav@gmail.com";
            email.Font = new Font("Segoe UI", 10);
            email.LinkColor = Color.Blue;
            email.Location = new Point(40, 180);
            email.AutoSize = true;
            email.LinkClicked += (s, e) => Process.Start(new ProcessStartInfo
            {
                FileName = "mailto:civil.ajayyadav@gmail.com",
                UseShellExecute = true
            });

            // How to use section
            Label howToUse = new Label();
            howToUse.Text = "How to use?";
            howToUse.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            howToUse.ForeColor = Color.FromArgb(0, 90, 180);
            howToUse.Location = new Point(40, 220);
            howToUse.AutoSize = true;

            Label steps = new Label();
            steps.Text = "1. Open the software and enter all your data in text boxes and table.\n" +
                         "2. Save *.dip file to access later via Load *.dip (optional).\n" +
                         "3. Click File>Generate from template (*.docx).\n" +
                         "4. The file will be saved in selected location.";
            steps.Font = new Font("Segoe UI", 10);
            steps.Location = new Point(60, 250);
            steps.Size = new Size(480, 100);

            // OK Button
            Button okButton = new Button();
            okButton.Text = "OK";
            okButton.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            okButton.BackColor = Color.FromArgb(0, 120, 215);
            okButton.ForeColor = Color.White;
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.Size = new Size(100, 35);
            okButton.Location = new Point(aboutForm.ClientSize.Width / 2 - 50, 350);
            okButton.Click += (s, e) => aboutForm.Close();

            // Add controls
            aboutForm.Controls.Add(titleLabel);
            aboutForm.Controls.Add(createdBy);
            aboutForm.Controls.Add(version);
            aboutForm.Controls.Add(download);
            aboutForm.Controls.Add(email);
            aboutForm.Controls.Add(howToUse);
            aboutForm.Controls.Add(steps);
            aboutForm.Controls.Add(okButton);

            // Show as modal dialog
            aboutForm.ShowDialog();
        }


        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowAboutDialog();
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
            Work_Completion_date = TxtInWorkCompletion.Text;
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
                    TxtInWorkCompletion.Text = DipGenIn.Work_Completion_date_ser;
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
