using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace RePrinter3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Select();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string materialNum = textBox1.Text;
            string batchNum = textBox2.Text.ToUpper();
            int quantity = 0;
            int firstDisc; int lastDisc; string materialCode = ""; string shade = ""; string diameter = ""; string thickness = ""; string brand = ""; string shadeCopy = ""; string barcodeIdentifier = "";
            string specListLocation = "G:/LabX_Export/LabsQ_LabX_Integration/";
            File.Copy(specListLocation + "Wieland_Ardent_List_Separate.csv", specListLocation + "Wieland_Ardent_List_Tempcopy.csv",true);
            var column1 = new List<string>();
            var column2 = new List<string>();
            var column3 = new List<string>();
            var column4 = new List<string>();
            var column5 = new List<string>();
            var column6 = new List<string>();
            var column7 = new List<string>();
            var column8 = new List<string>();
            var column9 = new List<string>();
            var column10 = new List<string>();
            var column11 = new List<string>();
            var column12 = new List<string>();
            using (var rd = new StreamReader(specListLocation + "Wieland_Ardent_List_Tempcopy.csv"))
            {
                while (!rd.EndOfStream)
                {
                    string theNextLine = rd.ReadLine();
                    if (theNextLine == "") break;
                    var splits = theNextLine.Split(',');
                    column1.Add(splits[0]);
                    column2.Add(splits[1]);
                    column3.Add(splits[2]);
                    column4.Add(splits[3]);
                    column5.Add(splits[4]);
                    column6.Add(splits[5]);
                    column7.Add(splits[6]);
                    column8.Add(splits[2]);
                    column9.Add(splits[3]);
                    column10.Add(splits[4]);
                    column11.Add(splits[5]);
                    column12.Add(splits[6]);
                }
            }
            File.Delete(specListLocation + "Wieland_Ardent_List_Tempcopy.csv");
            for (int i = 0; i < column1.Count; i++)
            {
                if (column1[i] == materialNum)
                {
                    materialCode = column2[i];
                    shade = column6[i];
                    shadeCopy = shade;
                    diameter = column12[i];
                    thickness = column3[i];
                    brand = column5[i];
                    barcodeIdentifier = column4[i].Replace('_',' ');
                    if (shadeCopy == "Sun-Chroma") shadeCopy = shadeCopy.Replace("-", " ");
                    if (Char.IsDigit(shadeCopy[1])) shadeCopy = shadeCopy.Insert(1, " ");
                    else if (Char.IsDigit(shadeCopy[2])) shadeCopy = shadeCopy.Insert(2, " ");
                    if (materialCode == "ZTR" || shadeCopy == "Sun" || shadeCopy == "Sun Chroma") shadeCopy = shadeCopy.ToLower();
                }
            }
            
            if (shade.Equals(""))
            { MessageBox.Show("Material Not Found in Wieland_Ardent_List", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
            string labelInfoLocation = "G:/Equipment/GiS-Topex RFID Printer/Generated RFID Files/" + shade + "/" + batchNum + "/";


            //Topex Printer
            if(radioButton1.Checked==true)  
            {
                if(textBox3.Text == "")
                {
                    MessageBox.Show("Please enter range of discs to print", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); 
                    return;
                }
                if (textBox4.Text == "") textBox4.Text = textBox3.Text;
                try
                { firstDisc = int.Parse(textBox3.Text); lastDisc = int.Parse(textBox4.Text); }
                catch
                { MessageBox.Show("Disc range not numbers", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
                if (firstDisc > lastDisc)
                { MessageBox.Show("First disc larger than last disc", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
                
                //generate new rfid file
                string[] lines;
                try { lines = File.ReadAllLines(labelInfoLocation + batchNum + "_rfid_complete.csv"); }
                catch 
                { 
                    MessageBox.Show(labelInfoLocation + batchNum + "_rfid_complete.csv\nhas not been measured.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                                
                File.WriteAllText(labelInfoLocation + batchNum + "_rfid_temp.csv",lines[0]+"\n");
                if (lastDisc > lines.Length-1)
                {
                    MessageBox.Show("Higher number exceeds batch size", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if(firstDisc<1)
                {
                    MessageBox.Show("No discs less than zero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                for (int i = 0; i <= lastDisc - firstDisc; i++)
                {
                    File.AppendAllText(labelInfoLocation + batchNum + "_rfid_temp.csv", lines[firstDisc + i] + "\n");
                }


                ProcessStartInfo theDruck = new ProcessStartInfo();
                theDruck.CreateNoWindow = true;
                theDruck.UseShellExecute = false;
                theDruck.WindowStyle = ProcessWindowStyle.Hidden;
                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";

                //print FC labels
                if(materialCode == "ZFC" && checkBox1.Checked == false)
                    theDruck.Arguments = " /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + labelInfoLocation + batchNum + "_rfid_temp.csv\"  /start /hidden";
                //print FC retain labels
                else if (materialCode == "ZFC" && checkBox1.Checked == true)
                    theDruck.Arguments = " /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring_Retain.txt\" /RFID Off  /B \"" + labelInfoLocation + batchNum + "_rfid_temp.csv\"  /start /hidden";
                //print wieland labels
                else if (checkBox1.Checked == false)
                    theDruck.Arguments = " /P \"G:\\Topex_Printer\\Wieland_CE0123.txt\" /RFID On  /B \"" + labelInfoLocation + batchNum + "_rfid_temp.csv\"  /start /hidden";
                //print wieland retains labels
                else
                    theDruck.Arguments = " /P \"G:\\Topex_Printer\\Wieland Label Templates\\Wieland_CE0123_Retain.txt\" /RFID On  /B \"" + labelInfoLocation + batchNum + "_rfid_temp.csv\"  /start /hidden";

                //run druckerstation
                try
                {
                    using (Process exeProcess = Process.Start(theDruck))
                    {
                        exeProcess.WaitForExit();
                    }
                }
                catch
                {
                    MessageBox.Show("Druckerstation.exe not found or misconfigured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                //delete temporary file
                File.Delete(labelInfoLocation + batchNum + "_rfid_temp.csv");
            }
            


            //Box & Barcode Label
            else if (radioButton2.Checked==true || radioButton3.Checked == true)
            {
                if (textBox5.Text=="")
                {
                    MessageBox.Show("Please enter a quantity", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return;
                }
                try
                { quantity = int.Parse(textBox5.Text); }
                catch
                { MessageBox.Show("Quantity not a number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
                if (quantity < 1)
                { MessageBox.Show("Please enter a valid quantity", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }

                string labelName = "";
                if (radioButton2.Checked == true)
                    labelName = "Box Output_" + batchNum + ".cmd";
                else if (radioButton3.Checked == true)
                    labelName = "Barcode Output_" + batchNum + ".cmd";
                string lines = "";
                string cmdOutputLocation = @"G:\Equipment\GiS-Topex RFID Printer\Datamax Template Files\";
                try { lines = File.ReadAllText(labelInfoLocation + labelName); }
                catch
                {
                    DialogResult question = MessageBox.Show("Warning!\n" + labelInfoLocation + labelName + "\nhas not been measured.\nPrint anyway?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation); ;
                    if (question == DialogResult.No) return;
                    if (question == DialogResult.Yes)
                    {
                        if (radioButton2.Checked == true)
                            lines = "LABELNAME = \"G:\\Equipment\\GiS-Topex RFID Printer\\Datamax Template Files\\Zenostar Box Label Template.lab\"" + Environment.NewLine +
                                    "PRINTER= \"Copy of Datamax-O'Neil I-4606e Mark II,USB004\"" + Environment.NewLine +
                                    "Thickness = \"" + thickness + "\"" + Environment.NewLine +
                                    "Material Number = \"" + materialNum + "\"" + Environment.NewLine +
                                    "Lot Number = \"" + batchNum + "\"" + Environment.NewLine +
                                    "Material Description = \"" + shadeCopy + "\"" + Environment.NewLine +
                                    "LABELQUANTITY = \"" + quantity + "\"";
                        else if (radioButton3.Checked == true)
                            lines =
                                "LABELNAME = \"G:\\Equipment\\GiS-Topex RFID Printer\\Datamax Template Files\\Zenostar Barcode Label Template.lab\"" + Environment.NewLine +
                                "PRINTER = \"Datamax-O'Neil I-4606e Mark II,USB003\"" + Environment.NewLine +
                                "Material Number = \"" + materialNum + "\"" + Environment.NewLine +
                                "Lot Number = \"" + batchNum + "\"" + Environment.NewLine +
                                "Piece Count = 1 pc.\n" + Environment.NewLine +
                                "Barcode Label Identifier = \"" + barcodeIdentifier + " \"" + Environment.NewLine +
                                "LABELQUANTITY = \"" + quantity + "\"";

                    }
                }
                
                
                lines = lines.Replace("LABELQUANTITY = \"2\"", "LABELQUANTITY = \""+quantity+"\"");
                File.WriteAllText(cmdOutputLocation+"Label Output/Reprint Output.cmd",lines);

                //run CS.exe
                ProcessStartInfo csStart = new ProcessStartInfo();
                csStart.CreateNoWindow = true;
                csStart.UseShellExecute = false;
                csStart.WindowStyle = ProcessWindowStyle.Hidden;
                csStart.FileName = @"C:\Program Files (x86)\Teklynx\CODESOFT 2014\CS.exe";
                csStart.Arguments = " /CMD " + cmdOutputLocation + "Label Output";

                try
                {
                    using (Process exeProcess = Process.Start(csStart))
                        exeProcess.WaitForExit(1);
                }
                catch
                {
                    File.Delete(cmdOutputLocation + "Label Output/Reprint Output.cmd");
                    MessageBox.Show("CS.exe not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                Thread.Sleep(5000);
                MessageBox.Show(new Form() { TopMost = true },"You can close CodeSoft when labels finish printing", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            else
            { MessageBox.Show("No Label Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
            
        }




    }
}
