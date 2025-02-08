using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TemplateTranslateApp
{
    public partial class Form1 : Form
    {
        string TemplateName="";
        string Config="";
        string Chamber = "";
        
        
        public Form1()
        {
            
            InitializeComponent();
            textBoxMachine.Enabled = false;
            textBoxDataType.Enabled = false;
        }

        private void buttonMannual_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";

            if (openFileDialog.ShowDialog()==DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                textBoxMannualPath.Text = filePath;
            }
        }

        private void buttonConfig_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Config = openFileDialog.FileName;
                string filePath = openFileDialog.FileName;
                textBoxConfig.Text = filePath;
            }
        }

        private void buttonSaveFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog()==DialogResult.OK)
            {
                string folderPath = folderBrowserDialog.SelectedPath;
                textBoxSaveFolder.Text = folderPath;
            }
        }

        private void buttonGetTemplate_Click(object sender, EventArgs e)
        {
            if (textBoxSaveFolder.Text == "")
            {
                MessageBox.Show("Save Path not Selected");
            }
            else
            {
                if (textBoxMachine.Text != "" && textBoxDataType.Text != "")
                {
                    TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
                    
                    if (textBoxMannualPath.Text != "")
                    {
                        if (File.Exists(textBoxMannualPath.Text))
                        {
                            if (textBoxDataType.Text == "RCAI" || textBoxDataType.Text == "DCMAI")
                            {
                                try { 
                                    RC_DCM_AITemplate.PrintTemplate(textBoxMannualPath.Text, textBoxSaveFolder.Text, TemplateName, textBoxDataType.Text);
                                    MessageBox.Show("Template Created");
                                }
                                catch (Exception E)
                                {
                                    if (E.Message.Contains("Error saving file"))
                                    {
                                        MessageBox.Show("Close the opened Excel file and try again");
                                    }
                                    else
                                    {
                                        MessageBox.Show(E.Message + "\r\n" + "Template Generate unsuccessfully, this problem maybe caused by Mannual file format error or wrong file chosen");
                                    }
                                }
                            }
                            else if (textBoxDataType.Text == "RCTI" || textBoxDataType.Text == "DCMTI")
                            {
                             
                                try
                                {
                                    TITemplate.PrintTemplate(textBoxMannualPath.Text, textBoxSaveFolder.Text, TemplateName, textBoxDataType.Text);
                                    MessageBox.Show("Template Created");
                                }
                                catch(Exception E)
                                {
                                    if (E.Message.Contains("Error saving file"))
                                    {
                                        MessageBox.Show("Close the opened Excel file and try again");
                                    }
                                    else
                                    {
                                        MessageBox.Show(E.Message + "\r\n" + "Template Generate unsuccessfully, this problem maybe caused by Mannual file format error or wrong file chosen");
                                    }
                                }
                            }

                            
                        }
                        else
                        {
                            MessageBox.Show("Manual txt file not selected or file does not exit in the selected path");
                        }
                    }
                    else
                    {
                        if (textBoxDataType.Text == "WHCAI" || textBoxDataType.Text == "WHCSAI")
                        {
                            try
                            {
                                WHCTemplate.PrintTemplate(textBoxSaveFolder.Text, TemplateName, textBoxMachine.Text, textBoxDataType.Text);
                                MessageBox.Show("Template Created");
                            }
                            catch (Exception E)
                            {

                                if (E.Message.Contains("Error saving file"))
                                {
                                    MessageBox.Show("Close the opened Excel file and try again");
                                }
                                else
                                {
                                    MessageBox.Show(E.Message + "\r\n" + "Template generate unsuccessfully, this problem maybe caused by Mannual file format error or wrong file chosen");
                                }
                            }
                            

                        }
                        else
                        {
                            MessageBox.Show("Manual txt file not selected or file does not exit in the selected path");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select Machine and Data Type");
                }
            }

            
        }

        private void buttonGetConfig_Click(object sender, EventArgs e)
        {
            bool isXP4 = false;
            bool isSynergis = false;

            try{
                
                if (Config != "")
                {
                    
                    if (textBoxMachine.Text != "" && textBoxDataType.Text != "")
                    {

                        if (textBoxMachine.Text == "XP4")
                        {
                            isXP4 = true;
                        }

                        if (textBoxMachine.Text == "Synergis")
                        {
                            isSynergis = true;
                        }

                        if (textBoxSaveFolder.Text!="")
                        {
                            if (File.Exists(textBoxSaveFolder.Text + "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx"))
                            {
                                if (textBoxDataType.Text == "RCAI" || textBoxDataType.Text == "DCMAI")
                                {
                                    int ChamberNum = int.Parse(Chamber);
                                    try
                                    {
                                        
                                        RC_DCM_AITemplate.AssignConfig(Config, textBoxSaveFolder.Text, TemplateName, isXP4, isSynergis, ChamberNum);
                                        MessageBox.Show("Config Assigned");
                                    }
                                    catch (Exception E)
                                    {
                                        if (E.Message.Contains("Error saving file"))
                                        {
                                            MessageBox.Show("Close the opened Excel file and try again");
                                        }
                                        else
                                        {
                                            MessageBox.Show(E.Message + "\r\n" + "Config Assignment unsuccessful, this problem maybe caused by Mannual file format error or wrong file chosen");
                                        }
                                        
                                        
                                    }

                                }
                                else if (textBoxDataType.Text == "RCTI" || textBoxDataType.Text == "DCMTI")
                                {
                                    int ChamberNum = int.Parse(Chamber);
                                    try
                                    {
                                        
                                        TITemplate.AssignConfig(Config, textBoxSaveFolder.Text, TemplateName, isSynergis, ChamberNum);
                                        MessageBox.Show("Config Assigned");
                                    }
                                    catch (Exception E)
                                    {
                                        if (E.Message.Contains("Error saving file"))
                                        {
                                            MessageBox.Show("Close the opened Excel file and try again");
                                        }
                                        else
                                        {
                                            MessageBox.Show(E.Message + "\r\n" + "Config Assignment unsuccessful, this problem maybe caused by Mannual file format error or wrong file chosen");
                                        }
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        WHCTemplate.AssignConfig(Config, textBoxSaveFolder.Text, TemplateName, textBoxDataType.Text);
                                        MessageBox.Show("Config Assigned");
                                    }
                                    catch (Exception E)
                                    {
                                        if (E.Message.Contains("Error saving file"))
                                        {
                                            MessageBox.Show("Close the opened Excel file and try again");
                                        }
                                        else
                                        {
                                            MessageBox.Show(E.Message + "\r\n" + "Config Assignment unsuccessful, this problem maybe caused by Mannual file format error or wrong file chosen");
                                        }
                                    }
                                }

                            }
                            else
                            {
                                MessageBox.Show("Template not exists,please get template first");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Save folder not selected");
                        }
 
                    }
                    else
                    {
                        MessageBox.Show("Machine and Data Type not Selected");
                    }
                }
                else
                {
                    MessageBox.Show("Config File not Selected");
                }
            }
            catch(FormatException f)
            {
                MessageBox.Show(f.Message+"\r\n"+ "Chamber not selected");
            }


            
        }

        private void wHCAIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP4";
            textBoxDataType.Text = "WHCAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCSAIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP4";
            textBoxDataType.Text = "WHCSAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCAIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP8";
            textBoxDataType.Text = "WHCAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCSAIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP8";
            textBoxDataType.Text = "WHCSAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCAIToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Intrepid";
            textBoxDataType.Text = "WHCAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCSAIToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Intrepid";
            textBoxDataType.Text = "WHCSAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCAIToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Synergis";
            textBoxDataType.Text = "WHCAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void wHCSAIToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Synergis";
            textBoxDataType.Text = "WHCSAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = false;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = false;
            buttonMannual.Enabled = false;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBoxConfig.Clear();
            textBoxDataType.Clear();
            textBoxMachine.Clear();
            textBoxMannualPath.Clear();
            textBoxSaveFolder.Clear();
            MessageBox.Show("All Inputs been cleared");
            Config = "";
            TemplateName = "";
            comboBoxChamber.Text = "";
            
        }

        private void rCnAIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP4";
            textBoxDataType.Text = "RCAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void rCnTIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP4";
            textBoxDataType.Text = "RCTI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void dCMnAIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP8";
            textBoxDataType.Text = "DCMAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void dCMnTIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "XP8";
            textBoxDataType.Text = "DCMTI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void rCnAIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Intrepid";
            textBoxDataType.Text = "RCAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void rCnTIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Intrepid";
            textBoxDataType.Text = "RCTI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void dCMnAIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Synergis";
            textBoxDataType.Text = "DCMAI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void dCMnAIToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            textBoxMachine.Text = "Synergis";
            textBoxDataType.Text = "DCMTI";
            comboBoxChamber.Text = "";
            Chamber = "";
            comboBoxChamber.Enabled = true;
            textBoxMannualPath.Text = "";
            textBoxMannualPath.Enabled = true;
            buttonMannual.Enabled = true;
            TemplateName = TemplateName = "\\" + textBoxMachine.Text + "_" + textBoxDataType.Text + ".xlsx";
        }

        private void comboBoxChamber_SelectedIndexChanged(object sender, EventArgs e)
        {
            Chamber = comboBoxChamber.Text;
        }
        private void comboBoxChamber_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

    }
}
