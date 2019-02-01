/*
Created by: John Walker
Last modified: 1/31/2018
Description: FORM2.CS is the form for the Credential Manager
*/


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO;

namespace AdminRDP
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public bool Save = false;

        private void Load_ComboBox()
        {
            //clear combox boxes for refresh everytime button is clicked
            comboBox1.Items.Clear();
            comboBox1.Refresh();
            try
            {
                //grab XML info from SavedCredentials folder
                string[] strXMLExists = Directory.GetFiles("C:\\Program Files\\AdminRDP\\SavedCredentials\\", "*.xml", SearchOption.AllDirectories);

                //populdate dropbox data from xml if there are xml files present
                if (strXMLExists != null)
                {
                    //grab XML data
                    XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));
                    string[] dir = Directory.GetFiles("C:\\Program Files\\AdminRDP\\SavedCredentials\\", "*.xml");

                    //loop through current directory and grab info from saved credential xml files
                    foreach (string d in dir)
                    {
                        using (FileStream read = new FileStream(d, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            //add items to combo box
                            Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);
                            comboBox1.Items.Add(credman.txtUsername + "_" + credman.txtDomain);
                            read.Close();
                        }

                    }

                }
            }catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        //save or update credentials
        private void button1_Click(object sender, EventArgs e)
        {
            //Form validation
            if (txtUsername.Text == "")
            {
                MessageBox.Show("Please enter a username!");
                return;
            }

            if (txtPassword.Text == "")
            {
                MessageBox.Show("Please enter a password!");
                return;
            }

            if (txtDomain.Text == "")
            {
                MessageBox.Show("Please enter a domain!");
                return;
            }

            //attempt to save credentials
            try
            {
                //save credentials to XML
                Credential_Manager_Save credman = new Credential_Manager_Save();
                credman.txtUsername = txtUsername.Text;
                credman.txtPassword = Encryption.EncryptString(txtPassword.Text);
                credman.txtDomain = txtDomain.Text;
                string strPath = "C:\\Program Files\\AdminRDP\\SavedCredentials\\" + txtUsername.Text + "_" + txtDomain.Text + ".xml";
                SaveXML.SaveData(credman, strPath);                
                MessageBox.Show("Credentials have been saved!");
                txtUsername.Text = "";
                txtPassword.Text = "";
                txtDomain.Text = "";
                Save = true;
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       
        private void Form2_Load(object sender, EventArgs e)
        {

            
        }

        //load data back in text fields when selected
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {           
            XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));
            using (FileStream read = new FileStream("C:\\Program Files\\AdminRDP\\SavedCredentials\\" + comboBox1.SelectedItem.ToString() + ".xml", FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);
                txtUsername.Text = credman.txtUsername;
                txtPassword.Text = Encryption.DecryptString(credman.txtPassword);
                txtDomain.Text = credman.txtDomain;
                read.Close();
            }          
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            Load_ComboBox();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //delete credentials
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Please select a username to delete!");
            }
            else
            {
                File.Delete("C:\\Program Files\\AdminRDP\\SavedCredentials\\" + comboBox1.SelectedItem.ToString() + ".xml");
                txtUsername.Text = "";
                txtPassword.Text = "";
                txtDomain.Text = "";
                MessageBox.Show("Deleted!");
            }

        }

        //Set default credentials
        private void button3_Click(object sender, EventArgs e)
        {
            //checks to see if combobox is selected
            if (comboBox1.SelectedItem != null)
            {
                try
                {
                    //set file path
                    string strFilePath = @"C:\Program Files\AdminRDP\SavedCredentials\Default.txt";

                    //if file doesn't exit create it.
                    if (!File.Exists(strFilePath))
                    {
                        //create file
                        using (var file = File.Create(strFilePath)){};
                    }

                    //write default username to text file
                    using (StreamWriter write = new StreamWriter(strFilePath))
                    {
                        write.WriteLine(comboBox1.SelectedItem.ToString());
                        write.Close();
                    }

                    MessageBox.Show(comboBox1.SelectedItem.ToString() + " has been set as the default credentials for login!");

                }
                catch(Exception e2)
                {
                    MessageBox.Show(e2.Message);
                }
            }
            else
            {
                MessageBox.Show("Please selected a profile!");
            }
            
        }
    }
}
