/*
Created by: John Walker
Last modified: 1/31/2018
Description: FORM3.CS is the form for the Server List Manager
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
using System.Xml;
using System.IO;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;

namespace AdminRDP
{
    public partial class Form3 : Form
    {
        private StreamWriter sr;

        public Form3()
        {
            InitializeComponent();
            
            //Initilize event handlers for drag and drop
            treeView1.ItemDrag += new ItemDragEventHandler(treeView1_ItemDrag);
            treeView1.DragEnter += new DragEventHandler(treeView1_DragEnter);
            treeView1.DragOver += new DragEventHandler(treeView1_DragOver);
            treeView1.DragDrop += new DragEventHandler(treeView1_DragDrop);    
        }

        //loop through nodes if nodes count is greater than 0
        private void LoopComboBox (TreeNodeCollection tnc)
        {
            foreach (TreeNode tn2 in tnc)
            {
                if (tn2.Nodes.Count > 0)
                {
                    comboBox1.Items.Add(tn2.Text);
                    LoopComboBox(tn2.Nodes);
                }
                else
                {
                    comboBox1.Items.Add(tn2.Text);
                    LoopComboBox(tn2.Nodes);
                }
            }
        }

        //load combo box and loop through treeview 
        //to display nodes
        private void LoadComboBox()
        {
            comboBox1.Items.Clear();
            comboBox1.Refresh();

            foreach (TreeNode tn in treeView1.Nodes)
            {              
                LoopComboBox(tn.Nodes);
            }

            if(comboBox1.Items.Count != 0)
            {
                comboBox1.SelectedIndex = 0;
            }
            
        }

        //export treeview to xml
        public void exportToXml(System.Windows.Forms.TreeView tv, string filename)
        {
            sr = new StreamWriter(filename, false, System.Text.Encoding.UTF8);
            //Write the header
            sr.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            //Write our root node
            sr.WriteLine("<" + treeView1.Nodes[0].Text + ">");
            foreach (TreeNode node in tv.Nodes)
            {
                saveNode(node.Nodes);
            }
            //Close the root node
            sr.WriteLine("</" + treeView1.Nodes[0].Text + ">");
            sr.Close();
        }

        //loop through child nodes
        private void saveNode(TreeNodeCollection tnc)
        {
            foreach (TreeNode node in tnc)
            {
                //If we have child nodes, we'll write 
                //a parent node, then iterrate through
                //the children
                if (node.Nodes.Count > 0)
                {
                    sr.WriteLine("<" + node.Text + ">");
                    saveNode(node.Nodes);
                    sr.WriteLine("</" + node.Text + ">");
                }
                else //No child nodes, so we just write the text
                    sr.WriteLine("<" + node.Text + "/>");
            }
        }


        //load xml document on page load
        //and populate treeview
        private void Form3_Load(object sender, EventArgs e)
        {
            try
            {
                //declare xml doc
                XmlDocument xmldoc = new XmlDocument();
                XmlNode xmlnode;

                //load xml file 
                using (FileStream fs = new FileStream(@"C:\Program Files\AdminRDP\SavedServers\ServerList.xml", FileMode.Open, FileAccess.Read))
                {
                    xmldoc.Load(fs);
                    xmlnode = xmldoc.ChildNodes[1];
                    //populate treeview from xml file
                    treeView1.Nodes.Clear();
                    treeView1.Nodes.Add(new TreeNode(xmldoc.DocumentElement.Name));
                    TreeNode tNode;
                    tNode = treeView1.Nodes[0];
                    AddNode(xmlnode, tNode);

                }
                treeView1.ExpandAll();

            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message);
            }

            LoadComboBox();
        }

        //function to populate treeview control from XML file
        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList nodeList;
            int i = 0;
            if (inXmlNode.HasChildNodes)
            {
                nodeList = inXmlNode.ChildNodes;
                for (i = 0; i <= nodeList.Count - 1; i++)
                {
                    xNode = inXmlNode.ChildNodes[i];
                    inTreeNode.Nodes.Add(new TreeNode(xNode.Name));
                    tNode = inTreeNode.Nodes[i];
                    AddNode(xNode, tNode);
                }
            }
        }

        //Add SERVER or group from user input
        private void button2_Click(object sender, EventArgs e)
        {
                //grab selected node
               TreeNode treeNode = treeView1.SelectedNode;

                //If no node is selected select ROOT
               if (treeNode == null)
                  {
                    treeNode = treeView1.Nodes[0];
                  }

               string inputBox = Interaction.InputBox("Please enter server/group name:");

            try
            {
                if (inputBox.Length > 0)
                {
                    if (ValidateString.IsValid(inputBox) == true)
                    {
                        treeNode.Nodes.Add(inputBox.Replace(" ", String.Empty));
                        treeNode.ExpandAll();
                        MessageBox.Show("Entries added!");
                    }
                    else
                    {
                        MessageBox.Show("Servers or groups cannot contain any special characters or start with numbers!");
                    }
                }
            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
            }
        

        //Remove server or server group
        private void button3_Click(object sender, EventArgs e)
        {
            //check if ROOT is attempted to be removed
            if (treeView1.SelectedNode.Text == "Root")
            {
                MessageBox.Show("Cannot remove ROOT group.");
            }
            else
            {
                treeView1.SelectedNode.Remove();
            }

        }

        //save treeview to xml button
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                exportToXml(treeView1, @"C:\Program Files\AdminRDP\SavedServers\ServerList.xml");
                MessageBox.Show("Changes saved!");
            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
            
        }

        //Functions below handle drag and drop functionality
        private void treeView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            // Move the dragged node when the left mouse button is used.
            if (e.Button == MouseButtons.Left)
            {
                DoDragDrop(e.Item, DragDropEffects.Move);
            }
        }

        // Set the target drop effect to the effect 
        // specified in the ItemDrag event handler.
        private void treeView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.AllowedEffect;
        }

        // Select the node under the mouse pointer to indicate the 
        // expected drop location.
        private void treeView1_DragOver(object sender, DragEventArgs e)
        {
            // Retrieve the client coordinates of the mouse position.
            Point targetPoint = treeView1.PointToClient(new Point(e.X, e.Y));

            // Select the node at the mouse position.
            treeView1.SelectedNode = treeView1.GetNodeAt(targetPoint);
        }

        private void treeView1_DragDrop(object sender, DragEventArgs e)
        {
            // Retrieve the client coordinates of the drop location.
            Point targetPoint = treeView1.PointToClient(new Point(e.X, e.Y));

            // Retrieve the node at the drop location.
            TreeNode targetNode = treeView1.GetNodeAt(targetPoint);

            // Retrieve the node that was dragged.
            TreeNode draggedNode = (TreeNode)e.Data.GetData(typeof(TreeNode));

            // Confirm that the node at the drop location is not 
            // the dragged node or a descendant of the dragged node.
            if (!draggedNode.Equals(targetNode) && !ContainsNode(draggedNode, targetNode))
            {
                // If it is a move operation, remove the node from its current 
                // location and add it to the node at the drop location.
                if (e.Effect == DragDropEffects.Move)
                {
                    draggedNode.Remove();
                    targetNode.Nodes.Add(draggedNode);
                }

                // If it is a copy operation, clone the dragged node 
                // and add it to the node at the drop location.
                else if (e.Effect == DragDropEffects.Copy)
                {
                    targetNode.Nodes.Add((TreeNode)draggedNode.Clone());
                }

                // Expand the node at the location 
                // to show the dropped node.
                targetNode.Expand();
            }
        }

        // Determine whether one node is a parent 
        // or ancestor of a second node.
        private bool ContainsNode(TreeNode node1, TreeNode node2)
        {
            // Check the parent node of the second node.
            if (node2.Parent == null) return false;
            if (node2.Parent.Equals(node1)) return true;

            // If the parent node is not null or equal to the first node, 
            // call the ContainsNode method recursively using the parent of 
            // the second node.
            return ContainsNode(node1, node2.Parent);
        }


        //open file dialog box
        private void importText_Click(object sender, EventArgs e)
        {
            string strFileName = openFileDialog1.ShowDialog().ToString();
        }

        //import entries from text file
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                //import text file to text box
                string readLine = null;
                string filePath = openFileDialog1.FileName.ToString();
                System.IO.StreamReader reader = new System.IO.StreamReader(filePath);

                while ((readLine = reader.ReadLine()) != null)
                {
                    textBox1.Text += readLine + "\r\n";
                }

            }
            catch
            {
                MessageBox.Show("Import failed!");
            }

        }

        //renames selected node
        private void renameToolStripMenuItem_Click(object sender, EventArgs e)
        {
           //grab selected node
           TreeNode treeNode = treeView1.SelectedNode;

           string inputBox = Interaction.InputBox("Please enter new name:", "AdminRDP",treeNode.Text);

           if (inputBox.Length > 0)
              {
                    treeNode.Text = inputBox.ToString();
              }
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button3_Click(sender, e);
        }

        //recurse through nodes to find selected index from combo box
        private void RecursiveSearch(TreeNodeCollection tnc)
        {
            foreach (TreeNode tn in tnc)
            {
                if (tn.Nodes.Count > 0)
                {
                    if (tn.Text == comboBox1.SelectedItem.ToString())
                    {
                        treeView1.SelectedNode = tn;
                    }
                    RecursiveSearch(tn.Nodes);
                }
            }
        }

        //add servers from textbox to treeview
        private void addServers_Click(object sender, EventArgs e)
        {
            try
            {
                bool returnValue = true;
                //find selected node
                foreach (TreeNode tn in treeView1.Nodes)
                {
                    RecursiveSearch(tn.Nodes);
                }

                //set selected node for nodes to be added under
                TreeNode tn2 = new TreeNode();
                tn2 = treeView1.SelectedNode;

                //split textbox
                var textLines = textBox1.Text.Split('\n');

                //loop through lines to add
                foreach (var ln in textLines)
                {
                    if (ln != "")
                    {
                        if (ValidateString.IsValid(ln) == true)
                        {
                            tn2.Nodes.Add(ln.Replace(" ", String.Empty).ToString());
                        }
                        else
                        {
                            MessageBox.Show("Entries cannot contain any special characters or begin with numbers!");
                            returnValue = false;
                        }

                    }
                }

                if(returnValue != false)
                {
                    treeView1.ExpandAll();
                    MessageBox.Show("Entries added!");
                    textBox1.Clear();
                }

            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
            
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            LoadComboBox();
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(treeView1.SelectedNode.Text);
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                exportToXml(treeView1, @"C:\Program Files\AdminRDP\SavedServers\ServerList.xml");
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message);
            }

        }
    }
}
