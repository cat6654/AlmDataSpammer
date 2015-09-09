using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TDAPIOLELib;
using SACLIENTLib;
using System.Web;
using System.Net;
using System.IO;
using System.Xml;
using System.Threading;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;


namespace WindowsFormsApplication5
{

    public partial class Form1 : Form
    {
        int nCount;
        int fCount;
        int rCount;
        int cCount;
        int tsCount;
        int dCount;
        int pCount;
        int liCount;
        int usCount;
        int ugCount;
        int hCount;
        int reqParent;
        int tsetParent;
        int rest_bCount;
        string server_;
        string user_;
        string password_;
        string domain_;
        string project;
        int DB_TYPE;
        string DB_SERVER_NAME;
        string TABLESPACE_;
        string TEMP_TABLESPACE;
        string db_user_;
        string db_password_;
        string list_name;
        string item_name;
        string user_group_list_name;
        string user_name;
        string userGroup_name;
        string testParent;
        int foldername;
        int tsfoldername;
        int rname;
        int relname;
        int dname;
        int pname;
        List<int> test_l1 = new List<int>();
        List<int> req_l1 = new List<int>();
        List<int> cycle_l1 = new List<int>();
        List<string> domain_l1 = new List<string>();
        List<int> req_l2 = new List<int>();
        List<int> test_l2 = new List<int>();
        List<int> tset_l1 = new List<int>();
        List<int> test_l3 = new List<int>();
        
        static private TDConnection connection_ = new TDConnection();
        static private SAapi sconnection = new SAapi();
        static private rest1 REST = new rest1();
        static private xml_parser XPARSER = new xml_parser();


        public List<TDField> reqFieldsList = new List<TDField>();
        public List<CustomizationField> reqFieldsListreqs = new List<CustomizationField>();
        public List<string> GetAllNodesNames(CustomizationListNode rootNode, bool isRoot, [Optional, DefaultParameterValue(null)]  List<string> nodesNames)
        {
            var names = new List<string>();
            if (nodesNames != null)
                names = nodesNames;
            if (!isRoot)
                names.Add(rootNode.Name);
            if (rootNode.ChildrenCount > 0)
            {
                foreach (CustomizationListNode curNode in rootNode.Children)
                {
                    names = GetAllNodesNames(curNode, false, names);
                }
            }

            return names;
        }



        public Form1()
        {

            try {
                
            InitializeComponent();
            listBox1.Items.Add(String.Format("{0} {1}", "ID ", "Name"));
            /*    //Write config to txt example
            string log = @"C:\ALM_DATA_SPAMMER_log.txt"; oh yeah
            using (FileStream fs = new FileStream(log, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.Write(textBox2.Text + sw.NewLine);
                    sw.Write(textBox3.Text + sw.NewLine);
                    sw.Write(textBox4.Text + sw.NewLine);
                    sw.Write(textBox5.Text + sw.NewLine);
                    sw.Write(textBox6.Text + sw.NewLine);
                }
            }

                */       
            //Write config to xml 
           
           // XmlTextWriter writer = new XmlTextWriter(@"C:\ALM_DATA_SPAMMER_config.xml", System.Text.Encoding.UTF8);
            if (System.IO.File.Exists(@"C:\ALM_DATA_SPAMMER_config.xml"))
            {
                //XmlDocument doc = new XmlDocument();
                //doc.LoadXml(@"C:\ALM_DATA_SPAMMER_config.xml");
                
               
                //XmlReader reader = (XmlReader)XmlReader.Create(new StringReader(doc.ToString()));
                XmlTextReader reader = new XmlTextReader(@"C:\ALM_DATA_SPAMMER_config.xml");
                
                while (reader.Read())
               {
                
                    reader.ReadToFollowing("URL");
                    textBox2.Text = reader.ReadElementContentAsString();
                    reader.ReadToFollowing("DOMAIN");
                    textBox3.Text = reader.ReadElementContentAsString();
                    reader.ReadToFollowing("Project");
                    textBox4.Text = reader.ReadElementContentAsString();
                    reader.ReadToFollowing("user");
                    textBox5.Text = reader.ReadElementContentAsString();
                    reader.ReadToFollowing("password");
                    textBox6.Text = reader.ReadElementContentAsString();
                    reader.Close();
               }
            }
            

            progressBar1.Enabled = false;
            progressBar1.Visible = false;
            textBox14.Enabled = false;
            textBox16.Enabled = false;
            } 
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show("register both OTA API and SA API *.dlls under your ALM client's %appdata%/HP/ALM-Client/<your server> dir");
            }

                  }

        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Enabled = true;
            progressBar1.Visible = true;
                      
            nCount = Convert.ToInt32(textBox1.Text);
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = nCount + 1;


            /*
            BugFactory bf = (BugFactory)connection_.BugFactory;

            Random rand = new Random();
            rand.Next();

            string[] arr = new string[] { "New", "Open", "Fixed", "Closed", "Reopen", "Rejected" };
            string[] arr2 = new string[] { "1-Low", "2-Medium", "3-High", "4-Very High", "5-Urgent" };

            for (int i = 0; i < nCount; ++i)
            {
                ++progressBar1.Value;
                Bug bg = (Bug)bf.AddItem(DBNull.Value);
               
                bg.Summary = "Defect_" + i.ToString() + Convert.ToString(rand.Next());
                bg["BG_SEVERITY"] = arr2[i % arr2.Length];
                bg["BG_DETECTION_DATE"] = "2010-11-10";
                //arr[i % arr.Length] adding a comment for ALi commit change
                bg["BG_STATUS"] = arr[i % arr.Length];
                bg.Post();

                listBox1.Items.Add(String.Format("{0} {1}", bg.ID, bg.Summary));

                //Console.WriteLine(bg.ID + " " + bg.Summary); 

            }
             */
           
            CreateEntities myBug = new CreateEntities(connection_);
            Random rand = new Random();
            rand.Next();
            for (int i = 0; i < nCount; ++i)
            {
                myBug.DefectSummary = "Defect " + i + rand.Next();
                myBug.PostDefect();

                listBox1.Items.Add(String.Format("{0} {1}", myBug.DefectId, myBug.DefectSummary));

            }
            progressBar1.Visible = false;
        }

        

        private void button2_Click(object sender, EventArgs e)
        {
          try
            {
            server_ = textBox2.Text;
            user_ = textBox5.Text;
            password_ = textBox6.Text;
            domain_ = textBox3.Text;
            project = textBox4.Text;
            
                connection_.InitConnectionEx(server_);
                connection_.Login(user_, password_);
                connection_.Connect(domain_, project);


                if (connection_.Connected)
                {
                    MessageBox.Show("Connected");
                    button1.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    button5.Enabled = true;
                    button6.Enabled = true;
                    button7.Enabled = true;
                    button8.Enabled = true;
                    button13.Enabled = true;
                    button2.Enabled = false;
                    button23.Enabled = true;
                    button24.Enabled = true;
                    button25.Enabled = true;
                    button26.Enabled = true;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (connection_.Connected) connection_.Disconnect();
            if (connection_.LoggedIn)
            {
                connection_.Logout();
                MessageBox.Show("disconnected ;)");
                button3.Enabled = false;
                button4.Enabled = false;
                button1.Enabled = false;
                
                button5.Enabled = false;
                button6.Enabled = false;
                button7.Enabled = false;
                button8.Enabled = false;
                button23.Enabled = false;
                button24.Enabled = false;
                button25.Enabled = false;
                button26.Enabled = false;

                button2.Enabled = true;

                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;

                test_l3.Clear();

                listBox1.ResetText();
                listBox1.Select();
                listBox1.ClearSelected();
            }
            
            connection_.ReleaseConnection();
        }

      

        private void button5_Click(object sender, EventArgs e)
        {
            fCount = Convert.ToInt32(textBox7.Text);
            hCount = Convert.ToInt32(textBox23.Text);

            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = fCount + 1;

            Random rand = new Random();
            foldername = rand.Next();
            TreeManager tm = (TreeManager)connection_.TreeManager;

            Customization customization = connection_.Customization;
            

            try
            {
                if (checkBox4.Checked == false)
                {
                    for (int b = 0; b < hCount; ++b)
                    {

                        if (b == 0)
                        {
                            SubjectNode sn = (SubjectNode)tm.get_NodeByPath("Subject");
                            SubjectNode folder = (SubjectNode)sn.AddNode("testFolder" + foldername);
                            folder.Post();
                            test_l2.Add(folder.NodeID);
                            test_l3.Add(folder.NodeID);
                            testParent = folder.Path;
                            listBox1.Items.Add(String.Format("{0} {1}", folder.NodeID, folder.Name));
                        }
                        else
                        {
                            SubjectNode sn = (SubjectNode)tm.get_NodeByPath(testParent);
                            SubjectNode folder = (SubjectNode)sn.AddNode("testFolder" + foldername);
                            folder.Post();
                            test_l2.Add(folder.NodeID);
                            test_l3.Add(folder.NodeID);
                            testParent = folder.Path;
                            listBox1.Items.Add(String.Format("{0} {1}", folder.NodeID, folder.Name));
                        }

                    }
                }

                for (int i = 0; i < fCount; ++i)
                {

                    ++progressBar1.Value;
                 
                   // Creating Dependant Entity From Stateless Factory
                    TestFactory testFactory = (TestFactory)connection_.TestFactory;
                    Test test = (Test)testFactory.AddItem(DBNull.Value);
                    test.Name = "TEST_" + i.ToString() + foldername;
                    test.Type = "MANUAL";
                    //test["TS_SUBJECT"] = folder.NodeID;
                    if (checkBox4.Checked)
                    {
                        test["TS_SUBJECT"] = test_l3[i % test_l3.Count];
                    }
                    else
                    {
                        test["TS_SUBJECT"] = test_l2[i % test_l2.Count];
                    }
                    
                    //customization required fields start

                    foreach (TDField field in testFactory.Fields)
                    {
                        if (field.Property.IsRequired())
                        {
                            reqFieldsList.Add(field);
                        }
                    }       

            foreach (TDField field in reqFieldsList)
            {
                switch (field.Type)
                {
                    case 0:
                    case 1:
                    case 2: //Int,int,float
                        test[field.Name] = i;
                        break;

                    case 3: //string ["TS_NAME"]
                        if(!field.Name.Equals("TS_TYPE"))
                        {
                        if (customization.Fields.Field["Test", field.Name].RootId != null)
                        {
                            CustomizationListNode rootNode = customization.Fields.Field["Test", field.Name].List.RootNode;
                            
                            List<string> nodesNamesList = GetAllNodesNames(rootNode, true);
                            string[] nodes = nodesNamesList.ToArray();
                            test[field.Name] = nodes[rand.Next(0, nodes.Length - 1)];

                            
                        }
                        else
                        {
                           
                            test[field.Name] = field.Name + i.ToString() + foldername;
                        }
                        }
                        break;

                    case 4: //memo
                        test[field.Name] = field.Name + i.ToString() + foldername;
                        break;

                    case 5: //Date
                        test[field.Name] = DateTime.Now.ToString("MM-dd-yyyy");
                        break;

                    case 8: // List of users
                        test[field.Name] = connection_.UserName;
                        break;
                }
            }
          
                                           
                    //end

                    test.Post();

                    test_l1.Add(test.ID);

                    DesignStepFactory stepf = (DesignStepFactory)test.DesignStepFactory;
                    DesignStep my_step = (DesignStep)stepf.AddItem(DBNull.Value);
                    my_step.StepName = "Step Name " + i.ToString();
                    my_step["DS_DESCRIPTION"] = "BLABLABLA_DESCRIPTION" + foldername + "BLABLABLA :)";
                    my_step["DS_EXPECTED"] = "EXPECTED RESULT IS SOOOOO EXPECTED... Um... you know like... Long cat is SOOOOOOOOOOOOOOOOO LOOOOOOOOOOOOOONG" + foldername;
                    my_step.Post();

                    listBox1.Items.Add(String.Format("{0} {1}", test.ID, test.Name));

                }
                progressBar1.Visible = false;
                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
                checkBox4.Enabled = true;
                test_l2.Clear();
                reqFieldsList.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                reqFieldsList.Clear();
            }

            }
        


        private void button6_Click(object sender, EventArgs e)
        {
            Random rand2 = new Random();
            rname = rand2.Next();
            ReqFactory rf = (ReqFactory)connection_.ReqFactory;
            rCount = Convert.ToInt32(textBox8.Text);
            hCount = Convert.ToInt32(textBox23.Text);

            Customization customization = connection_.Customization;
            
            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = rCount + 1;

            //string[] arr2 = new string[] { "Undefined", "Functional", "Group", "Folder", "Testing", "Business", "Performance" };
            int[] arr2 = new int[] {0,3};
            try
            {
                for (int b = 0; b < hCount; ++b)
                {

                    Req rqf = (Req)rf.AddItem(DBNull.Value);
                    rqf.Name = "Requirement_Folder_" + rname; //for alm 11 or alm 11.50
                    rqf["RQ_REQ_NAME"] = "Requirement_Folder_" + rname; //for qc10
                    rqf["RQ_TYPE_ID"] = 1;
                    rqf.ParentId = reqParent;

                   

                    rqf.Post();

                    listBox1.Items.Add(String.Format("{0} {1}", rqf.ID, rqf.Name));

                    reqParent = rqf.ID;
                    req_l2.Add(rqf.ID);
                }

                //pizdec nachinaetsya tut...
                CustomizationTypes ct = (CustomizationTypes)customization.Types;
                //CustomizationReqType crt = (CustomizationReqType)ct.GetEntityCustomizationType(0,0);

                List reqTypes = (List)ct.GetEntityCustomizationTypes(0);

                for (int i = 0; i < 4; i=i+3)
                {
                    CustomizationReqType reqType2 = (CustomizationReqType)ct.GetEntityCustomizationType(0, i);

                    foreach (CustomizationTypedField reqTypeField in reqType2.Fields)
                    {
                        CustomizationField currentField = (CustomizationField)reqTypeField.Field;
                        if (reqTypeField.IsRequired)
                        {
                            reqFieldsListreqs.Add(currentField);
                           // listBox1.Items.Add(String.Format("{0} {1} {2} {3}", currentField.ColumnName, reqTypeField.IsRequired, currentField.Type, currentField.ColumnName));
                        }
                    }
                }

                //konec pizdetsa

                for (int i = 0; i < rCount; ++i)
                {
                    
                    ++progressBar1.Value;
                    
                    Req rq = (Req)rf.AddItem(DBNull.Value);
                    rq.Name = "Requirement_" + i.ToString() + rname;  //for alm 11 or alm 11.50
                    //rq.Type = arr2[i % arr2.Length];
                                                        
                        rq["RQ_TYPE_ID"] = arr2[i % arr2.Length];
                    
                    // rq.ParentId = rqf.ID;
                    rq.ParentId = req_l2[i % req_l2.Count];
                    rq["RQ_REQ_AUTHOR"] = user_;

                    //customization required fields start

                    foreach (CustomizationField field in reqFieldsListreqs)
                    {
                        switch (field.Type)
                        {
                            case 0:
                            case 1:
                            case 2: //Int,int,float
                                if (!field.ColumnName.Equals("RQ_TYPE_ID")) rq[field.ColumnName] = rand2.Next(12);
                                break;
                                
                            case 3: //string 

                                if (field.RootId != null)
                                {
                                    
                                    CustomizationListNode rootNode = customization.Fields.Field["Req", field.ColumnName].List.RootNode;

                                    List<string> nodesNamesList = GetAllNodesNames(rootNode, true);
                                    string[] nodes = nodesNamesList.ToArray();
                                    rq[field.ColumnName] = nodes[rand2.Next(0, nodes.Length - 1)];                               
                                }
                                else
                                {
                                    rq[field.ColumnName] = "Requirement_" + i.ToString() + rname;
                                }

                                break;

                            case 4: //memo
                                rq[field.ColumnName] = "Requirement_" + i.ToString() + rname;
                                break;

                            case 5: //Date
                                rq[field.ColumnName] = DateTime.Now.ToString("MM-dd-yyyy");
                                break;

                            case 7: //lookup list in requirements is different from others for some reason - its type is 7 - wtf?
                              if (field.RootId != null)
                                {
                                    CustomizationListNode rootNode = customization.Fields.Field["Req", field.ColumnName].List.RootNode;

                                    List<string> nodesNamesList = GetAllNodesNames(rootNode, true);
                                    string[] nodes = nodesNamesList.ToArray();
                                    rq[field.ColumnName] = nodes[rand2.Next(0, nodes.Length - 1)];


                                }
                                else
                                {
                                    rq[field.ColumnName] = "Requirement_" + i.ToString() + rname;
                                }
                              break;
                            case 8: // List of users
                                rq[field.ColumnName] = connection_.UserName;
                                break;
                        }
                    }


                    //end


                    rq.Post();

                    //req_l1.Add(rq.ID);
                    if (checkBox1.Checked) rq.AddCoverage(test_l1[i % test_l1.Count], 0);

                    listBox1.Items.Add(String.Format("{0} {1}", rq.ID, rq.Name));
                }
                reqFieldsListreqs.Clear();
                progressBar1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                reqFieldsListreqs.Clear();
                progressBar1.Visible = false;
            }
        }

        private void Form1_Closed(object sender, EventArgs e)
        {
            try
            {
                
                XmlTextWriter writer = new XmlTextWriter(@"C:\ALM_DATA_SPAMMER_config.xml", System.Text.Encoding.UTF8);
                writer.WriteStartDocument(true);
                writer.Formatting = Formatting.Indented;
                writer.Indentation = 2;
                writer.WriteStartElement("Config");
                writer.WriteStartElement("URL");
                writer.WriteString(textBox2.Text);
                writer.WriteEndElement();
                writer.WriteStartElement("DOMAIN");
                writer.WriteString(textBox3.Text);
                writer.WriteEndElement();
                writer.WriteStartElement("Project");
                writer.WriteString(textBox4.Text);
                writer.WriteEndElement();
                writer.WriteStartElement("user");
                writer.WriteString(textBox5.Text);
                writer.WriteEndElement();
                writer.WriteStartElement("password");
                writer.WriteString(textBox6.Text);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
                


                if (connection_.Connected) connection_.Disconnect();
                if (connection_.LoggedIn)
                {
                    connection_.Logout();
                    connection_.ReleaseConnection();
                }
            }
            catch
            {
                MessageBox.Show("Disconnected ;)");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            tsCount = Convert.ToInt32(textBox9.Text);
            Random rand3 = new Random();
            tsfoldername = rand3.Next();

            hCount = Convert.ToInt32(textBox23.Text);

            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = tsCount + 1;

            TestSetTreeManager tsm = (TestSetTreeManager)connection_.TestSetTreeManager;
            try
            {
                for (int b = 0; b < hCount; ++b)
                {
                    if (b == 0)
                    {
                        TestSetFolder rn = (TestSetFolder)tsm.get_NodeById(0);
                        //TestSetFolder rn = (TestSetFolder)tsm.Root;
                        //TestSetFolder tsf = new TestSetFolder();
                        TestSetFolder tsfolder = (TestSetFolder)rn.AddNode("TSFolder" + tsfoldername.ToString());

                        if (checkBox3.Checked) tsfolder.TargetCycle = cycle_l1[tsfoldername % cycle_l1.Count];

                        tsfolder.Post();

                        tsetParent = tsfolder.NodeID;
                        tset_l1.Add(tsfolder.NodeID);

                        listBox1.Items.Add(String.Format("{0} {1}", tsfolder.NodeID, tsfolder.Name));
                        // TestSetFactory ts = (TestSetFactory)tsfolder.TestSetFactory;
                    }
                    else
                    {
                        TestSetFolder rn = (TestSetFolder)tsm.get_NodeById(tsetParent);
                        TestSetFolder tsfolder = (TestSetFolder)rn.AddNode("TSFolder" + tsfoldername.ToString());

                        if (checkBox3.Checked) tsfolder.TargetCycle = cycle_l1[tsfoldername % cycle_l1.Count];

                        tsfolder.Post();

                        tsetParent = tsfolder.NodeID;
                        tset_l1.Add(tsfolder.NodeID);

                        listBox1.Items.Add(String.Format("{0} {1}", tsfolder.NodeID, tsfolder.Name));
                    }
                }
                // TestSetFactory ts = (TestSetFactory)connection_.TestSetFactory;
                for (int i = 0; i < tsCount; ++i)
                {
                    ++progressBar1.Value;
                    TestSetFolder rn = (TestSetFolder)tsm.get_NodeById(tset_l1[i % tset_l1.Count]);
                    TestSetFactory ts = (TestSetFactory)rn.TestSetFactory;
                    TestSet test_set = (TestSet)ts.AddItem(DBNull.Value);

                    // test_set.Name = "TEST_SET" + i.ToString() + tsfoldername;
                    test_set["CY_CYCLE"] = "TEST_SET" + i.ToString() + tsfoldername;
                    // test_set["CY_FOLDER_ID"] = tsfolder.NodeID;
                    // test_set["CY_FOLDER_ID"] = tset_l1[i % tset_l1.Count];
                    // test_set.TestSetFolder = tset_l1[i % tset_l1.Count];


                    if (project == "TD_easyweb_E2E")
                    {
                        test_set["CY_USER_04"] = "CIAM";
                        test_set["CY_USER_02"] = "Pilot";
                        test_set["CY_USER_01"] = "Y";
                        test_set["CY_USER_06"] = "Y";
                        test_set["CY_USER_03"] = "1";
                    }

                    test_set.Post();

                    if (checkBox2.Checked)
                    {
                        for (int a = 0; a < test_l1.Count; ++a)
                        {
                            TSTestFactory testInstanceF = test_set.TSTestFactory;
                            TSTest tstInstance = (TSTest)testInstanceF.AddItem(DBNull.Value);
                            tstInstance["TC_TEST_ID"] = test_l1[a % test_l1.Count];
                            tstInstance.Status = "No Run";
                            tstInstance["TC_CYCLE_ID"] = test_set.ID;
                            tstInstance.Post();
                            tstInstance["TC_CYCLE_ID"] = test_set.ID;
                        }
                    }


                    listBox1.Items.Add(String.Format("{0} {1}", test_set.ID, test_set.Name));
                }
                progressBar1.Visible = false;
                tset_l1.Clear();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                cCount = Convert.ToInt32(textBox10.Text);

                progressBar1.Enabled = true;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = cCount + 1;

                Random rand4 = new Random();
                relname = rand4.Next();
                ReleaseFolderFactory relf = (ReleaseFolderFactory)connection_.ReleaseFolderFactory;
                ReleaseFolder root_release_folder = (ReleaseFolder)relf.Root;
                ReleaseFolderFactory relff = (ReleaseFolderFactory)root_release_folder.ReleaseFolderFactory;
                ReleaseFolder release_folder = (ReleaseFolder)relff.AddItem(DBNull.Value);
                release_folder.Name = "Rel_FOLDER_" + relname;
                release_folder.Post();

                ReleaseFactory rel = (ReleaseFactory)release_folder.ReleaseFactory;
                Release my_release = (Release)rel.AddItem(DBNull.Value);

                my_release.Name = "RELEASE_" + relname;

                my_release.StartDate = dateTimePicker1.Value;
                my_release.EndDate = dateTimePicker2.Value;
                my_release.Post();

                listBox1.Items.Add(String.Format("{0} {1}", my_release.ID, my_release.Name));

                CycleFactory cf = (CycleFactory)my_release.CycleFactory;

                for (int i = 0; i < cCount; ++i)
                {
                    ++progressBar1.Value;
                    Cycle my_cycle = (Cycle)cf.AddItem(DBNull.Value);
                    my_cycle.Name = "CYCLE_" + i.ToString() + relname;
                    my_cycle.StartDate = dateTimePicker1.Value;
                    my_cycle.EndDate = dateTimePicker1.Value.AddDays(i);

                    if (my_cycle.EndDate > my_cycle.StartDate) my_cycle.EndDate = dateTimePicker2.Value;
                    
                    my_cycle.Post();

                    cycle_l1.Add(my_cycle.ID);

                    listBox1.Items.Add(String.Format("{0} {1}", my_cycle.ID, my_cycle.Name));

                    checkBox3.Enabled = true;
                }
                progressBar1.Visible = false;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }         
            
        }

        

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                dCount = Convert.ToInt32(textBox11.Text);

                progressBar1.Enabled = true;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = dCount + 1;

                Random rand5 = new Random();
               

                for (int i = 0; i < dCount; ++i)
                {
                    ++progressBar1.Value;
                    dname = rand5.Next();
                    sconnection.CreateDomain("Domain" + dname, "", "", 500);

                    domain_l1.Add("Domain" + dname);

                    listBox1.Items.Add(String.Format("{0} {1}", "  ", "Domain" + dname));
                }
                progressBar1.Visible = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Please, pick correct domain numbers" + MessageBox.Show(ex.Message));
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {

            server_ = textBox2.Text;
            user_ = textBox5.Text;
            password_ = textBox6.Text;
            domain_ = textBox3.Text;
            project = textBox4.Text;

            try
            {
               
                sconnection.Login(server_, user_, password_);
                MessageBox.Show("Connected");
                button9.Enabled = true;
                button11.Enabled = true;
                button12.Enabled = true;
                button10.Enabled = false;

            }
            catch (Exception ex)
            {
                label48.Visible = true;
                MessageBox.Show(ex.Message);
                label48.Visible = false;
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
                        
            try
            {

                sconnection.Logout();
                MessageBox.Show("Disconnected");
                button9.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
                button13.Enabled = false;
                button10.Enabled = true;
            }
            catch
            {
                MessageBox.Show("Damn it, something is not right, man");
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            
            DB_TYPE = Convert.ToInt32(comboBox1.Text); 
            DB_SERVER_NAME = textBox13.Text;
            db_user_ = textBox18.Text;
            db_password_ = textBox17.Text;
            domain_ = textBox15.Text;
            TABLESPACE_ = textBox14.Text;
            TEMP_TABLESPACE = textBox16.Text;

            try
            {
                pCount = Convert.ToInt32(textBox12.Text);

                progressBar1.Enabled = true;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = pCount + 1;

                Random rand6 = new Random();

                

                for (int i = 0; i < pCount; ++i)
                {
                    ++progressBar1.Value;
                    pname = rand6.Next();
                    //sconnection.CreateProject(domain_l1[1],"Project" + pname, DB_TYPE, DB_SERVER_NAME,user_, password_, TABLESPACE_, 
                    sconnection.CreateProject(domain_, "Project" + pname, DB_TYPE, DB_SERVER_NAME, db_user_, db_password_, TABLESPACE_, TEMP_TABLESPACE, 0, 0, 1);
                    listBox1.Items.Add(String.Format("{0} {1}", "  ", "Project" + pname));
                    
                }
                progressBar1.Visible = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show("db_type is" + DB_TYPE);
                MessageBox.Show(ex.Message);
            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DB_TYPE = Convert.ToInt32(comboBox1.Text);
            if (DB_TYPE == 2)
            {
                textBox14.Enabled= false;
                textBox16.Enabled= false;

            }else{
                textBox14.Enabled= true;
                textBox16.Enabled= true;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:alexander.kostuchenko@hp.com?subject=hello&body=What's up";
            proc.Start();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                
                liCount = Convert.ToInt32(textBox19.Text);

                progressBar1.Enabled = true;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = liCount + 1;

                Random rand6 = new Random();
                list_name = textBox20.Text;
                

                Customization cust = (Customization)connection_.Customization;
                CustomizationLists clists = (CustomizationLists)cust.Lists;
                
                

                for (int i = 0; i < liCount; ++i)
                {
                    ++progressBar1.Value;
                    item_name = "list_item_" + Convert.ToString(rand6.Next());
                    CustomizationList my_list = (CustomizationList)clists.get_List(list_name);
                    CustomizationListNode lnode = (CustomizationListNode)my_list.RootNode;
                    CustomizationListNode new_item = (CustomizationListNode)lnode.AddChild(item_name + i);
                   // cust.Commit();

                    listBox1.Items.Add(String.Format("{0} {1}", "  ", item_name));


                }
                listBox1.Items.Add(String.Format("{0} {1}", "  ", "Commiting items"));
                cust.Commit();
                listBox1.Items.Add(String.Format("{0} {1}", "  ", "DONE"));
            }

           
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            progressBar1.Visible = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
                                    
            try
            {

                usCount = Convert.ToInt32(textBox22.Text);
                user_group_list_name = textBox21.Text;
                Random rand7 = new Random();

                progressBar1.Enabled = true;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = usCount + 1;


                Customization cust = (Customization)connection_.Customization;
                CustomizationUsers cusers = (CustomizationUsers)cust.Users;
                CustomizationUsersGroups cusersgroups = (CustomizationUsersGroups)cust.UsersGroups;
                CustomizationUsersGroup usr_grp = (CustomizationUsersGroup)cusersgroups.get_Group(user_group_list_name);

                cust.Load();

                for (int i = 0; i < usCount; ++i)
                {
                    ++progressBar1.Value;
                    user_name = "user_" + Convert.ToString(rand7.Next());
                    cusers.AddSiteUser(user_name, user_name,"cat6654@gmail.com",Convert.ToString(rand7.Next()) ,"7686654",usr_grp);
                    cust.Commit();
                    cusers.AddUser(user_name);

                }
                listBox1.Items.Add(String.Format("{0} {1}", "  ", "Commiting items"));
                cust.Commit();
                listBox1.Items.Add(String.Format("{0} {1}", "  ", "DONE"));

                progressBar1.Visible = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
           
            try
            {
                server_ = textBox2.Text;
                user_ = textBox5.Text;
                password_ = textBox6.Text;
                domain_ = textBox3.Text;
                project = textBox4.Text;
                
                
                REST.auth(server_ + "/authentication-point/alm-authenticate", "<?xml version='1.0' encoding='utf-8'?><alm-authentication><user>" + user_ + "</user><password>" + password_ + "</password></alm-authentication>");
                REST.post(server_ + "/rest/site-session", null);

               // listBox1.Items.Add(String.Format("{0} {1}", "  ", "your cookie is " + REST.myheader));
                MessageBox.Show("Connected to REST");
            }
            catch (Exception ex)
            {
                label48.Visible = true;
                MessageBox.Show(ex.Message);
                label48.Visible = false;
            }
            button15.Enabled = true;
            button16.Enabled = true;
            button17.Enabled = true;
            button18.Enabled = true;
            button14.Enabled = false;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            button14.Enabled = true;
            button15.Enabled = false;
            button16.Enabled = false;
            button17.Enabled = false;
            button18.Enabled = false;

            try
            {
                server_ = textBox2.Text;
                user_ = textBox5.Text;
                password_ = textBox6.Text;
                domain_ = textBox3.Text;
                project = textBox4.Text;

                REST.get(server_ + "/authentication-point/logout");
                listBox1.Items.Add(String.Format("{0} {1}", "REST connection is closed ;)", " "));
                MessageBox.Show("REST connection is closed ;)");
                              
            }
            catch (Exception ex)
            {
                label48.Visible = true;
                MessageBox.Show(ex.Message);
                label48.Visible = false;
            }


        }

        private void button16_Click(object sender, EventArgs e)
        {

            server_ = textBox2.Text;
            user_ = textBox5.Text;
            password_ = textBox6.Text;
            domain_ = textBox3.Text;
            project = textBox4.Text;

            try
            {
               

                REST.get(server_ + "/rest/domains/" + domain_ +"/projects/" + project + "/defects/?page-size=1000");
                XPARSER.readXml(REST.backstr);

               for (int i = 0; i < XPARSER.ids.Count; i++)
                {
                    listBox1.Items.Add(String.Format("{0} {1}", XPARSER.ids[i], XPARSER.names[i]));
                }
                XPARSER.ids.Clear();
                XPARSER.names.Clear();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        private void button17_Click(object sender, EventArgs e)
        {
            server_ = textBox2.Text;
            user_ = textBox5.Text;
            password_ = textBox6.Text;
            domain_ = textBox3.Text;
            project = textBox4.Text;

            rest_bCount = Convert.ToInt32(textBox24.Text);

            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = rest_bCount + 1;


            Random rand = new Random();
            rand.Next();

          

            string[] arr = new string[] { "New", "Open", "Fixed", "Closed", "Reopen", "Rejected" };
            string[] arr2 = new string[] { "1-Low", "2-Medium", "3-High", "4-Very High", "5-Urgent"};
            DateTime date = DateTime.Now;
            string format = "yyyy-MM-dd";
            try 
            {
                
                for (int i = 0; i < rest_bCount; ++i)
                   {

                       ++progressBar1.Value;
                    REST.post(server_ + "/rest/domains/" + domain_ + "/projects/" + project + "/defects/",
                        @"<?xml version='1.0' encoding='utf-8'?>
                        <Entity Type='defect'>
                        <Fields>
                        <Field Name='description'>
                        <Value>desc</Value>
                        </Field>
                        <Field Name='name'>
                        <Value>defect created from rest" + i + Convert.ToString(rand.Next()) + "</Value></Field><Field Name='status'><Value>" + arr[i % arr.Length] + "</Value></Field><Field Name='severity'><Value>" + arr2[i % arr2.Length] + "</Value></Field><Field Name='detected-by'><Value>" + user_ + "</Value></Field><Field Name='creation-time'><Value>" + date.ToString(format) + "</Value></Field></Fields></Entity>"
                        );
                    XPARSER.readXml(REST.backstr);

                    if (checkBox5.Checked && i > 0)
                    {
                        REST.post(server_ + "/rest/domains/" + domain_ + "/projects/" + project + "/defect-links/",
                            @"<?xml version='1.0' encoding='utf-8'?>
                            <Entity Type='defect-link'>
                            <Fields>
                            <Field Name='first-endpoint-id'>
                            <Value>" + XPARSER.ids[i] + "</Value></Field><Field Name='second-endpoint-type'><Value>defect</Value></Field><Field Name='second-endpoint-id'><Value>" + XPARSER.ids[i-1] + "</Value></Field></Fields><RelatedEntities /></Entity>"

                            );
                    }

                    listBox1.Items.Add(String.Format("{0} {1}", XPARSER.ids[i], XPARSER.names[i]));
                    }
                
                
                XPARSER.ids.Clear();
                XPARSER.names.Clear();

                progressBar1.Visible = false;

            }

                catch (Exception ex)
                {
                    progressBar1.Visible = false;
                    MessageBox.Show(ex.Message);
                }

            }

        private void button18_Click(object sender, EventArgs e)
        {
            server_ = textBox2.Text;
            user_ = textBox5.Text;
            password_ = textBox6.Text;
            domain_ = textBox3.Text;
            project = textBox4.Text;

            int rest_rCount = Convert.ToInt32(textBox25.Text);

            Random rand = new Random();
            rand.Next();

            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = rest_rCount + 1;

            REST.post(server_ + "/rest/domains/" + domain_ + "/projects/" + project + "/requirements/",
                @"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
                    <Entity Type='requirement'>
                    <Fields>
                    <Field Name='description'>
                    <Value>desc</Value>
                    </Field>
                    <Field Name='parent-id'>
                    <Value>0</Value>
                    </Field>
                    <Field Name='type-id'><Value>1</Value></Field>
                    <Field Name='req-reviewed'><Value>Reviewed</Value></Field>
                    <Field Name='owner'><Value>" + user_ +"</Value></Field><Field Name='name'><Value>Req_rest_folder" + Convert.ToString(rand.Next()) +"</Value></Field></Fields><RelatedEntities/></Entity>"

                );
            XPARSER.readXml(REST.backstr);

            string parent_folder_id = XPARSER.ids[0];
            listBox1.Items.Add(String.Format("{0} {1}", XPARSER.ids[0], XPARSER.names[0]));

            XPARSER.ids.Clear();
            XPARSER.names.Clear();

            try
            {
                for (int i = 0; i < rest_rCount; ++i)
                   {
                       ++progressBar1.Value;

                       REST.post(server_ + "/rest/domains/" + domain_ + "/projects/" + project + "/requirements/",
                   @"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
                    <Entity Type='requirement'>
                    <Fields>
                    <Field Name='description'>
                    <Value>desc</Value>
                    </Field>
                    <Field Name='type-id'><Value>3</Value></Field>
                    <Field Name='req-reviewed'><Value>Reviewed</Value></Field>
                    <Field Name='parent-id'>
                    <Value>" + parent_folder_id + "</Value></Field><Field Name='owner'><Value>" + user_ + "</Value></Field><Field Name='name'><Value>Req_rest_func" + Convert.ToString(rand.Next()) + "</Value></Field></Fields><RelatedEntities/></Entity>"

                   );
                       XPARSER.readXml(REST.backstr);
                       listBox1.Items.Add(String.Format("{0} {1}", XPARSER.ids[i], XPARSER.names[i]));

                   }

                XPARSER.ids.Clear();
                XPARSER.names.Clear();

                progressBar1.Visible = false;

            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show(ex.Message);
            }


        }

        private void button19_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Items.Add(String.Format("{0} {1}", "ID ", "Name"));
        }

        
    

            
              
    


        private void button20_Click(object sender, EventArgs e)
        {

            server_ = textBox2.Text;
            user_ = textBox5.Text;
            password_ = textBox6.Text;
            domain_ = textBox3.Text;
            project = textBox4.Text;

            
            try
            {

                label54.Visible = true;
                label55.Visible = true;
                WebGateCredential.ClearWebServerCredential();
                QCCH_API.initial();
                QCCH_API.download(server_);
                QCCH_API.getServiceManager();
                QCCH_API.login(user_, password_, domain_, project);
               
               

                label54.Visible = false;
                label55.Visible = false;
                button21.Enabled = true;
                

                if (QCCH_API.connectionManagementService.IsLoggedIn == true)
                {
                    button21.Enabled = true;
                    button20.Enabled = false;
                    button22.Enabled = true;

                    MessageBox.Show("Connected via ConnectivityHelper ^__^");
                }
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show(ex.Message);
            }

        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                QCCH_API.logout();
                QCCH_API.end();

                button21.Enabled = false;
                button20.Enabled = true;

                MessageBox.Show("Disconnected ^__^");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                QCCH_API.getFactory();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        //change sanitization
        private void button23_Click(object sender, EventArgs e)
        {
            //Customization cust = (Customization)connection_.Customization;
            //CustomizationFields cfields = (CustomizationFields)cust.Fields;

            //cust.Load();

            //try
            //{
            //    CustomizationField cfield = (CustomizationField)cfields.get_Field(textBox30.Text, textBox27.Text);
            //    cfield.OutputSanitizationType = comboBox2.Text;
            //    cust.Commit();

            //    listBox1.Items.Add(String.Format("{0} {1}", "Changed Sanitization to ", cfield.OutputSanitizationType));
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }

        //create udf
        private void button24_Click(object sender, EventArgs e)
        {
            Random rand7 = new Random();
            int UDFCount = Convert.ToInt32(textBox29.Text);
            string listType = comboBox3.Text;
            

            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = UDFCount + 1;

            Customization cust = (Customization)connection_.Customization;
            CustomizationFields cfields = (CustomizationFields)cust.Fields;

            cust.Load();

          //  dynamic listName = cust.Lists.List("Status");

         //   CustomizationList cList = (CustomizationList)cust.Lists.List("Status");
            
            try
            {
                for (int i = 0; i < UDFCount; ++i)
                {
                    if (i >= 99)
                    {
                        MessageBox.Show("Maximum UDF count is 99");
                        break;
                    }

                    ++progressBar1.Value;
                    CustomizationField cfield = (CustomizationField)cfields.AddActiveField(textBox28.Text);

                    //cfield.UserLabel = textBox28.Text + "_" + "UDF" + i + rand7.Next();
                    cfield.UserLabel = cfield.ColumnName;
                    if (listType == "string")
                    {
                        cfield.Type = 3;
                      //  cfield.OutputSanitizationType = "html";
                        cfield.FieldSize = 255;
                    }
                    if (listType == "number") cfield.Type = 0;
                    if (listType == "float(pre12.01)") cfield.Type = 2;
                    
                    if (listType == "memo")
                    {
                        cfield.Type = 4;
                   //     cfield.OutputSanitizationType = "text";
                    }
                    if (listType == "userList") cfield.Type = 8;
                    if (listType == "list")
                    {
                        cfield.Type = 3;
                        cfield.List = cust.Lists.List("Status");
                   //     cfield.OutputSanitizationType = "none";
                        cfield.FieldSize = 255;
                    }
                    if (listType == "date") cfield.Type = 5;
                    //cfield.Type = 21;// 3 - string, 2 - int, 4 - memo, 5 - date, 8 - userList , 21 - list, 22 - multivaluelist
       
                    listBox1.Items.Add(String.Format("{0} {1}", "  ", cfield.ColumnName));
                }
                cust.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                progressBar1.Visible = false;
            }
            progressBar1.Visible = false;

        }

        //create user group
        private void button25_Click(object sender, EventArgs e)
        {

            try
            {

                ugCount = Convert.ToInt32(textBox32.Text);
                user_group_list_name = textBox31.Text;
                Random rand8 = new Random();

                progressBar1.Enabled = true;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = ugCount + 1;

                Customization cust = (Customization)connection_.Customization;
                CustomizationUsers cusers = (CustomizationUsers)cust.Users;
                CustomizationUsersGroups cusersgroups = (CustomizationUsersGroups)cust.UsersGroups;
  
                cust.Load();

                for (int i = 0; i < ugCount; ++i)
                {
                    ++progressBar1.Value;
                    userGroup_name = "user_group" +  i + Convert.ToString(rand8.Next());
                    cusersgroups.AddGroupAsGroup(userGroup_name, user_group_list_name);
                    CustomizationUsersGroup usr_grp = (CustomizationUsersGroup)cusersgroups.get_Group(userGroup_name);
                    cust.Commit();
                    listBox1.Items.Add(String.Format("{0} {1}", usr_grp.ID, usr_grp.Name));
                    
                }
                
                listBox1.Items.Add(String.Format("{0} {1}", "  ", "DONE"));

                progressBar1.Visible = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //attach images
        private void button26_Click(object sender, EventArgs e)
        {

            //list of images file names
            List<string> images = new List<string>(); 

            //path to images
            if (!System.IO.Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "Images")))
            {
                throw new Exception("You should have Images folder in the same place where your DataSpammer exe is !!!");
            }
            string pathToImages = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            DirectoryInfo d = new DirectoryInfo(pathToImages);
            //get all files in a folder
            foreach (var file in d.GetFiles("*.*"))
            {
                images.Add(file.Name);
            }

            Customization cust = (Customization)connection_.Customization;
            CustomizationFields cfields = (CustomizationFields)cust.Fields;

            cust.Load();

            try
            {
                CustomizationField cfield = (CustomizationField)cfields.get_Field("BUG", "BG_DESCRIPTION");
                if (cfield.OutputSanitizationType != "none")
                {
                    cfield.OutputSanitizationType = "none";
                    cust.Commit();
                    listBox1.Items.Add(String.Format("{0} {1}", "Changed Sanitization of bug description to ", cfield.OutputSanitizationType));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //get existing defects' IDs
            BugFactory bf = (BugFactory)connection_.BugFactory;
            //create empty filter
            TDFilter filter = (TDFilter)bf.Filter;
            //get defects using empty filter. This means we will get ALL the defects.
            List bugList = filter.NewList();

            progressBar1.Enabled = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = bugList.Count + 1;

            Regex r = new Regex(@"^.*?(?=</body></html>)");
            
            foreach (Bug b in bugList)
            {
                AttachmentFactory attachFact = (AttachmentFactory)b.Attachments;
                Attachment a = attachFact.AddItem(DBNull.Value);
                a.FileName = Path.Combine(pathToImages, images[b.ID % images.Count]);
                a.Type = 1; //this means file.
                a.Post();
                //get current description
                string currentDescription = b["BG_DESCRIPTION"];
                //remove \r\n 
                currentDescription = currentDescription.Replace("\r\n",string.Empty);
                Match parsedBody = r.Match(currentDescription);
                b["BG_DESCRIPTION"] = parsedBody + "<img src=\"file://[IMAGE_BASE_PATH_PLACEHOLDER]" + images[b.ID % images.Count] + "\" width=\"96\" height=\"96\" border=\"0\" alt=\"graphic\"/></html></body>";
                b.Post();
                progressBar1.Value++;
            }
            progressBar1.Visible = false;
            listBox1.Items.Add("Added attachments to " + bugList.Count + " Defects");

        }     
        }
    

            
              
    }

