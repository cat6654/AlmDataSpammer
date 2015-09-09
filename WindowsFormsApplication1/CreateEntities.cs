using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using TDAPIOLELib;
using System.Windows.Forms;

namespace WindowsFormsApplication5
{
    public class CreateEntities
    {
        private readonly TDConnection _connection;

        public string DefectId { get; set; }
        public string DefectSummary { get; set; }
        public string fieldname;
        public string fieldtype;
        public int selectedrelID; 

        public CreateEntities(TDConnection connection)
        {
            _connection = connection;
        }

        public void PostDefect()
        {
            BugFactory factory = _connection.BugFactory;
            Customization customization = _connection.Customization;
            var reqFieldsList = new List<TDField>();
            var rnd = new Random();
            var bg = (Bug)factory.AddItem(DBNull.Value);

            foreach (TDField field in factory.Fields)
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
                        bg[field.Name] = rnd.Next(150);
                        break;

                    case 3: //string

                        if (customization.Fields.Field["Bug", field.Name].RootId != null)
                        {
                            CustomizationListNode rootNode = customization.Fields.Field["Bug", field.Name].List.RootNode;
                            List<string> nodesNamesList = GetAllNodesNames(rootNode, true);
                            
                            string[] nodes = nodesNamesList.ToArray();
                            bg[field.Name] = nodes[rnd.Next(0, nodes.Length - 1)];
                        }
                        else
                        {
                            bg[field.Name] = field.Name + rnd.Next(150);
                        }
                        break;

                    case 4: //memo
                        bg[field.Name] = field.Name + DateTime.Now.ToShortDateString();
                        break;

                    case 5: //Date
                        bg[field.Name] = DateTime.Now.ToString("MM-dd-yyyy");
                        break;

                    case 8: // List of users
                        bg[field.Name] = _connection.UserName;
                        break;

                   // case 15:
                  //  case 16:
                        //ReleaseFactory rel = (ReleaseFactory)_connection.ReleaseFactory;
                        //TDFilter relfilter= rel.Filter;
                        //relfilter["REL_NAME"] = "";

                        // List relList = rel.NewList(relfilter.Text);
                        // List<int> relnamelist = new List<int>();
                        // foreach (Release release in relList)
                        // {
                        //     relnamelist.Add(release.ID);
                        // }
                        // selectedrelID = relnamelist[rnd.Next(0, relnamelist.Count - 1)];
                        // bg[field.Name] = selectedrelID;
                        
                      //  break;
                    
                    case 17:
                    case 18: //release or cycle
                         //CustomizationListNode rootNode2 = customization.Fields.Field["Bug", field.Name].List.RootNode;
                         //   List<string> nodesNamesList2 = GetAllNodesNames(rootNode2, true);
                            
                         //   string[] nodes2 = nodesNamesList2.ToArray();
                         //   bg[field.Name] = nodes2[rnd.Next(0, nodes2.Length - 1)];

                        CycleFactory cycf = (CycleFactory)_connection.CycleFactory;
                        TDFilter cycfilter = cycf.Filter;
                        //cycfilter["RCYC_PARENT_ID"] = selectedrelID.ToString();
                        cycfilter["RCYC_NAME"] = "";

                         List cycList = cycf.NewList(cycfilter.Text);
                         List<int> cycIDList = new List<int>();
                         foreach (Cycle cycle in cycList)
                         {
                             cycIDList.Add(cycle.ID);
                         }

                         bg[field.Name] = cycIDList[rnd.Next(0, cycIDList.Count - 1)];
                         
                         break;
                }
            }
            bg.Summary = DefectSummary;
            bg.Post();
            DefectId = bg.ID.ToString();
            DefectSummary = bg.Summary;
        }

        private List<string> GetAllNodesNames(CustomizationListNode rootNode, bool isRoot, [Optional, DefaultParameterValue(null)]  List<string> nodesNames)
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
    }
}
