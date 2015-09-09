using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using TDAPIOLELib;

namespace OTA_Client
{
    public class CreateEntities
    {
        private readonly TDConnection _connection;

        public string DefectId { get; set; }
        public string DefectSummary { get; set; }

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
                            bg[field.Name] = field.Name + DateTime.Now.ToShortDateString();
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
                }
            }
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
