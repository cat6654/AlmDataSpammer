using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace WindowsFormsApplication5
{
  public  class xml_parser
    {
      public string url;
      public string domain;
      public string project;
      public string user;
      public string password;
      
      public string id;
      public string name;
      public string my_xml;
      public string type;
      public List<string> ids = new List<string>();
      public List<string> names = new List<string>();

      public void readXml(string my_xml)
      {

         
          XmlReader reader = (XmlReader)XmlReader.Create(new StringReader(my_xml));
          
                //  reader.ReadToFollowing("Entity");
              //    reader.MoveToFirstAttribute();
                //  type = reader.Value;

                  while (reader.Read())
                  {
                      reader.MoveToFirstAttribute();
                      reader.ReadToFollowing("Field");
                      reader.MoveToFirstAttribute();
                      if (reader.Value == "id" || reader.Value == "vc-checkout-date")
                      {
                          reader.ReadToFollowing("Value");
                          id = reader.ReadElementContentAsString();
                          ids.Add(id);
                      }
                      
                      if (reader.Value == "name")
                      {
                          reader.ReadToFollowing("Value");
                          name = reader.ReadElementContentAsString();
                          names.Add(name);
                      }
                      
                                        
                     }
        

      }


    }
}
