/*
Created by: John Walker
Last modified: 1/31/2018
Description: Class to save XML files
*/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;

namespace AdminRDP
{
    public class SaveXML
    {
        public static void SaveData(object obj, string filename)
        {
            XmlSerializer sr = new XmlSerializer(obj.GetType());
            TextWriter writer = new StreamWriter(filename);
            sr.Serialize(writer, obj);
            writer.Close();
        }
    }
}
