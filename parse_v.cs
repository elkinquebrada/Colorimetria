using System;
using System.IO;
using System.Xml;

class Program
{
    static void Main()
    {
        string xmlPath = @"C:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\Color\LogicDocs\xlsx_extracted\xl\worksheets\sheet1.xml";
        var doc = new XmlDocument();
        doc.Load(xmlPath);

        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
        nsmgr.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        using (StreamWriter writer = new StreamWriter(@"C:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\parsed_values.txt"))
        {
            // Specifically get cells we care about
            string[] targets = {"C16", "C17", "C20", "C24", "D24", "E24", "I20", "I24", "H24", "H20", "D7", "D8", "D9", "D53", "D54", "D55", "H53", "L53"};
            
            foreach (string t in targets)
            {
                XmlNode cNode = doc.SelectSingleNode("//x:c[@r='" + t + "']", nsmgr);
                if (cNode != null)
                {
                    XmlNode vNode = cNode.SelectSingleNode("x:v", nsmgr);
                    string vText = vNode != null ? vNode.InnerText : "NO_VAL";
                    
                    XmlNode fNode = cNode.SelectSingleNode("x:f", nsmgr);
                    string fText = fNode != null ? fNode.InnerText : "";
                    
                    writer.WriteLine(t + " = " + vText + " (Formula: " + fText + ")");
                }
                else
                {
                    writer.WriteLine(t + " = NOT FOUND");
                }
            }
        }
    }
}
