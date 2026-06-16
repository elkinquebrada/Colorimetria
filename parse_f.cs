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

        using (StreamWriter writer = new StreamWriter(@"C:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\parsed_formulas.txt"))
        {
            foreach (XmlNode cNode in doc.SelectNodes("//x:c", nsmgr))
            {
                XmlNode fNode = cNode.SelectSingleNode("x:f", nsmgr);
                if (fNode != null)
                {
                    string fText = fNode.InnerText;
                    string rAttr = "";
                    if (cNode.Attributes["r"] != null) {
                        rAttr = cNode.Attributes["r"].Value;
                    }
                    writer.WriteLine(rAttr + ": " + fText);
                }
            }
        }
    }
}
