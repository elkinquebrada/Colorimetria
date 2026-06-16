using System;
using System.Xml;

var doc = new XmlDocument();
doc.Load(@"C:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\Color\LogicDocs\xlsx_extracted\xl\worksheets\sheet1.xml");
XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
nsmgr.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

foreach(XmlNode node in doc.SelectNodes("//x:c[x:f]", nsmgr))
{
    Console.WriteLine(node.Attributes["r"].Value + ": " + node.SelectSingleNode("x:f", nsmgr).InnerText);
}
