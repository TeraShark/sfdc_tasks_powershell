using System;
using System.IO;
using System.Reflection;
using System.Xml;
namespace XmlTools
{
    public class XmlTools
    {
        private string baseDirectory;
        public XmlTools(string currentDirectory)
        {
            baseDirectory = currentDirectory;
            baseDirectory = Path.Combine(baseDirectory, "XmlTools");
        }

        public enum CAMLType
        {
            Default,
            Today,
            SinceDate,
            MyItems,
            ByDealID
        }
        public string getCAML(CAMLType camlType)
        {
            return getCAML(camlType, "");
        }
        public string getCAML(CAMLType camlType, string arg)
        {
            var doc = new XmlDocument();

            switch (camlType)
            {
                case CAMLType.Default:
                    doc.Load(baseDirectory + "\\DefaultCAML.xml");
                    break;
                case CAMLType.Today:
                    doc.Load(baseDirectory + "\\TodayCAML.xml");
                    break;
                case CAMLType.SinceDate:
                    doc.Load(baseDirectory + "\\SinceDateCAML.xml");
                    // Set the Date for the query:
                    var dateNode = doc.SelectSingleNode("//Value[@Type='DateTime']");
                    if (dateNode != null)
                    {
                        dateNode.InnerText = arg;
                    }
                    break;
                case CAMLType.MyItems:
                    doc.Load(baseDirectory + "\\MyItemsCAML.xml");
                    break;
                case CAMLType.ByDealID:
                    doc.Load(baseDirectory + "\\ByDealIDCAML.xml");
                    var textNode = doc.SelectSingleNode("//Value[@Type='Text']");
                    if (textNode != null)
                    {
                        textNode.InnerText = arg;
                    }
                    break;
            }
            return doc.OuterXml;

        }
    }
}
