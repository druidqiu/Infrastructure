using System.Collections;
using System.Xml;

namespace SYDQ.Infrastructure.Configuration
{
    public static class AppConfigReader
    {
        private static Hashtable _configItems;

        public static void Load(string appPath)
        {
            _configItems = new Hashtable();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(appPath);
            XmlNodeList itemNodes = xmlDoc.SelectNodes("/configuration/item");
            if (itemNodes == null) return;
            foreach (XmlNode node in itemNodes)
            {
                if (node.Attributes != null)
                    _configItems.Add(node.Attributes["key"].Value, node.Attributes["value"].Value);
            }
        }

        public static string Config(string key)
        {
            if (_configItems.ContainsKey(key))
                return (string) _configItems[key];
            return string.Empty;
        }
    }
}
