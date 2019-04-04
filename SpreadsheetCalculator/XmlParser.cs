using System;
using System.Collections.Generic;
using System.Xml;

namespace SpreadsheetCalculator
{
    public class XmlParser
    {
        private readonly XmlNamespaceManager _namespaceManager;
        private readonly XmlElement _root;

        private const string NamespaceName = "namespace";

        public XmlParser(string fileName)
        {
            var document = new XmlDocument();
            document.Load(fileName);

            _root = document.DocumentElement;
            if (_root == null)
                throw new Exception("Root is null.");

            var namespaceUri = document.DocumentElement != null ? document.DocumentElement.NamespaceURI : "";

            _namespaceManager = new XmlNamespaceManager(new NameTable());
            _namespaceManager.AddNamespace(NamespaceName, namespaceUri);
        }

        public IEnumerable<XmlNode> GetNodes(string tagName)
        {
            return GetNodes(_root, tagName);
        }

        public IEnumerable<XmlNode> GetNodes(XmlNode node, string tagName)
        {
            var rowNodes = node.SelectNodes($".//{NamespaceName}:{tagName}", _namespaceManager);
            if (rowNodes != null)
            {
                for (int i = 0; i < rowNodes.Count; i++)
                    yield return rowNodes.Item(i);
            }
        }

        public XmlNode GetSingleNode(XmlNode node, string tagName)
        {
            //            for (int i = 0; i < node.ChildNodes.Count; i++)
            //            {
            //                var childNode = node.ChildNodes.Item(i);
            //                if (childNode.Name == tagName)
            //                    return childNode;
            //            }

            //            return null;

            return node.SelectSingleNode($".//{NamespaceName}:{tagName}", _namespaceManager);
        }

        public Dictionary<string, string> GetAttributes(XmlNode node)
        {
            var attributes = new Dictionary<string, string>();

            if (node.Attributes != null)
            {
                for (int i = 0; i < node.Attributes.Count; i++)
                {
                    var attribute = node.Attributes[i];

                    attributes.Add(attribute.Name, attribute.InnerText);
                }
            }

            return attributes;
        }
    }
}