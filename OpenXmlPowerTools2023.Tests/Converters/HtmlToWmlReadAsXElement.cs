// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

/*******************************************************************************************
 * HtmlToWmlConverter expects the HTML to be passed as an XElement, i.e. as XML.  While the HTML test files that
 * are included in Open-Xml-PowerTools are able to be read as XML, most HTML is not able to be read as XML.
 * The best solution is to use the HtmlAgilityPack, which can parse HTML and save as XML.  The HtmlAgilityPack
 * is licensed under the Ms-PL (same as Open-Xml-PowerTools) so it is convenient to include it in your solution,
 * and thereby you can convert HTML to XML that can be processed by the HtmlToWmlConverter.
 * 
 * A convenient way to get the DLL that has been checked out with HtmlToWmlConverter is to clone the repo at
 * https://github.com/EricWhiteDev/HtmlAgilityPack
 * 
 * That repo contains only the DLL that has been checked out with HtmlToWmlConverter.
 * 
 * Of course, you can also get the HtmlAgilityPack source and compile it to get the DLL.  You can find it at
 * http://codeplex.com/HtmlAgilityPack
 * 
 * We don't include the HtmlAgilityPack in Open-Xml-PowerTools, to simplify installation.  The XUnit tests in
 * this module do not require the HtmlAgilityPack to run.
*******************************************************************************************/

namespace OpenXmlPowerTools2023.Tests.Converters
{
    public class HtmlToWmlReadAsXElement
    {
        public static XElement ReadAsXElement(string sourceHtmlFi)
        {
            string htmlString = File.ReadAllText(sourceHtmlFi);
            XElement html = null;
            try
            {
                html = XElement.Parse(htmlString);
            }
            catch (XmlException e)
            {
                throw e;
            }
            html = (XElement)ConvertToNoNamespace(html);
            return html;
        }

        private static object ConvertToNoNamespace(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                return new XElement(element.Name.LocalName,
                    element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                    element.Nodes().Select(n => ConvertToNoNamespace(n)));
            }
            return node;
        }
    }
}