// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using OpenXmlPowerTools.Documents;
using System;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                PrintUsage();
                Environment.Exit(0);
            }

            FileInfo templateDoc = new FileInfo(args[0]);
            if (!templateDoc.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[0]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo dataFile = new FileInfo(args[1]);
            if (!dataFile.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[1]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo assembledDoc = new FileInfo(args[2]);
            if (assembledDoc.Exists)
            {
                Console.WriteLine("Error, {0} exists.", args[2]);
                PrintUsage();
                Environment.Exit(0);
            }

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            bool templateError;
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
            }

            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }

        private static void Example01()
        {

        }

        private static void Example02()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            FileInfo templateDoc = new FileInfo("../../TemplateDocument.docx");
            FileInfo dataFile = new FileInfo("../../Data.xml");

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            bool templateError;
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See AssembledDoc.docx to determine the errors in the template.");
            }

            FileInfo assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, "AssembledDoc.docx"));
            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }

        private static void Example03()
        {

        }

        private static void Example04()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            FileInfo templateDoc = new FileInfo("../../TemplateDocument.docx");
            FileInfo dataFile = new FileInfo(Path.Combine(tempDi.FullName, "Data.xml"));

            // The following method generates a large data file with random data.
            // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
            XElement data = GenerateDataFromDataSource(dataFile);

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            int count = 1;
            foreach (var customer in data.Elements("Customer"))
            {
                FileInfo assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, string.Format("Letter-{0:0000}.docx", count++)));
                Console.WriteLine(assembledDoc.Name);
                bool templateError;
                WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out templateError);
                if (templateError)
                {
                    Console.WriteLine("Errors in template.");
                    Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
                }
                wmlAssembledDoc.SaveAs(assembledDoc.FullName);
            }
        }

        private static string[] s_productNames = new[] {
            "Unicycle",
            "Bicycle",
            "Tricycle",
            "Skateboard",
            "Roller Blades",
            "Hang Glider",
        };

        private static XElement GenerateDataFromDataSource(FileInfo dataFi)
        {
            int numberOfDocumentsToGenerate = 500;
            var customers = new XElement("Customers");
            Random r = new Random();
            for (int i = 0; i < numberOfDocumentsToGenerate; ++i)
            {
                var customer = new XElement("Customer",
                    new XElement("CustomerID", i + 1),
                    new XElement("Name", "Eric White"),
                    new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                    new XElement("Orders"));
                var orders = customer.Element("Orders");
                int numberOfOrders = r.Next(10) + 1;
                for (int j = 0; j < numberOfOrders; j++)
                {
                    var order = new XElement("Order",
                        new XAttribute("Number", j + 1),
                        new XElement("ProductDescription", s_productNames[r.Next(s_productNames.Length)]),
                        new XElement("Quantity", r.Next(10)),
                        new XElement("OrderDate", "September 26, 2015"));
                    orders.Add(order);
                }
                customers.Add(customer);
            }
            customers.Save(dataFi.FullName);
            return customers;
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage: DocumentAssembler TemplateDocument.docx Data.xml AssembledDoc.docx");
        }
    }
}
