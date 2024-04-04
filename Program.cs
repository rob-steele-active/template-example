// See https://aka.ms/new-console-template for more information
using ClosedXML.Report;

namespace Main
{
    class Program
    {
        static void Main(string[] args)
        {
            const string outputFile = @"./TestOutput.xlsx";
            var template = new XLTemplate(@"./TestTemplate.xlsx");

            List<Item> data = new List<Item>() { new Item("Alice", 1), new Item("Jeffrey", 9000) };
            var aCampaign = new Campaign("TEST", "EXCEL", "2024", data);

            template.AddVariable(aCampaign);
            template.Generate();
            template.SaveAs(outputFile);

            Console.WriteLine("Done!");
        }
    }
    class Campaign
    {
        public string ClientCode;
        public string ProductCode;
        public string EstimateCode;

        public List<Item> Data;

        public Campaign(string clientCode, string productCode, string estimateCode, List<Item> data)
        {
            this.ClientCode = clientCode;
            this.ProductCode = productCode;
            this.EstimateCode = estimateCode;
            this.Data = data;
        }
    }

    class Item
    {
        public string Name;
        public int Age;

        public Item(string name, int age)
        {
            this.Name = name;
            this.Age = age;
        }
    }
}
