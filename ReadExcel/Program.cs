using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //Read Excel and return as List<entity>
            string path = @"e:\url.xlsx";
            ExcelHelper excel = ExcelHelper.Default;
            string str_Sql = "select * from [table$]";
            OleDbDataReader myRead = excel.getDataRead(str_Sql,path);

            //Set result into list
            List<entity> urlList = new List<entity>();

            while (myRead.Read())
            {
                entity set = new entity();
                set.url = myRead[0].ToString();
                set.component = myRead[1].ToString();
                set.className = myRead[2].ToString();
                set.elementId = myRead[3].ToString();
                urlList.Add(set);
            }
      
            
            foreach (entity ent in urlList)
            {
                Console.WriteLine("url = " + ent.url);
                Console.WriteLine("componet = " + ent.component);
                Console.WriteLine("className = " + ent.className);
                Console.WriteLine("elementId = " + ent.elementId);
                Console.WriteLine("======================");
            }

            Console.ReadKey();
        }
    }

    class entity
    {
        public string url { get; set; }
        public string component { get; set; }
        public string className { get; set; }
        public string elementId { get; set; }
    }
}
