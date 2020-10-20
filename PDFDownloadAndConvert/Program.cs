using Aspose.Pdf;
using ExcelDataReader;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Linq;
using System.Collections.Generic;

namespace PDFDownloadAndConvert
{
    public class VirusDataPerGmina
    {
        public string gmina { get; set; }
        public string wynikDodatni { get; set; }
        public string podKwarantanna { get; set; }
        public string nadzor_epid { get; set; }
        public string hospitalizacja { get; set; }
        public string ozdrowiency { get; set; }

        public string zgony { get; set; }
    }
    public class WholeRecords
    {
        public string Data{get;set;}
        public List<VirusDataPerGmina> records { get; set; }

        public WholeRecords(string Data)
        {
            this.Data = Data;
        }
    }

    public class MasterRoot
    {
        public List<WholeRecords> MasterRecord { get; set; }
    }
    
    public  class StringHelpers
    {
        string _str;
        public StringHelpers()
        {
        }
        public  string PreformStringCleanup(string str)
        {
            this._str = str;
            this.ReplaceNullWithZero();
            this.TrimAfterSlash();
            return this._str;
        }
        private  void ReplaceNullWithZero()
        {
            if(this._str == ""|| this._str == null) { this._str= "0"; }

        }
        private void TrimAfterSlash()
        {
            if (this._str.Contains('/'))
            {
               
                string[] strArray = this._str.Split('/');
                
               this._str= strArray[0];
            }
        }
    }
    

    
    class Program
    {

        public static void ConvertAndSaveAllFiles()
        {
           string Path = "F:\\PDFS";
            string SavePath = "F:\\XLSFiles";


            ExcelSaveOptions opt = new ExcelSaveOptions();
            opt.Format = ExcelSaveOptions.ExcelFormat.XLSX;

            string[] filesList = Directory.GetFiles(Path);
            int i = 0;
            foreach (var file in filesList)
            {
                i++;
                var name = file.Split('\\');
                name = name[name.Length - 1].Split(".p");

                Document pdf = new Document(file);
                pdf.Save(SavePath+'\\'+ name[0]+".xlsx", opt);
                Console.WriteLine($"Processing file {file} \n {i} of {filesList.Length}");
                pdf.Dispose();



            }
        }

        static void Main(string[] args)
        {


            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            StringHelpers stringHelper = new StringHelpers();

            MasterRoot masterRoot = new MasterRoot();
            masterRoot.MasterRecord = new List<WholeRecords>();

            string[] fileList = Directory.GetFiles("F:\\XLSFiles\\");
            foreach (var file in fileList)
            {
                var name = file.Split('\\');
                name = name[name.Length - 1].Split(".x");

                WholeRecords allRecords = new WholeRecords(name[0]);
                allRecords.records = new List<VirusDataPerGmina>();




                using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var startRecorging = false;


                    do
                    {
                        while (reader.Read())
                        {
                            if (reader.GetString(0).Contains("Bielsko-Biała"))
                            {
                                startRecorging = true;
                            }
                  
                            if (startRecorging)
                            {

                                allRecords.records.Add(new VirusDataPerGmina
                                {
                                    gmina = reader.GetString(0),
                                    wynikDodatni = stringHelper.PreformStringCleanup(reader.GetString(1)),
                                    podKwarantanna = stringHelper.PreformStringCleanup(reader.GetString(2)),
                                    nadzor_epid = stringHelper.PreformStringCleanup(reader.GetString(3)),
                                    hospitalizacja = stringHelper.PreformStringCleanup(reader.GetString(4)),
                                    ozdrowiency = stringHelper.PreformStringCleanup(reader.GetString(5)),
                                    zgony = stringHelper.PreformStringCleanup(reader.GetString(6)),


                                });
                                if (reader.GetString(0).Contains("RAZEM",StringComparison.OrdinalIgnoreCase))
                                {

                                    break;
                                }
                            }
                           
                        }
                    } while (reader.NextResult());


                    }
                   
                    masterRoot.MasterRecord.Add(allRecords);


                }

                var jstring = System.Text.Json.JsonSerializer.Serialize(masterRoot);
                System.IO.File.WriteAllText(@"F:\PDFS\data.json", jstring);

            }

        }
    }
}
