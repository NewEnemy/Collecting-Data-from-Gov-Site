using Microsoft.VisualBasic;
using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Net;
using System.Threading;

namespace PDFDownloadAndConvert
{
    class DownloadFiles
    {

        public static class NonBlockingConsole
        {
            private static BlockingCollection<string> m_Queue = new BlockingCollection<string>();


            static NonBlockingConsole()
            {
                var thread = new Thread(
                    () => {

                        while (true) Console.WriteLine(m_Queue.Take());
                    }


                    );
                thread.IsBackground = true;
                thread.Start();
            }
            public static void WriteLine(string value)
            {
                m_Queue.Add(value);
            }
        }

        static DateTime date_now = DateTime.Today;
        static string date = "19.10.2020";
        static DateTime past_date = DateTime.Today;

        DownloadFiles(int range)
        {

            WebClient webclient = new WebClient();



            for (int i = 0; i < range; i++)
            {

                past_date = past_date.AddDays(-1);
                date = past_date.AddDays(-1).ToString("dd.MM.yyyy");
                NonBlockingConsole.WriteLine(date);
                webclient.DownloadFile(new System.Uri($"https://www.psse.bielsko.pl/pdf/Raport_gminy_Bielsko_{date}.pdf"), $"F:\\PDFS\\{date}.pdf");
                NonBlockingConsole.WriteLine($"{i} of {range} \n Date {date}");
            }
        }
    }
    
}
