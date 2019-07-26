using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Finances
{
    partial class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<Transaction> transactions = getTransactionsFromFile(@"CardComun_24-07-2019_10-00-28.xls");
                //List<Transaction> transactions = getTransactionsFromFile(@"CardPersonal_24-07-2019_12-45-50.xls");

                List<Transaction> categorizedTransactions = categorizeTransactions(transactions);

                List<List<Transaction>> transactionsByWeek = findWeeklyTransactions(categorizedTransactions);

                List<List<Transaction>> transactionsByMonth = findMonthlyTransactions(categorizedTransactions);

                string weeklyReport = GenerateReport(transactionsByWeek, ReportType.Weekly);
                //Console.WriteLine(weeklyReport);

                using (StreamWriter sw = new StreamWriter("weeklyReport.txt"))
                {
                    sw.Write(weeklyReport);
                    sw.Flush();
                    sw.Close();
                }

                string monthlyReport = GenerateReport(transactionsByMonth, ReportType.Monthly);
                //Console.WriteLine(monthlyReport);

                using (StreamWriter sw = new StreamWriter("monthlyReport.txt"))
                {
                    sw.Write(monthlyReport);
                    sw.Flush();
                    sw.Close();
                }

                Console.Read();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.Read();
            }
        }
    }
}
