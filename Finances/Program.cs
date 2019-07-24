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

                List<Transaction> categorizedTransactions = categorizeTransactions(transactions);

                List<List<Transaction>> transactionsByWeek = findWeeklyTransactions(categorizedTransactions);

                Console.WriteLine(GenerateReport(transactionsByWeek));

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
