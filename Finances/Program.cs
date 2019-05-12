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
            //List<Transaction> transactions = getTransactionsFromFile(@"Tranzactii_11-05-2019_23-36-40_card_comun.xls");
            //List<Transaction> transactions = getTransactionsFromFile(@"Tranzactii_11-05-2019_20-03-53.xls");
            List<Transaction> transactions = getTransactionsFromFile(@"Tranzactii_11-05-2019_16-32-26.xls");
            

            List<Transaction> categorizedTransactions = categorizeTransactions(transactions);

            List<List<Transaction>> transactionsByWeek = findWeeklyTransactions(categorizedTransactions);

            Console.WriteLine(GenerateReport(transactionsByWeek));

            Console.Read();
        }

    }
}
