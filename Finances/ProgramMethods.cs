using MailKit.Net.Pop3;
using MimeKit;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;

namespace Finances
{
    partial class Program
    {
        //private static void GetEmail()
        //{
        //    using (var emailClient = new Pop3Client())
        //    {
        //        emailClient.Connect(.PopServer, _emailConfiguration.PopPort, true);

        //        emailClient.AuthenticationMechanisms.Remove("XOAUTH2");

        //        emailClient.Authenticate(_emailConfiguration.PopUsername, _emailConfiguration.PopPassword);

        //        List<MailMessage> emails = new List<EmailMessage>();
        //        for (int i = 0; i < emailClient.Count && i < maxCount; i++)
        //        {
        //            var message = emailClient.GetMessage(i);
        //            var emailMessage = new EmailMessage
        //            {
        //                Content = !string.IsNullOrEmpty(message.HtmlBody) ? message.HtmlBody : message.TextBody,
        //                Subject = message.Subject
        //            };
        //            emailMessage.ToAddresses.AddRange(message.To.Select(x => (MailboxAddress)x).Select(x => new EmailAddress { Address = x.Address, Name = x.Name }));
        //            emailMessage.FromAddresses.AddRange(message.From.Select(x => (MailboxAddress)x).Select(x => new EmailAddress { Address = x.Address, Name = x.Name }));
        //        }
        //    }
        //}

        private static List<Transaction> getTransactionsFromFile(string path)
        {
            List<Transaction> transactionList = new List<Transaction>();

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("transactions");

            int iDateColumnIndex = -1;
            int iTransactionDetailsColumnIndex = -1;
            int iDebitColumnIndex = -1;
            int iCreditColumnIndex = -1;

            for (int col = 0; col < sheet.GetRow(1).Cells.Count - 1; col++)
            {
                NPOI.SS.UserModel.ICell cell = sheet.GetRow(1).GetCell(col);

                if (cell.StringCellValue == "Data")
                {
                    Console.WriteLine($"Date found on column {col}");
                    iDateColumnIndex = col;
                }

                if (cell.StringCellValue == "Detalii tranzactie")
                {
                    Console.WriteLine($"Transaction details found on column {col}");
                    iTransactionDetailsColumnIndex = col;
                }

                if (cell.StringCellValue == "Debit")
                {
                    Console.WriteLine($"Debit found on column {col}");
                    iDebitColumnIndex = col;
                }

                if (cell.StringCellValue == "Credit")
                {
                    Console.WriteLine($"Credit found on column {col}");
                    iCreditColumnIndex = col;
                }
            }

            if (iDateColumnIndex == -1 || iTransactionDetailsColumnIndex == -1 || iDebitColumnIndex == -1 || iCreditColumnIndex == -1)
            {
                throw new Exception("Could not find one of the columns.");
            }

            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                for (int col = 0; col < sheet.GetRow(row).Cells.Count - 1; col++)
                {
                    NPOI.SS.UserModel.ICell cell = sheet.GetRow(row).GetCell(col);

                    if (cell.CellType != CellType.Blank)
                    {
                        if (col == iDebitColumnIndex)
                        {
                            if (cell.CellType == CellType.Numeric)
                            {
                                cell.SetCellType(CellType.String);
                                transactionList.Add(new Transaction());
                                transactionList[transactionList.Count - 1].Date = sheet.GetRow(row).GetCell(iDateColumnIndex).DateCellValue;
                                transactionList[transactionList.Count - 1].Debit = double.Parse(cell.StringCellValue, CultureInfo.InvariantCulture);
                                transactionList[transactionList.Count - 1].To = sheet.GetRow(row + 2).GetCell(iTransactionDetailsColumnIndex).StringCellValue;
                                transactionList[transactionList.Count - 1].Type = sheet.GetRow(row).GetCell(iTransactionDetailsColumnIndex).StringCellValue;
                            }
                        }
                        if (col == iCreditColumnIndex)
                        {
                            if (cell.CellType == CellType.Numeric)
                            {
                                cell.SetCellType(CellType.String);
                                transactionList.Add(new Transaction());
                                transactionList[transactionList.Count - 1].Date = sheet.GetRow(row).GetCell(iDateColumnIndex).DateCellValue;
                                transactionList[transactionList.Count - 1].Credit = double.Parse(cell.StringCellValue, CultureInfo.InvariantCulture);
                                transactionList[transactionList.Count - 1].From = sheet.GetRow(row + 1).GetCell(iTransactionDetailsColumnIndex).StringCellValue;
                                transactionList[transactionList.Count - 1].Type = sheet.GetRow(row).GetCell(iTransactionDetailsColumnIndex).StringCellValue;
                            }
                        }
                    }
                }
            }
            return transactionList;
        }

        private static double findSpendingsByContainingString(List<Transaction> c, string v)
        {
            double spendings = 0;
            foreach (var item in c)
            {
                if (item.Debit != null && item.To.Contains(v))
                {
                    spendings += item.Debit.Value;
                }
            }

            return spendings;
        }

        private static double findSpendingsByOrdonator(List<Transaction> c, string v)
        {
            double spendings = 0;
            foreach (var item in c)
            {
                if (item.Debit != null && item.To == v)
                {
                    spendings += item.Debit.Value;
                }
            }
            return spendings;
        }

        private static List<Transaction> findSpendingsByUniqueDestination(List<Transaction> t)
        {
            List<Transaction> response = new List<Transaction>();
            List<string> uniqueOrdonators = new List<string>();
            var debits = t.FindAll(o => o.To != null).ToList();

            foreach (var debit in debits)
            {
                uniqueOrdonators.Add(debit.To);
            }

            foreach (var item in uniqueOrdonators.Distinct())
            {
                double spendings = findSpendingsByOrdonator(t, item);
                response.Add(new Transaction
                {
                    To = item,
                    Debit = spendings
                });
            }

            return response.OrderBy(o => o.Debit).Reverse().ToList();
        }

        private static List<List<Transaction>> findWeeklyTransactions(List<Transaction> _transactions)
        {
            List<List<Transaction>> weeklyTransactions = new List<List<Transaction>>();
            CultureInfo cul = CultureInfo.CurrentCulture;
            List<int> cws = new List<int>();

            foreach (var item in _transactions)
            {
                item.CalendarWeek = cul.Calendar.GetWeekOfYear(item.Date, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
                cws.Add(item.CalendarWeek);
            }

            foreach (var cw in cws.Distinct().OrderBy(o => o))
            {
                weeklyTransactions.Add(new List<Transaction>());
                foreach (var transaction in _transactions)
                {
                    if (transaction.CalendarWeek == cw)
                    {
                        weeklyTransactions.Last().Add(transaction);
                    }
                }
            }

            return weeklyTransactions;
        }

        private static List<Transaction> categorizeTransactions(List<Transaction> transactions)
        {
            foreach (var item in transactions)
            {
                switch (item.Type)
                {
                    case "Cumparare POS":
                        item.TypeOfTransaction = TransactionType.CumpararePOS;
                        break;
                    case "Retragere numerar":
                        item.TypeOfTransaction = TransactionType.RetragereNumerar;
                        break;
                    case "Transfer Home'Bank":
                        item.TypeOfTransaction = TransactionType.TransferHomeBank;
                        break;
                    default:
                        break;
                }
            }

            foreach (var item in transactions)
            {
                if (item.To != null)
                {
                    if (item.To.Contains("KAUFLAND")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("CORA")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("CARREFOUR")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("LIDL")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("PROFI")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("CHOPSTIX")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("TIMAS")) { item.SpendingType = SpentOn.Car; }
                    else if (item.To.Contains("MEGA")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("AUCHAN")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("PREMIER")) { item.SpendingType = SpentOn.Food; }
                    else if (item.To.Contains("EAT ETC")) { item.SpendingType = SpentOn.Work; }
                    else if (item.To.Contains("ERIC STEFAN")) { item.SpendingType = SpentOn.Work; }
                    else if (item.To.Contains("INMEDIO")) { item.SpendingType = SpentOn.Work; }
                    else if (item.To.Contains("LUCA")) { item.SpendingType = SpentOn.Work; }
                    else if (item.To.Contains("OMV")) { item.SpendingType = SpentOn.Car; }
                    else if (item.To.Contains("AUTOKARMA")) { item.SpendingType = SpentOn.Car; }
                    else if (item.To.Contains("LAGARDERE")) { item.SpendingType = SpentOn.Work; }
                    else
                    {
                        item.SpendingType = SpentOn.Other;
                    }
                }
            }

            return transactions;
        }

        private static string GenerateReport(List<List<Transaction>> byWeek)
        {
            string report = "";
            string moreOrLess = "";
            string incOrDec = "";

            List<Transaction> lastWeek = new List<Transaction>();
            List<List<Transaction>> last4Weeks = new List<List<Transaction>>();

            double? totalSpent = 0;

            for (int i = 0; i < byWeek.Count; i++)
            {
                totalSpent += byWeek[i].Sum(o => o.Debit);
            }

            double averagePerWeek = (double)totalSpent/byWeek.Count;

            foreach (var week in byWeek)
            {
                //Total spent
                report += ($"Week {week.First().CalendarWeek}{Environment.NewLine}");
                report += ($"Total spent: {week.Sum(o => o.Debit)} RON. ");

                //Percentage versus weekly average
                if (lastWeek.Count != 0)
                {
                    double dThisWeek = (double)week.Sum(o => o.Debit);
                    //double dLastWeek = (double)lastWeek.Sum(o => o.Debit);
                    double dDiff = dThisWeek - averagePerWeek;
                    double dPercent = (dDiff / averagePerWeek * 100);
                    int iPercent = (int)Math.Round(dPercent);
                    string sPercent = iPercent.ToString();

                    if (dDiff < 0)
                    {
                        moreOrLess = "\u2193";
                        incOrDec = "decrease";
                        sPercent = sPercent.Replace("-", "");
                    }
                    else
                    {
                        moreOrLess = "\u2191";
                        incOrDec = "increase";
                    }

                    report += ($" {moreOrLess} {sPercent}%. Weekly average: {(int)Math.Round(averagePerWeek)}RON.{Environment.NewLine}");
                }

                //string averagePerDay = (week.Sum(o => o.Debit) / 7).ToString();
                //averagePerDay = averagePerDay.Substring(0, averagePerDay.IndexOf(".") + 2);

                //report += ($"Average per day spent: {averagePerDay} RON.{Environment.NewLine}");
                //report += ($"{(week.Sum(o => o.Debit) - week.Average(o => o.Debit)) / week.Average(o => o.Debit) * 100}% more than average.{Environment.NewLine}");

                report += Environment.NewLine;

                //foreach (var trans in Enum.GetValues(typeof(TransactionType)))
                //{
                //    report += Enum.GetName(typeof(TransactionType), trans) + ": " + week.Where(o => o.TypeOfTransaction == (TransactionType)trans).Sum(o => o.Debit) + " RON ";
                //    if (lastWeek.Count != 0)
                //    {
                //        double dThisWeek = (double)week.Where(o => o.TypeOfTransaction == (TransactionType)trans).Sum(o => o.Debit);
                //        double dLastWeek = (double)lastWeek.Where(o => o.TypeOfTransaction == (TransactionType)trans).Sum(o => o.Debit);
                //        double dDiff = dThisWeek - dLastWeek;

                //        if (dDiff != 0)
                //        {
                //            if (dDiff > 0)
                //            {
                //                incOrDec = "increase";
                //            }
                //            else
                //            {
                //                incOrDec = "decrease";
                //            }
                //            double dPercent = (double)(dDiff / dLastWeek * 100);
                //            string sPercent = dPercent.ToString().Substring(0, dPercent.ToString().IndexOf(".") + 2);
                //            report += ($"{sPercent}% {incOrDec}.");
                //        }
                //    }
                //    report += Environment.NewLine;
                //}

                //report += Environment.NewLine;

                //foreach (var trans in Enum.GetValues(typeof(SpentOn)))
                //{
                //    report += Enum.GetName(typeof(SpentOn), trans) + ": " + week.Where(o => o.SpendingType == (SpentOn)trans).Sum(o => o.Debit) + " RON ";
                //    if (lastWeek.Count != 0)
                //    {
                //        double dThisWeek = (double)week.Where(o => o.SpendingType == (SpentOn)trans).Sum(o => o.Debit);
                //        double dLastWeek = (double)lastWeek.Where(o => o.SpendingType == (SpentOn)trans).Sum(o => o.Debit);
                //        double dDiff = dThisWeek - dLastWeek;

                //        if (dDiff != 0)
                //        {
                //            if (dDiff > 0)
                //            {
                //                incOrDec = "increase";
                //            }
                //            else
                //            {
                //                incOrDec = "decrease";
                //            }
                //            double dPercent = (double)(dDiff / dLastWeek * 100);
                //            string sPercent = dPercent.ToString().Substring(0, dPercent.ToString().IndexOf(".") + 2);
                //            report += ($"{sPercent}% {incOrDec}.");
                //        }
                //    }
                //    report += Environment.NewLine;
                //}

                //report += Environment.NewLine;

                report += "Most spent on: " + Environment.NewLine;
                foreach (var ordonator in findSpendingsByUniqueDestination(week))
                {
                    //if (lastWeek.Count != 0)
                    //{
                    //    double dThisWeek = (double)ordonator.Debit;
                    //    double dLastWeek = (double)findSpendingsByUniqueDestination(lastWeek).Where(o => o.To == ordonator.To).Sum(o => o.Debit);
                    //    double dDiff = dThisWeek - dLastWeek;

                    //    if (dDiff != 0)
                    //    {
                    //        if (dDiff > 0)
                    //        {
                    //            incOrDec = "increase";
                    //        }
                    //        else
                    //        {
                    //            incOrDec = "decrease";
                    //        }
                    //        double dPercent = (double)(dDiff / dLastWeek * 100);
                    //        string sPercent = dPercent.ToString().Substring(0, dPercent.ToString().IndexOf(".") + 2);
                    //        report += ($"{sPercent}% {incOrDec}.");
                    //    }
                    //}
                    report += ($"{ordonator.To}: {ordonator.Debit} RON ");
                    report += Environment.NewLine;
                }

                report += Environment.NewLine;

                lastWeek.Clear();
                lastWeek = week;
            }

            return report;
        }
    }
}
