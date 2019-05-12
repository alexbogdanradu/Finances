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
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                for (int col = 0; col < sheet.GetRow(row).Cells.Count - 1; col++)
                {
                    NPOI.SS.UserModel.ICell cell = sheet.GetRow(row).GetCell(col);

                    if (cell.CellType != CellType.Blank)
                    {
                        if (col == 15)
                        {
                            if (cell.CellType == CellType.Numeric)
                            {
                                cell.SetCellType(CellType.String);
                                transactionList.Add(new Transaction());
                                transactionList[transactionList.Count - 1].Date = sheet.GetRow(row).GetCell(1).DateCellValue;
                                transactionList[transactionList.Count - 1].Debit = double.Parse(cell.StringCellValue);
                                transactionList[transactionList.Count - 1].To = sheet.GetRow(row + 2).GetCell(7).StringCellValue;
                                transactionList[transactionList.Count - 1].Type = sheet.GetRow(row).GetCell(7).StringCellValue;
                            }
                        }
                        if (col == 17)
                        {
                            if (cell.CellType == CellType.Numeric)
                            {
                                cell.SetCellType(CellType.String);
                                transactionList.Add(new Transaction());
                                transactionList[transactionList.Count - 1].Date = sheet.GetRow(row).GetCell(1).DateCellValue;
                                transactionList[transactionList.Count - 1].Credit = double.Parse(cell.StringCellValue);
                                transactionList[transactionList.Count - 1].From = sheet.GetRow(row + 1).GetCell(7).StringCellValue;
                                transactionList[transactionList.Count - 1].Type = sheet.GetRow(row).GetCell(7).StringCellValue;
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
                    else if (item.To.Contains("MEGA")) { item.SpendingType = SpentOn.Food; }
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

            List<Transaction> lastWeek = new List<Transaction>();
            List<List<Transaction>> last4Weeks = new List<List<Transaction>>();

            foreach (var week in byWeek)
            {
                report += ($"Week {week.First().CalendarWeek}{Environment.NewLine}");
                report += Environment.NewLine;

                report += ($"Total spent: {week.Sum(o => o.Debit)}{Environment.NewLine}");
                report += ($"Average per day spent: {(week.Sum(o => o.Debit) / 7)}{Environment.NewLine}");

                if (lastWeek.Count != 0)
                {
                    double dThisWeek = (double)week.Sum(o => o.Debit);
                    double dLastWeek = (double)lastWeek.Sum(o => o.Debit);
                    double diff = dThisWeek - dLastWeek;

                    string moreOrLess = "";

                    if (diff < 0)
                    {
                        moreOrLess = "less";
                    }
                    else
                    {
                        moreOrLess = "more";
                    }

                    report += ($"{(diff/dLastWeek * 100).ToString().Substring(0, 6)}% {moreOrLess} than last week.{Environment.NewLine}");
                }

                //report += ($"{(week.Sum(o => o.Debit) - week.Average(o => o.Debit))/week.Average(o => o.Debit) * 100}% more than average.{Environment.NewLine}");

                report += Environment.NewLine;

                foreach (var trans in Enum.GetValues(typeof(TransactionType)))
                {
                    report += Enum.GetName(typeof(TransactionType), trans) + ": " + week.Where(o => o.TypeOfTransaction == (TransactionType)trans).Sum(o => o.Debit) + Environment.NewLine;
                }

                report += Environment.NewLine;

                foreach (var trans in Enum.GetValues(typeof(SpentOn)))
                {
                    report += Enum.GetName(typeof(SpentOn), trans) + ": " + week.Where(o => o.SpendingType == (SpentOn)trans).Sum(o => o.Debit) + Environment.NewLine;
                }

                report += Environment.NewLine;

                //report += ("MEGA " + findSpendingsByContainingString(week, "MEGA") + Environment.NewLine);
                //report += ("CORA " + findSpendingsByContainingString(week, "CORA") + Environment.NewLine);
                //report += ("OMV " + findSpendingsByContainingString(week, "OMV") + Environment.NewLine);

                report += "Most spent on: " + Environment.NewLine;
                foreach (var ordonator in findSpendingsByUniqueDestination(week))
                {
                    report += ($"Total spendings at {ordonator.To}: {ordonator.Debit}{Environment.NewLine}");
                }

                report += Environment.NewLine;

                lastWeek.Clear();
                lastWeek = week;
            }

            return report;
        }
    }
}
