using System;
using System.Collections.Generic;
using System.Text;

namespace Finances
{
    public enum TransactionType
    {
        CumpararePOS,
        TransferHomeBank,
        RetragereNumerar
    }

    public enum SpentOn
    {
        Food,
        Work,
        Car,
        Other
    }

    public class Transaction
    {
        public double? Debit { get; set; }
        public double? Credit { get; set; }
        public string Type { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public DateTime Date { get; set; }
        public int CalendarWeek { get; set; }
        public SpentOn SpendingType { get; set; }
        public TransactionType TypeOfTransaction {get;set;}
}
}
