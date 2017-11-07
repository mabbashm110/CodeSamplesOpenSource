using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailMarketing.Library
{
    public class CPerson
    {
        public string ContactPerson { get; set; }
        public string EmailAddress { get; set; }
        public string Tel { get; set; }
        public DateTime? SubscriptionDate { get; set; }
        public EmailStatus Status { get; set; }
    }

    public enum EmailStatus
    {
        Subscribed,
        Pending,
        Processing,
        Sent,
        Unsubscribed,
        Error
    }
}
