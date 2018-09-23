using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MySendEmail.Models
{
    public class EmailModel
    {
        public string Sender { get; set; }
        public string Receiver { get; set; }
        public string CarbonCopy { get; set; }
        public string SendTime { get; set; }
        //OK=1; NG=0
        public int SendState { get; set; } 
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Attachment { get; set; }

    }
}
