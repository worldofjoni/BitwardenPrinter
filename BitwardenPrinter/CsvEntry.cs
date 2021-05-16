using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BitwardenPrinter
{
    class CsvEntry
    {
        public String folder { get; set; }
        public String favorite{ get; set; }
        public String type{ get; set; }
        public String name{ get; set; }
        public String notes{ get; set; }
        public String fields{ get; set; }
        public String login_uri{ get; set; }
        public String login_username{ get; set; }
        public String login_password{ get; set; }
        public String login_totp{ get; set; }

    }
}
