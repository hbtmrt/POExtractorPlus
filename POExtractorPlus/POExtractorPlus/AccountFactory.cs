using POExtractorPlus.Accounts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POExtractorPlus
{
    public class AccountFactory
    {
        public static IAccount Create(string type) {
            switch (type) {
                case "BRAZIL":
                    return new Brazil();
                case "EU-CANADA-USA-APD":
                    return new EUCanadaUsa();
                case "LACL-Mexico":
                    return new LaclMexico();
                default:
                    throw new Exception("Cannot happen this.");
            }
        }
    }
}
