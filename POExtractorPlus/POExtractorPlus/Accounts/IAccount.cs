using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POExtractorPlus.Accounts
{
    public interface IAccount
    {
        List<string> Extract(string destination, string[] files);
    }
}
