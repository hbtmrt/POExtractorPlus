﻿using POExtractorPlus.Accounts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POExtractorPlus
{
    public class Controller
    {
        public List<string> Extract(string accountType, string[] files, string destination)
        {
            IAccount account = AccountFactory.Create(accountType);
            return account.Extract(destination, files);
        }
    }
}
