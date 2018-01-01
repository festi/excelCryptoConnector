﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CryptoConnector
{
    class ManualConnector : AccountConnector
    {
        public override string Name => "manual";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => false;
        public override bool SupportFills => false;

        protected override void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id)
        {
        }

        protected override List<AccountId> RefreshBalance_Internal(Worksheet sheet)
        {
            return new List<AccountId>();
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
        }
    }
}
