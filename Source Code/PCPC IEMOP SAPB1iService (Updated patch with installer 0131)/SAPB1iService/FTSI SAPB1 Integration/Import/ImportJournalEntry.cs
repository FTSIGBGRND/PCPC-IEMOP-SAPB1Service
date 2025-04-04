using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPbobsCOM;
using System.Windows;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Data.SqlClient;

namespace SAPB1iService
{
    class ImportJournalEntry
    {
        public static void _ImportJournalEntry()
        {
            ImportUserDefinedJournalEntry.importFromReversal();
        }
    }
}
