using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.SqlTypes;
using System.Collections.Specialized;
using System.Threading;

namespace test_oracle
{
    class Program
    {
        static void Main(string[] args)
        {
            ADManager nuovo = new ADManager("parametri_connessione.xml");
            nuovo.CaricaMappaturaDBAD();
            nuovo.AggiornaActiveDirectoryDaOracle();
            Console.WriteLine("Fine procedura");
            Thread.Sleep(3000);
        }
    }
}
