using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Data;
using Walkershop.EDI.Converter;

namespace Walkershop.EDI.Converter.Data.ManagerObjects
{
    public class ManagerBase
    {
        private static Database database;

        private string CONNECTION_STRING_NAME = "EDIConverter";
        
        protected Database CurrentDB
        {
            get
            {
                if (database == null)
                {
                    string connectionString = string.IsNullOrEmpty(WalkerApplication.CONNECTION_STRING_NAME) ? CONNECTION_STRING_NAME : WalkerApplication.CONNECTION_STRING_NAME;
                    database = DatabaseFactory.CreateDatabase(connectionString);
                }

                return database;
            }
        }
    }
}
