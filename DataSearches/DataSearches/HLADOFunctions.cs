using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using HLFileFunctions;
using ADODB;

namespace HLADOFunctions
{
    class ADOFunctions
    {
        # region constructor

        FileFunctions myFileFuncs;

        public ADOFunctions()
        {
            myFileFuncs = new FileFunctions();
        }
        #endregion


        public ADODB.Connection OpenADOConnection(string aFullDatabaseName, string UserID = "", string PassWord = "", int Options = 0, bool Messages = false)
        {
            if (aFullDatabaseName == null)
            {
                if (Messages) MessageBox.Show("Please pass database name", "Open ADO Connection");
                return null;
            }

            if (!myFileFuncs.FileExists(aFullDatabaseName))
            {
                if (Messages) MessageBox.Show("Database " + aFullDatabaseName + " doesn't exist", "Open ADO Connection");
                return null;
            }

            string AccessConnect = "Provider='Microsoft.Jet.OLEDB.4.0';data source='" + aFullDatabaseName + "'";

            ADODB.Connection myConn = new ADODB.Connection();
            //myConn.ConnectionString = AccessConnect;
            
            myConn.Open(AccessConnect, UserID, PassWord, Options);
            return myConn;
        }

        public ADODB.Recordset OpenADORecordset(ref ADODB.Connection aConnection, string aQuery, bool Messages = false)
        // Note this opens for editing as well. May be some overheads to this.
        {
            if (aConnection == null | aQuery == null)
            {
                if (Messages) MessageBox.Show("Please pass required parameters", "Open ADO Recordset For Editing");
                return null;
            }

            ADODB.Recordset myRS = new ADODB.Recordset();
            myRS.Open(aQuery, aConnection, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic);

            return myRS;
        }
    }
}
