using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Linq;
using Oracle.ManagedDataAccess.Client;

namespace COAProcessing
{


    [ComVisible(true)]
    [ProgId("DeletePdfPath.DeletePdfPathCls")]
    public class DeletePdfPathCls : IEntityExtension
    {

        #region private members
        private INautilusServiceProvider _sp;
        private OracleConnection _connection;
        private string sql = "";
        private OracleCommand cmd;
        private OracleDataReader reader;
        private string _connectionString;

        #endregion
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;

        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            _sp = Parameters["SERVICE_PROVIDER"];
            var records = Parameters["RECORDS"];
            Connect();
            while (!records.EOF)
            {

                var id = records.Fields["SDG_ID"].Value;

                string sql = string.Format("UPDATE LIMS_SYS.SDG_USER SET U_PDF_PATH=NULL WHERE SDG_ID='{0}'",
                                               id.ToString());
                cmd = new OracleCommand(sql, _connection);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                records.MoveNext();
            }
            _connection.Close();



        }
        public void Connect()
        {
            try
            {
                INautilusDBConnection dbConnection;
                if (_sp != null)
                {
                    dbConnection = _sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                }
                else
                {
                    dbConnection = null;
                }
                if (dbConnection != null)
                {
                    // _username= dbConnection.GetUsername();
                    _connection = GetConnection(dbConnection);
                    //set oracleCommand's connection
                    cmd = _connection.CreateCommand();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Err At Connect: " + e.Message);
            }
        }
        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {
                    _connectionString = ntlsCon.GetADOConnectionString();
                    var splited = _connectionString.Split(';');
                    _connectionString = "";
                    for (int i = 1; i < splited.Count(); i++)
                    {
                        _connectionString += splited[i] + ';';
                    }

                    //create connection
                    connection = new OracleConnection(_connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                        MessageBox.Show("Inconsistent role Security : " + f.Message);
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                    MessageBox.Show("Err At GetConnection: " + e.Message);
                }
            }
            return connection;
        }
    }
}
