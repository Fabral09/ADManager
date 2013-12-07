//-------------------------------------------------------------------------
/*
/ Filename: ADManager.cs
/ Author: Fabrizio Alonzi
/ Date: 2010
*/
//-------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.SqlTypes;
using System.Threading;
using System.IO;
using System.Xml;

namespace test_oracle
{
    class ADManager
    {
        private DirectoryEntry entrypoint, nuovo_utente;
        private Queue<string> nomi_campi_active_directory = new Queue<string>();
        private Queue<string> nomi_campi_db = new Queue<string>();
        private TextWriter scrittore_file = new StreamWriter( DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + ".log" );
        int indice_campo_mail = 0;
        private int indice_utenti_totali = 0, indice_utenti_aggiornati = 0, indice_utenti_non_trovati = 0;
        private string db_adress = null, db_type = null, db_port = null, db_user_id = null, db_password = null, db_service_name = null, db_table = null, db_date_field = null, db_connection_string = null;
 
        private void ImpostaProprieta( DirectoryEntry de, string PropertyName, string PropertyValue )
        {
            if ( PropertyValue != null || PropertyValue.Trim() == "" )
            {
                if ( de.Properties.Contains( PropertyName ) )
                {
                    de.Properties[PropertyName][0] = ( PropertyValue != "" ) ? PropertyValue : " ";
                }
                else
                {
                    de.Properties[PropertyName].Add( ( PropertyValue != "" ) ? PropertyValue : " " );
                }
            }
        }

        private void AggiornaProprieta( DirectoryEntry de, string PropertyName, string PropertyValue )
        {
            try
            {
                if ( PropertyValue != null || PropertyValue.Trim() == "" )
                {
                    de.Properties[PropertyName].Value = ( ( PropertyValue != "" ) ? PropertyValue : " " );
                    de.CommitChanges();
                }
                
            }
            catch ( DirectoryServicesCOMException ex )
            {
                if ( ex.ErrorCode == -2147016682 ) de.Rename( PropertyName + "=" + PropertyValue );
            }
        }

        public void CaricaMappaturaDBAD()
        {
            int indice_campo = 0;

            foreach ( string chiave in System.Configuration.ConfigurationSettings.AppSettings.AllKeys )
            {
                string relazione = System.Configuration.ConfigurationSettings.AppSettings[chiave];
                string[] tmp_array = relazione.Split( '$' );
                nomi_campi_db.Enqueue( tmp_array[0].Trim() );
                nomi_campi_active_directory.Enqueue( tmp_array[1].Trim() );
                if ( tmp_array[1].Trim() == "mail" ) this.indice_campo_mail = indice_campo;
                indice_campo++;
            }
        }

        public void ImportaTabellaOracle( int indice_user_name )
        {

            string query = "SELECT ";

            foreach ( string s in nomi_campi_db )
                query += s + ", ";

            query = query.Remove( query.Length - 2, 1 );
            query += " from " + db_table;

            OracleConnection con = new OracleConnection( db_connection_string );
            con.Open();

            OracleCommand cmd = new OracleCommand( query, con );
            OracleDataReader reader = cmd.ExecuteReader();

            Queue<string> valori_campi = new Queue<string>();

            while ( reader.Read() )
            {
                for ( int i = 0; i < reader.FieldCount; i++ )
                    valori_campi.Enqueue( reader.GetValue( i ).ToString() );
                CreaNuovoUtente( reader.GetValue( indice_user_name ).ToString(), Guid.NewGuid().ToString(), valori_campi );
            }
        }

        public void ImportaTabellaSqlServer( int indice_user_name )
        {
            string query = "SELECT ";

            foreach ( string s in nomi_campi_db )
                query += s + ", ";

            query = query.Remove( query.Length - 2, 1 );
            query += " from " + db_table;

            SqlConnection con = new SqlConnection( db_connection_string );
            con.Open();

            SqlCommand cmd = new SqlCommand( query, con );
            SqlDataReader reader = cmd.ExecuteReader();

            Queue<string> valori_campi = new Queue<string>();

            while ( reader.Read() )
            {
                for ( int i = 0; i < reader.FieldCount; i++ )
                    valori_campi.Enqueue( reader.GetValue( i ).ToString() );
                CreaNuovoUtente( reader.GetValue( indice_user_name ).ToString(), Guid.NewGuid().ToString(), valori_campi );
                
            }
        }

        public void AggiornaActiveDirectoryDaSQLServer( string stringa_di_connessione, string nome_tabella, string nome_campo_data )
        {
            string query = "SELECT ";

            foreach ( string s in nomi_campi_db )
                query += s + ", ";

            query = query.Remove( query.Length - 2, 1 );
            query += " from " + db_table;

            SqlConnection con = new SqlConnection( this.db_connection_string );
            con.Open();
            SqlCommand cmd = new SqlCommand( "SELECT TO_CHAR(MAX(" + db_date_field + "), 'DD-MON-YYYY') AS DATA_ATTUALE FROM " + db_table, con );
            SqlDataReader reader = cmd.ExecuteReader();
            reader.Read();
            string data = reader.GetValue( 0 ).ToString();
            query += " where " + db_date_field + " = '" + data + "'";
            cmd.CommandText = query;
            reader = cmd.ExecuteReader();
            Queue<string> valori_campi = new Queue<string>();

            while ( reader.Read() )
            {
                for ( int i = 0; i < reader.FieldCount; i++ )
                    valori_campi.Enqueue( reader.GetValue( i ).ToString() );
                DirectorySearcher search = new DirectorySearcher( entrypoint );
                search.Filter = String.Format( "(mail={0})", valori_campi.ElementAt( indice_campo_mail ) );
                SearchResult result = search.FindOne();
                indice_utenti_totali++;
                if ( result != null )
                {
                    nuovo_utente = result.GetDirectoryEntry();
                    ModificaUtente( nomi_campi_active_directory, valori_campi );
                    valori_campi.Clear();
                    indice_utenti_aggiornati++;
                    Console.Clear();
                    Console.WriteLine( "\n\n  UTENTI AGGIORNATI: {0} \n\n", indice_utenti_aggiornati );
                }
                else
                {
                    scrittore_file.WriteLine( "UTENTE CON E-MAIL " + valori_campi.ElementAt( indice_campo_mail ).ToString() + " NON PRESENTE NELL'ACTIVE  DIRECTORY!" );
                    indice_utenti_non_trovati++;
                    valori_campi.Clear();
                }
            }
            scrittore_file.WriteLine();
            string sommatoria = String.Format( "UTENTI TOTALI: {0}  | UTENTI AGGIORNATI: {1}  |  UTENTI NON TROVATI IN AD: {2}", indice_utenti_totali, indice_utenti_aggiornati, indice_utenti_non_trovati );
            scrittore_file.WriteLine( sommatoria.ToString());
            scrittore_file.Flush();
            scrittore_file.Close();
        }

        public void AggiornaActiveDirectoryDaOracle()
        {
            string query = "SELECT ";

            foreach ( string s in nomi_campi_db )
                query += s + ", ";

            query = query.Remove( query.Length - 2, 1 );
            query += " from " + db_table;

            OracleConnection con = new OracleConnection( this.db_connection_string );
            con.Open();
            OracleCommand cmd = new OracleCommand( "SELECT TO_CHAR(MAX(" + db_date_field + "), 'DD-MON-YYYY') AS DATA_ATTUALE FROM " + db_table, con );
            OracleDataReader reader = cmd.ExecuteReader();
            reader.Read();
            string data = reader.GetValue( 0 ).ToString();
            query += " where " + db_date_field + " = '" + data + "'";
            cmd.CommandText = query;
            reader = cmd.ExecuteReader();
            Queue<string> valori_campi = new Queue<string>();

            while ( reader.Read() )
            {
                for ( int i = 0; i < reader.FieldCount; i++ )
                    valori_campi.Enqueue( reader.GetValue(i).ToString() );
                DirectorySearcher search = new DirectorySearcher( entrypoint );
                search.Filter = String.Format( "(mail={0})", valori_campi.ElementAt( indice_campo_mail ) );
                SearchResult result = search.FindOne();
                indice_utenti_totali++;
                if ( result != null )
                {
                    nuovo_utente = result.GetDirectoryEntry();
                    ModificaUtente( nomi_campi_active_directory, valori_campi );
                    valori_campi.Clear();
                    indice_utenti_aggiornati++;
                    Console.Clear();
                    Console.WriteLine( "\n\n  UTENTI AGGIORNATI: {0} \n\n", indice_utenti_aggiornati );
                }
                else
                {
                    scrittore_file.WriteLine( "UTENTE CON E-MAIL " + valori_campi.ElementAt( indice_campo_mail ).ToString() + " NON PRESENTE NELL'ACTIVE  DIRECTORY!" );
                    indice_utenti_non_trovati++;
                    valori_campi.Clear();
                }
            }
            scrittore_file.WriteLine();
            string sommatoria = String.Format( "UTENTI TOTALI: {0}  | UTENTI AGGIORNATI: {1}  |  UTENTI NON TROVATI IN AD: {2}", indice_utenti_totali, indice_utenti_aggiornati, indice_utenti_non_trovati );
            scrittore_file.WriteLine( sommatoria.ToString() );
            scrittore_file.Flush();
            scrittore_file.Close();
        }

        private void CreaNuovoUtente( string nome_utente, string password, Queue<string> valori_campi )
        {
            try
            {
                DirectoryEntry de = entrypoint;
                DirectoryEntries utenti = de.Children;
                DirectoryEntry nuovo_utente = utenti.Add( "CN=" + nome_utente, "user" );

                foreach ( string s in this.nomi_campi_active_directory )
                    ImpostaProprieta( nuovo_utente, s, valori_campi.Dequeue() );
                entrypoint.CommitChanges();

                nuovo_utente.CommitChanges();
                DirectoryEntry utente = new DirectoryEntry();
                utente.Path = nuovo_utente.Path;
                utente.AuthenticationType = AuthenticationTypes.Secure;
                Object ret = utente.Invoke( "SetPassword", password );
                utente.Properties["userAccountControl"].Value = 544;
                utente.CommitChanges();
                de.Close();
                this.nuovo_utente = nuovo_utente;
                Console.WriteLine( "UTENTE " + nome_utente + " INSERITO IN ACTIVE DIRECTORY!" );
            }  
            catch( DirectoryServicesCOMException ex )
            {
                if ( ex.ErrorCode == -2147019886 )
                    Console.WriteLine( "ERRORE! L'utente " + nome_utente + " esiste già nell'Active Directory!" );
                else
                    Console.WriteLine( ex.Message );
            }
                 
        }

        public void ModificaUtente( Queue<string> nomi_campi, Queue<string> valori_campi )
        {
            foreach ( string s in nomi_campi )
            {
                try
                {
                    AggiornaProprieta( this.nuovo_utente, s, valori_campi.Dequeue() );
                }
                catch ( Exception ex )
                {
                    Console.WriteLine( "IL CAMPO " + s + " NON PUO' ESSERE AGGIORNATO!" );
                }
                
            }
        }

        public ADManager( string ldap_adress, string username, string password )
        {
            entrypoint = new DirectoryEntry( ldap_adress, username, password );
        }

        public ADManager( string xml )
        {
            string ldap_adress, username, password;
            XmlDocument doc = new XmlDocument();
            XmlNodeList list = null;
            doc.Load( xml );
            list = doc.SelectNodes( "Settaggi" );
            ldap_adress = list.Item( 0 ).ChildNodes.Item( 0 ).InnerText;
            username = list.Item( 0 ).ChildNodes.Item( 1 ).InnerText;
            password = list.Item( 0 ).ChildNodes.Item( 2 ).InnerText;
            db_type = list.Item( 0 ).ChildNodes.Item( 3 ).InnerText;
            db_adress = list.Item( 0 ).ChildNodes.Item( 4 ).InnerText;
            db_service_name = list.Item( 0 ).ChildNodes.Item( 5 ).InnerText;
            db_port = list.Item( 0 ).ChildNodes.Item( 6 ).InnerText;
            db_user_id = list.Item( 0 ).ChildNodes.Item( 7 ).InnerText;
            db_password = list.Item( 0 ).ChildNodes.Item( 8 ).InnerText;
            db_table = list.Item( 0 ).ChildNodes.Item( 9 ).InnerText;
            db_date_field = list.Item( 0 ).ChildNodes.Item( 10 ).InnerText;

            if ( db_type.ToString().ToUpper() == "ORACLE" )
                db_connection_string = String.Format( "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST={0})(PORT={1})))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME={2})));User ID={3}; Password={4}", db_adress, db_port, db_service_name, db_user_id, db_password );

            entrypoint = new DirectoryEntry( ldap_adress, username, password );
        }
    }
}
