using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlServerCe;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Mercure
{
    class GestionBDD
    {
        //enum FichierType { Excel, Sdf }
        public string ConnSqlStr{get;set;}
        public string ConnDbStr {get;set;}
        public OleDbConnection ConnOleDb {get;set;}
        public OleDbCommand CmdOleDb {get;set;}
        public OleDbDataReader ReaderExcel {get;set;}

        public SqlCeConnection ConnSql {get;set;}
        public SqlCeCommand CmdRecherche {get;set;}
        public SqlCeDataReader RdrReche{get;set;}
        
        /// <summary>
        /// lecture du fichier excel et stockage les contenu dans la base de données .sdf
        /// </summary>
        /// <param name="PathExcel">le complète path du fichier Excel</param>
        /// <param name="PathSdf">le complète path du fichier Sdf</param>
        public void LectureExcel(string PathExcel, string PathSdf)
        {

            Connection(PathExcel, PathSdf);//connecter le fichier .xls et le fichier .sdf
            try
             {    
                while (ReaderExcel.Read())
                {
                    //Console.Out.WriteLine(ReaderExcel.GetString(0) + ";" + ReaderExcel.GetString(1) + ";" + ReaderExcel.GetString(2) + ";" + ReaderExcel.GetString(3) + ";" + ReaderExcel.GetString(4) + ";" + ReaderExcel.GetDouble(5));
                    
                    SqlCeCommand CmdSql = ConnSql.CreateCommand();

                    InsererMarques(ReaderExcel.GetString(2)); // Insérer Marques

                    InsererFamilles(ReaderExcel.GetString(3));// Insérer Familles

                    InsererSousFamilles(ReaderExcel.GetString(4), ReaderExcel.GetString(3));// Insérer SousFamilles

                    InsererArticles(ReaderExcel.GetString(0), ReaderExcel.GetString(1), ReaderExcel.GetDouble(5),
                        ReaderExcel.GetString(2), ReaderExcel.GetString(4));// Insérer Articles
                    
                }
                ReaderExcel.Close();
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
            }
            finally
            {
                if (ConnOleDb.State == ConnectionState.Open)
                {
                    ConnOleDb.Close();
                    ConnOleDb.Dispose();
                }
                if (ConnSql.State == ConnectionState.Open)
                {
                    ConnSql.Close();
                    ConnSql.Dispose();
                }
            }
            Console.In.Read();
        }

        /// <summary>
        /// connecter le fichier .xls et le fichier .sdf et remplir le OleDbDataReader
        /// </summary>
        /// <param name="PathExcel"></param>
        /// <param name="PathSdf"></param>
        private void Connection(string PathExcel, string PathSdf)
        {
            //connecter avec le fichier .xls
            ConnDbStr = GetConnectionString(PathExcel);
            ConnOleDb = new OleDbConnection(ConnDbStr);
            ConnOleDb.Open();
            Console.Out.WriteLine(PathExcel+" est ouvert.");

            //obtenir le nom du sheet de .xls
            DataTable DtExcelSchema = ConnOleDb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string SheetName = DtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            string StrExcel = "select * from [" + SheetName + "];";

            //obtenir des contenus du .xls et mettre dans OleDbDataReader
            CmdOleDb = new OleDbCommand(StrExcel, ConnOleDb);
            ReaderExcel = CmdOleDb.ExecuteReader();
            Console.Out.WriteLine("OleDbDataReader est rempli.");

            //connecter avec le fichier .sdf
            ConnSqlStr = GetConnectionString(PathSdf);
            ConnSql = new SqlCeConnection(ConnSqlStr);
            ClearAllData(); //supprimer toutes les données dans tous les tableaux de la base de données
            Console.Out.WriteLine("Toutes les données sont supprimées.");
            ConnSql.Open();
            Console.Out.WriteLine(ConnSqlStr + " est ouvert.");
        }

        /// <summary>
        /// construire le string de connection
        /// </summary>
        /// <param name="Path">le complète path du data source</param>
        /// <returns>le string de la connection</returns>
        private string GetConnectionString(string Path)
        {
            FileSystemInfo FileInfo = new FileInfo(Path);
            string Extension = FileInfo.Extension;
            Dictionary<string, string> Props = new Dictionary<string, string>();
            switch (Extension)
            {
                // XLSX - Excel 2007, 2010, 2012, 2013
                /*props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
                props["Data Source"] = Path;*/
                case ".xls":
                    // XLS - Excel 2003 and Older
                    Props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                    Props["Extended Properties"] = "Excel 8.0";
                    Props["Data Source"] = Path;
                    break;

                case ".sdf":
                    Props["Data Source"] = Path;
                    break;
            }
            StringBuilder SB = new StringBuilder();

            foreach (KeyValuePair<string, string> Prop in Props)
            {
                SB.Append(Prop.Key);
                SB.Append('=');
                SB.Append(Prop.Value);
                SB.Append(';');
            }

            return SB.ToString();
        }

        /// <summary>
        /// Insérer les données de Marques dans la base de données, si ce nom exsite déjà,
        /// on ne l'ajoute pas 
        /// </summary>
        /// <param name="Value">La valeur de nom du Marque</param>
        private void InsererMarques(string Value)
        {
            SqlCeCommand CmdSql = ConnSql.CreateCommand();
            Value = Regex.Replace(Value, @"\s{2,}", " "); 
            if (!DonneeIsExist("Marques", "Nom", Value))
            {
                CmdSql.CommandText = "INSERT INTO Marques(Nom) values(@MNom)";
                CmdSql.Parameters.AddWithValue("@MNom", Value);
                CmdSql.ExecuteNonQuery();
                Console.Out.WriteLine("INSERT INTO Marques(Nom) values(" + Value + ")");
            }
            
        }

        /// <summary>
        /// Insérer les données de Familles dans la base de données, si ce nom exsite déjà,
        /// on ne l'ajoute pas
        /// </summary>
        /// <param name="Value">La valeur de nom du Familles<</param>
        private void InsererFamilles(string Value)
        {
            SqlCeCommand CmdSql = ConnSql.CreateCommand();
            Value = Regex.Replace(Value, @"\s+", " ");
            if (!DonneeIsExist("Familles", "Nom", Value))
            {
                CmdSql.CommandText = "INSERT INTO Familles(Nom) values(@FNom)";
                CmdSql.Parameters.AddWithValue("@FNom", Value);
                CmdSql.ExecuteNonQuery();
                Console.Out.WriteLine("INSERT INTO Familles(Nom) values(" + Value + ")");
            }
            
        }

        /// <summary>
        /// Insérer les données de SousFamilles dans la base de données, si ce nom exsite déjà,
        /// on ne l'ajoute pas
        /// </summary>
        /// <param name="Value">La valeur de nom du SousFamille</param>
        /// <param name="Famille">La valeur de nom du Famille</param>
        private void InsererSousFamilles(string Value, string Famille)
        {
            SqlCeCommand CmdSql = ConnSql.CreateCommand();
            Value = Regex.Replace(Value, @"\s{1,}", " "); 
            if (!DonneeIsExist("SousFamilles", "Nom", Value))
            {
                CmdSql.CommandText = "INSERT INTO SousFamilles(Nom,IDFamille) values(@SFNom,@IDFam)";
                CmdSql.Parameters.AddWithValue("@SFNom", Value);
                CmdSql.Parameters.AddWithValue("@IDFam", GetIdByName("Familles", Famille));
                CmdSql.ExecuteNonQuery();
                Console.Out.WriteLine("INSERT INTO SousFamilles(Nom,IDFamille) values(" + Value + "," + GetIdByName("Familles", Famille) + ")");
            }
            
        }

        /// <summary>
        /// Insérer les données de Articles dans la base de données, si la référence exsite déjà,
        /// on ne l'ajoute pas 
        /// </summary>
        /// <param name="Descrip">La valeur de la description</param>
        /// <param name="Ref">La valeur de la référence</param>
        /// <param name="Prix">La valeur de prix</param>
        /// <param name="Marque">La valeur de nom du Marque</param>
        /// <param name="SousFamille">La valeur de nom du SousFamille</param>
        private void InsererArticles(string Descrip, string Ref, double Prix, string Marque, string SousFamille)
        {
            SqlCeCommand CmdSql = ConnSql.CreateCommand();
            if (!DonneeIsExist("Articles", "Ref", Ref))
            {
                CmdSql.CommandText = "INSERT INTO Articles(Description,Ref,Prix,IDMarque,IDSousFamille)" 
                                    +"values(@ADspt,@ARef,@APrix,@AIDMq,@AIDSsFm);";
                CmdSql.Parameters.AddWithValue("@ADspt", Descrip);
                CmdSql.Parameters.AddWithValue("@ARef", Ref);
                CmdSql.Parameters.AddWithValue("@APrix", Prix);
                CmdSql.Parameters.AddWithValue("@AIDMq", GetIdByName("Marques", Marque));
                CmdSql.Parameters.AddWithValue("@AIDSsFm", GetIdByName("SousFamilles", SousFamille));
                CmdSql.ExecuteNonQuery();
                Console.Out.WriteLine("INSERT INTO Articles(Description,Ref,Prix,IDMarque,IDSousFamille)values(" + Descrip + ","+
                Ref + "," + Prix + "," + GetIdByName("Marques", Marque) +","+ GetIdByName("SousFamilles", SousFamille));
            }

        }

        /// <summary>
        /// vérifier le donnée a déjà exist dans la base de données
        /// </summary>
        /// <param name="Table">La table</param>
        /// <param name="Colonne">Le colonne de la table</param>
        /// <param name="Value">La valeur du colonne</param>
        /// <returns>bool</returns>
        private bool DonneeIsExist(string Table, string Colonne, string Value)
        {
            CmdRecherche = ConnSql.CreateCommand();
            //Rechercher ID familles en relation dans table "Familles" 
            CmdRecherche.CommandText = "SELECT COUNT( " + Colonne + " ) FROM " + Table + " WHERE " + Colonne + "=@Value;";
            CmdRecherche.Parameters.AddWithValue("@Value", Value);
            CmdRecherche.ExecuteNonQuery();
            RdrReche = CmdRecherche.ExecuteReader();
            try
            {
                int Count = -1;
                if (RdrReche.Read())
                    Count = (int)RdrReche.GetValue(0);
                if (Count > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// obtenir ID par le nom donnée de la table
        /// </summary>
        /// <param name="Table">Le nom de la table</param>
        /// <param name="NomValue">La valeur du nom</param>
        /// <returns>ID </returns>
        private int GetIdByName(string Table, string NomValue)
        {
            CmdRecherche = ConnSql.CreateCommand();
            //Rechercher ID familles en relation dans table "Familles" 
            CmdRecherche.CommandText = "SELECT [ID] FROM "+ Table+" WHERE Nom=@Nom;";
            CmdRecherche.Parameters.AddWithValue("@Nom", NomValue);
            CmdRecherche.ExecuteNonQuery();
            RdrReche = CmdRecherche.ExecuteReader();
            try
            {
                int IDFam = -1;
                while (RdrReche.Read())//normalement il y a qu'un resultat
                {
                    IDFam = (int)RdrReche.GetValue(0);
                }
                //set le paramètre IDFam dans l'insertion de SousFamilles
                return IDFam;
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// supprimer toutes les données de la base de données .sdf
        /// </summary>
        private void ClearAllData()
        {
            ConnSql.Open();

            SqlCeCommand Cmd = new SqlCeCommand("DELETE FROM Articles", ConnSql);
            Cmd.ExecuteNonQuery();

            Cmd.CommandText = "DELETE FROM Marques";
            Cmd.ExecuteNonQuery();

            Cmd.CommandText = "DELETE FROM SousFamilles";
            Cmd.ExecuteNonQuery();

            Cmd.CommandText = "DELETE FROM Familles";
            Cmd.ExecuteNonQuery();

            ConnSql.Close();
        }
        
    }
}
