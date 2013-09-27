using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

using Oracle;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;

using Ionic.Utils.Zip;

using GZM.Framework.Batch.Utilites;

namespace DV1L3001_ArchivageFichiersRdi
{
    public class ArchivageRDI
    {
        #region Private Members

        private string _emailFrom;
        private string _emailErrorTo;
        private int _codeRetourErreur;
        private string _backupFolder;
        private string _sourceFolder;
        private int _nbMoisAConserver;
        private long _noEmissionBase;
        private long _maxNbMissingFolder;

        #endregion

        #region Properties

        public string EmailFrom
        {
            get { return Properties.Settings.Default.EmailFrom; }
        }

        public string EmailErrorTo
        {
            get { return Properties.Settings.Default.EmailErrorTo; }
        }

        public int CodeRetourErreur
        {
            get { return Properties.Settings.Default.CODE_RETOUR_ERREUR; }
        }

        public string BackupFolder
        {
            get { return Properties.Settings.Default.BACKUP_DESTINATION_FOLDER; }
        }

        public string SourceFolder
        {
            get { return Properties.Settings.Default.SOURCE_FOLDER; }
        }

        public int NbMoisAConserver
        {
            get { return Properties.Settings.Default.NB_MOIS_A_CONSERVER; }
        }

        public long NoEmissionBase
        {
            get { return Properties.Settings.Default.NO_EMISSION_BASE; }
        }

        public long MaxNbMissingFolder
        {
            get { return Properties.Settings.Default.MAX_NB_MISSING_FOLDER_FOR_EXIT; }
        }

        

        #endregion

        #region Methods
        public ArchivageRDI()
        {
            //=== Init des variables dans le constructeur
            _emailFrom = EmailFrom;
            _emailErrorTo = EmailErrorTo;
            _codeRetourErreur = CodeRetourErreur;
            _backupFolder = BackupFolder;
            _sourceFolder = SourceFolder;
            _nbMoisAConserver = NbMoisAConserver;
            _noEmissionBase = NoEmissionBase;
            _maxNbMissingFolder = MaxNbMissingFolder;
        }

        public int Archivage()
        {
            int intCodeRetour = 0x00; //=== Code de retour Normalement Terminer

            //==== Periode selon la date du jour par défaut
            string strSQL = "SELECT cp.cal_no_emission FROM lg2_calen_prod cp WHERE cp.cal_dt_emission <= ADD_MONTHS(SYSDATE, :Mois_A_Conserver) ORDER BY cp.cal_dt_emission DESC";

            OracleConnection objConn = new OracleConnection(Properties.Settings.Default.CONNECTSTRING); ;
            OracleCommand objCmd = new OracleCommand(strSQL, objConn);
            OracleDataReader odr = null;

            // create a writer and open the file
            TextWriter log = new StreamWriter(".\\logs\\DV1L3001_" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log");

            try
            {

                // write a line of text to the file
                log.WriteLine(DateTime.Now.ToString() + "==== DEBUT DU TRAITEMENT");
                log.WriteLine(DateTime.Now + " _sourceFolder = " + _sourceFolder);
                log.WriteLine(DateTime.Now + " _backupFolder = " + _backupFolder);
                log.WriteLine(DateTime.Now + " _nbMoisAConserver = " + _nbMoisAConserver);


                #region Validation des informations
                string msgErreur = "";
                //=== On vérifie si répertoire de destination existe
                if (!Directory.Exists(_sourceFolder))
                {
                    msgErreur += "Le répertoire source est inexistant. [source_folder = " + _sourceFolder + "]<BR>";
                }
                if (!Directory.Exists(_backupFolder))
                {
                    msgErreur += "Le répertoire de sauvegarde est inexistant à l'endroit spécifié. [backup_destination_folder = " + _backupFolder + "]<BR>";
                }

                if (msgErreur != "")
                {
                    log.WriteLine(DateTime.Now + " ERROR : " + msgErreur.ToString());
                    throw new ApplicationException(msgErreur);
                }
                #endregion

                objConn.Open();
                //==== On soustrait les mois pour trouver la date de départ
                objCmd.Parameters.Add("Mois_A_Conserver", OracleDbType.Varchar2).Value = _nbMoisAConserver * -1;

                odr = objCmd.ExecuteReader(CommandBehavior.CloseConnection);

                long noEmission = 0;
                if (odr.Read())
                {
                    noEmission = _noEmissionBase + (int)odr["cal_no_emission"];
                    log.WriteLine(DateTime.Now + " Start with noEmission = " + noEmission.ToString());
                }
                else
                {
                    log.WriteLine(DateTime.Now + " Start with noEmission = AUCUNE ÉMISSION");
                }


                int nbMissingFolders = 0;
                for (long i = noEmission; i >= _noEmissionBase; i--)
                {
                    //=== Si le folder existe on le zip
                    string archiveFolder = _sourceFolder + "\\" + i.ToString();
                    
                    if (Directory.Exists(archiveFolder))
                    {
                        nbMissingFolders = 0;                        
                        log.WriteLine(DateTime.Now + " FOLDER EXIST archiveFolder = " + archiveFolder.ToString());
                        using (ZipFile zip = new ZipFile())
                        {
                            zip.UpdateDirectory(archiveFolder);//==== Répertoire qui sera compressé
                            zip.Save(_backupFolder + "\\" + i.ToString() + ".zip");  //=== Nom et emplacement du fichier ZIP
                            log.WriteLine(DateTime.Now + " ZIP TO : " + _backupFolder + "\\" + i.ToString() + ".zip");
                            zip.Dispose();
                        }
                        Directory.Delete(archiveFolder, true); //==== Supprime le répertoire qui a été compressé
                        log.WriteLine(DateTime.Now + " DELETE UNCOMPRESSED FOLDER : " + archiveFolder.ToString());
                    }
                    else
                    {
                        nbMissingFolders++; //=== Cumule le nombre de répertoire manquan
                        log.WriteLine(DateTime.Now + " _" + nbMissingFolders + "_ FOLDER DOESNT EXIST archiveFolder = " + archiveFolder.ToString());

                        if (nbMissingFolders >= _maxNbMissingFolder)//=== Si 5 répertoire consécutif n'existe pas on arrete le traitement.
                        {
                            log.WriteLine(DateTime.Now + " STOP CAUSE BY MORE THEN " + _maxNbMissingFolder.ToString() + " MISSING FOLDERS");
                            break;
                        }

                    }
                }

                log.WriteLine(DateTime.Now.ToString() + "==== FIN DU TRAITEMENT");

            }
            catch (Exception ex)
            {
                //=== Envoie de courriel sur erreur
                intCodeRetour = _codeRetourErreur;
                log.WriteLine(DateTime.Now.ToString() + " ERROR : " + ex.Message);
                GZMMailer mail = new GZMMailer(_emailFrom, _emailErrorTo, "Erreur dans l'application : DV1L3001_ArchivageFichiersRdi", ex.Message, "", 1, "", "");
            }
            finally
            {
                // close the stream
                if (log != null)
                {
                    log.Close();
                    log.Dispose();
                }

                //=== Ferme les objets
                if (odr != null)
                {
                    odr.Close();
                    odr.Dispose();
                }

                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
            }
            return intCodeRetour;

        }
        #endregion
    }
}
