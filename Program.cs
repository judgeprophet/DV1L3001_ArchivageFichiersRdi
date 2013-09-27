using System;
using System.Runtime.InteropServices;

namespace DV1L3001_ArchivageFichiersRdi
{
    /// AUTEUR: Sébastien Vincent
    /// DATE: 2008-03-28
    /// BUT: Déplacer les factures plus vieille que 36 mois dans un autre dossier sur un autre SAN. 
    ///      Dans le but d'optimiser l'espace disque (windows) et d'accélérer la vitesse de restore en cas de problème ou de relève.
    /// PARAMETRE:  (Optionnel) Date de la période désiré FORMAT: YYYYMMJJ.
    /// Si aucun paramètre n'est saisie la date du jour est utilisé
    class Program
    {
        #region Private Member
        [DllImport("kernel32.dll")]
        private static extern void ExitProcess(int a);
        #endregion

        static void Main(string[] args)
        {
            ArchivageRDI ardi = new ArchivageRDI();
            ExitProcess(ardi.Archivage());
        }
    }
}
