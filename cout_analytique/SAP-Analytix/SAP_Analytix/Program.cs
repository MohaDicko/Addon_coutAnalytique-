using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SAP_Analytix
{
    static class Program
    {
        public static SAPbouiCOM.Application sboApplication;
        public static SAPbobsCOM.Company oCompany;

        // Configuration des formulaires gÃ©rÃ©s avec leurs paramÃ¨tres
        private static readonly Dictionary<string, FormConfig> FormConfigurations = new Dictionary<string, FormConfig>
        {
            // Commandes clients
            { "139", new FormConfig { FormType = "139", FormName = "Commande Client (Sales Order)", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Commandes fournisseurs
            { "22", new FormConfig { FormType = "22", FormName = "Commande Fournisseur (Purchase Order)", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // RÃ©ception de marchandises
            { "20", new FormConfig { FormType = "20", FormName = "RÃ©ception de Marchandises (Goods Receipt)", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Livraison client
            { "16", new FormConfig { FormType = "16", FormName = "Livraison (Delivery)", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Facture client
            { "13", new FormConfig { FormType = "13", FormName = "Facture Client (Sales Invoice)", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Facture fournisseur
            { "18", new FormConfig { FormType = "18", FormName = "Facture Fournisseur (Purchase Invoice)", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } }
        };

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // 1. RÃ©cupÃ©ration de la chaÃ®ne de connexion passÃ©e par SAP B1
                string connectionString = string.Empty;
                string[] cmdArgs = Environment.GetCommandLineArgs();

                if (cmdArgs.Length > 1)
                {
                    connectionString = cmdArgs[1];
                }
                else
                {
                    // ChaÃ®ne de connexion par dÃ©faut pour l'environnement de dÃ©veloppement SAP B1 Studio
                    connectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                }

                // 2. Connexion Ã  l'UI API
                SboGuiApi sboGuiApi = new SboGuiApi();
                try
                {
                    sboGuiApi.Connect(connectionString);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Impossible de se connecter au client SAP Business One.\n" +
                        "Veuillez vous assurer que SAP Business One est ouvert et en cours d'exÃ©cution.\n\nDÃ©tail : " + ex.Message,
                        "SAP_Analytix - Connexion impossible",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                sboApplication = sboGuiApi.GetApplication(-1);

                // 3. Connexion Ã  la DI API (via Single Sign-On avec l'application UI)
                oCompany = (SAPbobsCOM.Company)sboApplication.Company.GetDICompany();

                if (oCompany == null || !oCompany.Connected)
                {
                    sboApplication.StatusBar.SetText(
                        "SAP_Analytix : Impossible d'Ã©tablir la connexion DI-API.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }

                // 4. Abonnement aux Ã©vÃ©nements UI et Cycle de vie de l'application
                sboApplication.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                sboApplication.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

                sboApplication.StatusBar.SetText(
                    "SAP_Analytix connectÃ© avec succÃ¨s (Gestion analytique active)",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);

                // DÃ©marrage de la boucle de messages Windows
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur critique SAP_Analytix : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CleanupResources();
            }
        }

        /// <summary>
        /// Gestion des Ã©vÃ©nements du cycle de vie de SAP Business One (fermeture, changement de sociÃ©tÃ©)
        /// </summary>
        private static void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                case BoAppEventTypes.aet_CompanyChanged:
                case BoAppEventTypes.aet_LanguageChanged:
                case BoAppEventTypes.aet_ServerTerminition:
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        /// <summary>
        /// Gestion des Ã©vÃ©nements UI sur les formulaires
        /// </summary>
        private static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // VÃ©rifier si le type de formulaire est configurÃ©
                if (string.IsNullOrEmpty(pVal.FormTypeEx) || !FormConfigurations.TryGetValue(pVal.FormTypeEx, out FormConfig config))
                    return;

                // On intercepte la validation d'une cellule de la colonne Article dans la matrice aprÃ¨s saisie
                if (pVal.EventType == BoEventTypes.et_VALIDATE
                    && !pVal.BeforeAction
                    && pVal.ItemChanged
                    && pVal.ItemUID == config.MatrixUID
                    && pVal.ColUID == config.ItemCodeCol
                    && pVal.Row > 0)
                {
                    SAPbouiCOM.Form oForm = null;
                    SAPbouiCOM.Matrix oMatrix = null;
                    SAPbouiCOM.EditText oEditItem = null;
                    SAPbouiCOM.EditText oEditCostCenter = null;

                    try
                    {
                        oForm = sboApplication.Forms.Item(FormUID);
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(config.MatrixUID).Specific;

                        if (pVal.Row > oMatrix.RowCount)
                            return;

                        oEditItem = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(config.ItemCodeCol, pVal.Row);
                        string itemCode = oEditItem?.Value != null ? oEditItem.Value.Trim() : string.Empty;

                        if (!string.IsNullOrEmpty(itemCode))
                        {
                            string costCenter = GetAnalyticalCodeFromItem(itemCode);

                            if (!string.IsNullOrEmpty(costCenter))
                            {
                                oForm.Freeze(true);
                                try
                                {
                                    oEditCostCenter = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(config.CostCenterCol, pVal.Row);
                                    oEditCostCenter.Value = costCenter;

                                    sboApplication.StatusBar.SetText(
                                        $"[{config.FormName}] Centre de coÃ»ts '{costCenter}' appliquÃ© Ã  la ligne {pVal.Row}",
                                        BoMessageTime.bmt_Short,
                                        BoStatusBarMessageType.smt_Success);
                                }
                                finally
                                {
                                    oForm.Freeze(false);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sboApplication.StatusBar.SetText(
                            $"SAP_Analytix - Erreur ligne {pVal.Row} : {ex.Message}",
                            BoMessageTime.bmt_Short,
                            BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        // LibÃ©ration systÃ©matique des objets COM UI
                        if (oEditCostCenter != null)
                        {
                            Marshal.ReleaseComObject(oEditCostCenter);
                            oEditCostCenter = null;
                        }
                        if (oEditItem != null)
                        {
                            Marshal.ReleaseComObject(oEditItem);
                            oEditItem = null;
                        }
                        if (oMatrix != null)
                        {
                            Marshal.ReleaseComObject(oMatrix);
                            oMatrix = null;
                        }
                        if (oForm != null)
                        {
                            Marshal.ReleaseComObject(oForm);
                            oForm = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                sboApplication?.StatusBar?.SetText("Erreur Add-on : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// RÃ©cupÃ¨re le code analytique (U_Analytique) depuis la fiche article (OITM)
        /// </summary>
        private static string GetAnalyticalCodeFromItem(string itemCode)
        {
            if (oCompany == null || !oCompany.Connected)
                return string.Empty;

            Recordset oRec = null;
            SAPbobsCOM.Fields fields = null;
            SAPbobsCOM.Field field = null;

            try
            {
                oRec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string safeItemCode = itemCode.Replace("'", "''");
                string query = $"SELECT \"U_Analytique\" FROM \"OITM\" WHERE \"ItemCode\" = '{safeItemCode}'";

                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    fields = oRec.Fields;
                    field = fields.Item("U_Analytique");

                    object val = field.Value;
                    if (val != null && !(val is DBNull))
                    {
                        return val.ToString().Trim();
                    }
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                sboApplication?.StatusBar?.SetText("Erreur lecture code article : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return string.Empty;
            }
            finally
            {
                if (field != null)
                {
                    Marshal.ReleaseComObject(field);
                    field = null;
                }
                if (fields != null)
                {
                    Marshal.ReleaseComObject(fields);
                    fields = null;
                }
                if (oRec != null)
                {
                    Marshal.ReleaseComObject(oRec);
                    oRec = null;
                }
            }
        }

        /// <summary>
        /// Nettoyage propre des objets COM Ã  l'arrÃªt du programme
        /// </summary>
        private static void CleanupResources()
        {
            try
            {
                if (oCompany != null)
                {
                    if (oCompany.Connected)
                    {
                        oCompany.Disconnect();
                    }
                    Marshal.ReleaseComObject(oCompany);
                    oCompany = null;
                }

                if (sboApplication != null)
                {
                    Marshal.ReleaseComObject(sboApplication);
                    sboApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
                // Ignorer les erreurs lors du nettoyage de fermeture
            }
        }
    }

    /// <summary>
    /// Configuration pour chaque type de formulaire SAP gÃ©rÃ©
    /// </summary>
    public class FormConfig
    {
        public string FormType { get; set; }
        public string FormName { get; set; }
        public string MatrixUID { get; set; }
        public string ItemCodeCol { get; set; }
        public string CostCenterCol { get; set; }
    }
}
