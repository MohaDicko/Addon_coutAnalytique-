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

        // Colonnes analytiques candidates sur les lignes de documents (par ordre de priorité)
        // 1. U_DocAnalytique (UDF ligne personnalisé)
        // 2. U_Analytique (au cas où l'UDF ligne porte le même nom que l'UDF article)
        // 3. 110000109 / 2004 (Colonnes standard SAP Règle de répartition / Centre de coûts Axe 1)
        private static readonly string[] CandidateCostCenterCols = new string[]
        {
            "U_DocAnalytique",
            "U_Analytique",
            "110000109",
            "2004"
        };

        // Configuration des formulaires gérés avec leurs paramètres (Codes FormTypeEx UI API + fallback ObjectType)
        private static readonly Dictionary<string, FormConfig> FormConfigurations = new Dictionary<string, FormConfig>
        {
            // Commandes clients (ORDR - FormTypeEx 139)
            { "139", new FormConfig { FormType = "139", FormName = "Commande Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "17",  new FormConfig { FormType = "17",  FormName = "Commande Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Commandes fournisseurs (OPOR - FormTypeEx 142)
            { "142", new FormConfig { FormType = "142", FormName = "Commande Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "22",  new FormConfig { FormType = "22",  FormName = "Commande Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Réception de marchandises / Entrée de marchandises OPOR (OPDN - FormTypeEx 143)
            { "143", new FormConfig { FormType = "143", FormName = "Réception de Marchandises", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "20",  new FormConfig { FormType = "20",  FormName = "Réception de Marchandises", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Livraison client (ODLN - FormTypeEx 140)
            { "140", new FormConfig { FormType = "140", FormName = "Livraison Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "15",  new FormConfig { FormType = "15",  FormName = "Livraison Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "16",  new FormConfig { FormType = "16",  FormName = "Retour / Livraison", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Facture client (OINV - FormTypeEx 133)
            { "133", new FormConfig { FormType = "133", FormName = "Facture Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "13",  new FormConfig { FormType = "13",  FormName = "Facture Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Facture fournisseur (OPCH - FormTypeEx 141)
            { "141", new FormConfig { FormType = "141", FormName = "Facture Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "18",  new FormConfig { FormType = "18",  FormName = "Facture Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },

            // Offres & Demandes d'achat / Devis / Avoirs (au cas où le consultant teste dessus)
            { "149", new FormConfig { FormType = "149", FormName = "Devis Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "179", new FormConfig { FormType = "179", FormName = "Avoir Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "180", new FormConfig { FormType = "180", FormName = "Retour Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "181", new FormConfig { FormType = "181", FormName = "Avoir Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "182", new FormConfig { FormType = "182", FormName = "Retour Marchandises Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "1470000200", new FormConfig { FormType = "1470000200", FormName = "Demande d'Achat", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "540000988",  new FormConfig { FormType = "540000988",  FormName = "Offre d'Achat", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } }
        };

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // 1. Récupération de la chaîne de connexion passée par SAP B1
                string connectionString = string.Empty;
                string[] cmdArgs = Environment.GetCommandLineArgs();

                if (cmdArgs.Length > 1 && !string.IsNullOrWhiteSpace(cmdArgs[1]))
                {
                    connectionString = cmdArgs[1];
                }
                else
                {
                    // Chaîne de connexion par défaut pour l'environnement de développement SAP B1 Studio
                    connectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                }

                // 2. Connexion à l'UI API
                SboGuiApi sboGuiApi = new SboGuiApi();
                try
                {
                    sboGuiApi.Connect(connectionString);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Impossible de se connecter au client SAP Business One.\n" +
                        "Veuillez vous assurer que SAP Business One est ouvert et en cours d'exécution.\n\nDétail : " + ex.Message,
                        "SAP_Analytix - Connexion impossible",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                sboApplication = sboGuiApi.GetApplication(-1);

                // 3. Connexion à la DI API (via Single Sign-On avec l'application UI)
                oCompany = (SAPbobsCOM.Company)sboApplication.Company.GetDICompany();

                if (oCompany == null || !oCompany.Connected)
                {
                    sboApplication.StatusBar.SetText(
                        "SAP_Analytix : Impossible d'établir la connexion DI-API.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }

                // 4. Abonnement aux événements UI et Cycle de vie de l'application
                sboApplication.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                sboApplication.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

                sboApplication.StatusBar.SetText(
                    "SAP_Analytix connecté avec succès (Gestion analytique active)",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);

                // Démarrage de la boucle de messages Windows
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
        /// Gestion des événements du cycle de vie de SAP Business One (fermeture, changement de société)
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
        /// Gestion des événements UI sur les formulaires
        /// </summary>
        private static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Ignorer les événements avant action (BeforeAction == true)
                if (pVal.BeforeAction)
                    return;

                // Vérifier si le type de formulaire est configuré
                if (string.IsNullOrEmpty(pVal.FormTypeEx) || !FormConfigurations.TryGetValue(pVal.FormTypeEx, out FormConfig config))
                    return;

                // Vérifier que l'événement concerne bien la colonne Article ("1") de la matrice des lignes ("38")
                if (pVal.ItemUID != config.MatrixUID || pVal.ColUID != config.ItemCodeCol || pVal.Row <= 0)
                    return;

                // On intercepte :
                // - et_CHOOSE_FROM_LIST : sélection d'un article via la loupe / touche Tab
                // - et_VALIDATE (avec ItemChanged) : saisie directe au clavier puis validation
                // - et_LOST_FOCUS : sortie de la cellule Article après sélection
                bool isTriggerEvent =
                    pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST ||
                    (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemChanged) ||
                    pVal.EventType == BoEventTypes.et_LOST_FOCUS;

                if (!isTriggerEvent)
                    return;

                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbouiCOM.EditText oEditItem = null;

                try
                {
                    oForm = sboApplication.Forms.Item(FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(config.MatrixUID).Specific;

                    if (pVal.Row > oMatrix.RowCount)
                        return;

                    string itemCode = string.Empty;

                    // Si c'est un ChooseFromList, on peut aussi lire le code article directement depuis le DataTable de l'événement
                    if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        IChooseFromListEvent oCFLEvento = (IChooseFromListEvent)pVal;
                        DataTable oDataTable = oCFLEvento.SelectedObjects;
                        if (oDataTable != null && oDataTable.Rows.Count > 0)
                        {
                            try
                            {
                                itemCode = Convert.ToString(oDataTable.GetValue("ItemCode", 0))?.Trim() ?? string.Empty;
                            }
                            catch
                            {
                                // Fallback sur la lecture de la cellule
                            }
                        }
                    }

                    // Lecture depuis la cellule de la matrice si non récupéré via CFL
                    if (string.IsNullOrEmpty(itemCode))
                    {
                        oEditItem = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(config.ItemCodeCol, pVal.Row);
                        itemCode = oEditItem?.Value != null ? oEditItem.Value.Trim() : string.Empty;
                    }

                    if (string.IsNullOrEmpty(itemCode))
                        return;

                    string costCenter = GetAnalyticalCodeFromItem(itemCode);

                    if (string.IsNullOrEmpty(costCenter))
                    {
                        // Avertir uniquement sur Validate/CFL pour ne pas spammer sur LostFocus
                        if (pVal.EventType != BoEventTypes.et_LOST_FOCUS)
                        {
                            sboApplication.StatusBar.SetText(
                                $"SAP_Analytix : Aucun code analytique (U_Analytique) défini sur l'article '{itemCode}'.",
                                BoMessageTime.bmt_Short,
                                BoStatusBarMessageType.smt_Warning);
                        }
                        return;
                    }

                    // Affectation dans la colonne analytique de la ligne
                    ApplyCostCenterToRow(oForm, oMatrix, config, pVal.Row, costCenter, pVal.EventType);
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
            catch (Exception ex)
            {
                sboApplication?.StatusBar?.SetText("Erreur Add-on : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Tente de renseigner le code analytique dans la colonne de la matrice
        /// </summary>
        private static void ApplyCostCenterToRow(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, FormConfig config, int row, string costCenter, BoEventTypes eventType)
        {
            SAPbouiCOM.EditText oEditCostCenter = null;
            try
            {
                // Chercher la colonne cible parmi les colonnes candidates (U_DocAnalytique en priorité, puis standard SAP)
                string targetColUID = FindTargetColumnUID(oMatrix, config.CostCenterCol);

                if (string.IsNullOrEmpty(targetColUID))
                {
                    if (eventType != BoEventTypes.et_LOST_FOCUS)
                    {
                        sboApplication.StatusBar.SetText(
                            $"SAP_Analytix : Colonne '{config.CostCenterCol}' introuvable ou masquée dans [{config.FormName}]. Activez-la dans Paramètres de formulaire.",
                            BoMessageTime.bmt_Short,
                            BoStatusBarMessageType.smt_Error);
                    }
                    return;
                }

                oEditCostCenter = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(targetColUID, row);
                string currentValue = oEditCostCenter?.Value != null ? oEditCostCenter.Value.Trim() : string.Empty;

                // Si la valeur est déjà renseignée à l'identique, ne pas réécrire inutilement
                if (string.Equals(currentValue, costCenter, StringComparison.OrdinalIgnoreCase))
                    return;

                oForm.Freeze(true);
                try
                {
                    oEditCostCenter.Value = costCenter;

                    sboApplication.StatusBar.SetText(
                        $"[{config.FormName}] Code analytique '{costCenter}' appliqué à la ligne {row} (colonne {targetColUID})",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                if (eventType != BoEventTypes.et_LOST_FOCUS)
                {
                    sboApplication.StatusBar.SetText(
                        $"SAP_Analytix : Impossible d'écrire dans la colonne analytique (vérifiez qu'elle est visible et modifiable dans Paramètres de formulaire). Détail : {ex.Message}",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (oEditCostCenter != null)
                {
                    Marshal.ReleaseComObject(oEditCostCenter);
                    oEditCostCenter = null;
                }
            }
        }

        /// <summary>
        /// Recherche la colonne analytique présente et visible dans la matrice
        /// </summary>
        private static string FindTargetColumnUID(SAPbouiCOM.Matrix oMatrix, string preferredCol)
        {
            SAPbouiCOM.Columns oCols = null;
            try
            {
                oCols = oMatrix.Columns;

                // 1. Tester la colonne préférée (ex: U_DocAnalytique)
                if (ColumnExistsAndEditable(oCols, preferredCol))
                    return preferredCol;

                // 2. Tester les autres colonnes candidates
                foreach (string candidate in CandidateCostCenterCols)
                {
                    if (ColumnExistsAndEditable(oCols, candidate))
                        return candidate;
                }

                return string.Empty;
            }
            finally
            {
                if (oCols != null)
                {
                    Marshal.ReleaseComObject(oCols);
                    oCols = null;
                }
            }
        }

        private static bool ColumnExistsAndEditable(SAPbouiCOM.Columns oCols, string colUID)
        {
            SAPbouiCOM.Column oCol = null;
            try
            {
                oCol = oCols.Item(colUID);
                return oCol != null && oCol.Visible && oCol.Editable;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oCol != null)
                {
                    Marshal.ReleaseComObject(oCol);
                    oCol = null;
                }
            }
        }

        /// <summary>
        /// Récupère le code analytique (U_Analytique) depuis la fiche article (OITM)
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
                // Sélectionne toutes les colonnes de l'article pour éviter une erreur SQL si U_Analytique porte un autre nom
                string query = $"SELECT * FROM \"OITM\" WHERE \"ItemCode\" = '{safeItemCode}'";

                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    fields = oRec.Fields;

                    // Essayer U_Analytique en priorité, puis U_DocAnalytique ou OcrCode si présent
                    string[] itemCandidateFields = new string[] { "U_Analytique", "U_DocAnalytique", "U_CentreCout", "CardCode" };
                    foreach (string fieldName in itemCandidateFields)
                    {
                        if (fieldName == "CardCode")
                            break; // Ne pas utiliser CardCode, c'était juste un marqueur de fin

                        try
                        {
                            field = fields.Item(fieldName);
                            object val = field?.Value;
                            if (val != null && !(val is DBNull))
                            {
                                string strVal = val.ToString().Trim();
                                if (!string.IsNullOrEmpty(strVal))
                                    return strVal;
                            }
                        }
                        catch
                        {
                            // Ce champ n'existe pas sur OITM, on passe au suivant
                        }
                        finally
                        {
                            if (field != null)
                            {
                                Marshal.ReleaseComObject(field);
                                field = null;
                            }
                        }
                    }
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                sboApplication?.StatusBar?.SetText("Erreur lecture fiche article : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// Nettoyage propre des objets COM à l'arrêt du programme
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
    /// Configuration pour chaque type de formulaire SAP géré
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
