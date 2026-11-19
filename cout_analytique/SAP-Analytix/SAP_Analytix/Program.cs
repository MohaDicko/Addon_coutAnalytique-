using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SAP_Analytix
{
    /// <summary>
    /// Contient toutes les données analytiques récupérées depuis la fiche article (OITM) :
    /// - Valeur(s) UDF personnalisée(s) (code brut + code/libellé résolu via OOCR/OPRC)
    /// - Valeurs pour chacun des 5 axes analytiques SAP B1 (Dimensions 1 à 5 : OcrCode à OcrCode5)
    /// </summary>
    public class ItemAnalyticalData
    {
        public string CustomUdfValue { get; set; } = string.Empty;
        public string ResolvedOcrCode { get; set; } = string.Empty;
        public string ResolvedOcrName { get; set; } = string.Empty;

        // Tous les champs U_* analytiques trouvés sur OITM (Nom du champ -> Valeur)
        public Dictionary<string, string> AllUdfValues { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Alias (OcrCode, OcrName) résolus pour chaque valeur brute rencontrée
        public Dictionary<string, string[]> ValueAliases { get; } = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

        // Codes par axe analytique SAP (1 = OcrCode, 2 = OcrCode2, 3 = OcrCode3, 4 = OcrCode4, 5 = OcrCode5)
        public Dictionary<int, string> DimensionValues { get; } = new Dictionary<int, string>();

        public bool HasAnyValue =>
            !string.IsNullOrEmpty(CustomUdfValue) ||
            DimensionValues.Count > 0 ||
            AllUdfValues.Count > 0;
    }

    internal class CostCenterCacheEntry
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public int DimCode { get; set; }
    }

    internal class MatrixColumnMap
    {
        public Dictionary<int, string> DimensionCols { get; } = new Dictionary<int, string>();
        public string PrimaryUdfCol { get; set; } = string.Empty;
        public HashSet<string> EditableUdfCols { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    internal class CachedItemEntry
    {
        public DateTime LoadedAtUtc { get; set; }
        public ItemAnalyticalData Data { get; set; }
    }

    public static class Program
    {
        public static SAPbouiCOM.Application sboApplication;
        public static SAPbobsCOM.Company oCompany;

        // Mémorise le dernier article sélectionné via ChooseFromList par ligne
        private static string _lastSelectedItemCode = string.Empty;
        private static int _lastSelectedRow = -1;

        // Caches en mémoire pour éviter les freezes COM et les requêtes SQL répétitives
        private static bool _oocrCacheLoaded = false;
        private static readonly Dictionary<string, CostCenterCacheEntry> _oocrCache = new Dictionary<string, CostCenterCacheEntry>(StringComparer.OrdinalIgnoreCase);
        private static List<string> _oitmAnalyticalColumns = null;
        private static readonly Dictionary<string, MatrixColumnMap> _matrixColumnCache = new Dictionary<string, MatrixColumnMap>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, CachedItemEntry> _itemCache = new Dictionary<string, CachedItemEntry>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> _lastAppliedByFormRow = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly TimeSpan ItemCacheTtl = TimeSpan.FromSeconds(60);

        /// <summary>
        /// Réinitialise tous les caches en mémoire (utile lors d'un changement de société ou pour les tests unitaires).
        /// </summary>
        public static void ResetCaches()
        {
            _oocrCacheLoaded = false;
            _oocrCache.Clear();
            _oitmAnalyticalColumns = null;
            _matrixColumnCache.Clear();
            _itemCache.Clear();
            _lastAppliedByFormRow.Clear();
            _lastSelectedItemCode = string.Empty;
            _lastSelectedRow = -1;
        }

        // Colonnes UDF personnalisées candidates sur les lignes de documents
        public static readonly string[] CandidateCustomUdfCols = new string[]
        {
            "U_DocAnalytique",
            "U_Analytique",
            "U_CodeAnalytique",
            "U_CentreCout"
        };

        // Correspondance des 5 axes analytiques SAP Business One dans la matrice des lignes (UIDs UI API & Alias DB)
        public static readonly Dictionary<int, string[]> DimensionColumnCandidates = new Dictionary<int, string[]>
        {
            // Axe 1 (OcrCode)
            { 1, new[] { "2004", "110000109", "OcrCode" } },
            // Axe 2 (OcrCode2)
            { 2, new[] { "2005", "110000110", "10001066", "OcrCode2" } },
            // Axe 3 (OcrCode3 - ex: Direction commerciale)
            { 3, new[] { "2006", "110000111", "10001068", "OcrCode3" } },
            // Axe 4 (OcrCode4 - ex: Maintenance)
            { 4, new[] { "2007", "110000112", "10001070", "OcrCode4" } },
            // Axe 5 (OcrCode5 - ex: Support)
            { 5, new[] { "2008", "110000113", "10001072", "OcrCode5" } }
        };

        // Liste globale pour compatibilité ascendante
        public static readonly string[] CandidateCostCenterCols = new string[]
        {
            "U_DocAnalytique",
            "U_Analytique",
            "U_CodeAnalytique",
            "U_CentreCout",
            "110000109", "2004",
            "110000110", "2005", "10001066",
            "110000111", "2006", "10001068",
            "110000112", "2007", "10001070",
            "110000113", "2008", "10001072"
        };

        // Configuration des formulaires gérés avec leurs paramètres (Codes FormTypeEx UI API + fallback ObjectType)
        public static readonly Dictionary<string, FormConfig> FormConfigurations = new Dictionary<string, FormConfig>
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

            // Offres & Demandes d'achat / Devis / Avoirs / Sorties & Entrées de stock
            { "149", new FormConfig { FormType = "149", FormName = "Devis Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "179", new FormConfig { FormType = "179", FormName = "Avoir Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "180", new FormConfig { FormType = "180", FormName = "Retour Client", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "181", new FormConfig { FormType = "181", FormName = "Avoir Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "182", new FormConfig { FormType = "182", FormName = "Retour Marchandises Fournisseur", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "720", new FormConfig { FormType = "720", FormName = "Sortie de Marchandises", MatrixUID = "13", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "721", new FormConfig { FormType = "721", FormName = "Entrée de Marchandises", MatrixUID = "13", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "1470000200", new FormConfig { FormType = "1470000200", FormName = "Demande d'Achat", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } },
            { "540000988",  new FormConfig { FormType = "540000988",  FormName = "Offre d'Achat", MatrixUID = "38", ItemCodeCol = "1", CostCenterCol = "U_DocAnalytique" } }
        };

        private static string LogFilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SAP_Analytix_debug.log");

        public static void Log(string message)
        {
            try
            {
                File.AppendAllText(LogFilePath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignorer les erreurs d'écriture du journal
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // Mode d'auto-test hors-ligne pour valider toute la logique C# sans SAP ouvert
                if (args != null && args.Length > 0 && args[0] == "--selftest")
                {
                    Environment.ExitCode = SelfTestRunner.RunAllTests();
                    return;
                }

                Log("Démarrage de SAP_Analytix (Mode optimisé haute performance, 5 axes)...");

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
                    Log("Échec connexion UI API : " + ex.Message);
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
                    Log("Échec connexion DI API.");
                    sboApplication.StatusBar.SetText(
                        "SAP_Analytix : Impossible d'établir la connexion DI-API.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }

                Log($"Connecté à la société : {oCompany.CompanyName} ({oCompany.CompanyDB})");

                // 4. Préchargement du cache OOCR/OPRC et découverte unique des colonnes OITM
                EnsureCostCenterCacheLoaded();

                // 5. Filtrage strict des événements UI (évite 99% du trafic COM inutile dans SAP)
                SetApplicationFilters();

                // 6. Abonnement aux événements UI et Cycle de vie de l'application
                sboApplication.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                sboApplication.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

                sboApplication.StatusBar.SetText(
                    "SAP_Analytix connecté avec succès (Gestion analytique 5 axes active)",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);

                // Démarrage de la boucle de messages Windows
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                Log("Erreur critique Main : " + ex);
                MessageBox.Show("Erreur critique SAP_Analytix : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CleanupResources();
            }
        }

        /// <summary>
        /// Filtre les événements UI API au niveau de SAP Business One afin que seuls
        /// CHOOSE_FROM_LIST, VALIDATE, COMBO_SELECT et FORM_CLOSE sur les formulaires gérés
        /// soient transmis à l'add-on (supprime les ralentissements sur tout le reste de SAP).
        /// </summary>
        private static void SetApplicationFilters()
        {
            try
            {
                EventFilters oFilters = new EventFilters();
                EventFilter oFilterCfl = oFilters.Add(BoEventTypes.et_CHOOSE_FROM_LIST);
                EventFilter oFilterVal = oFilters.Add(BoEventTypes.et_VALIDATE);
                EventFilter oFilterCombo = oFilters.Add(BoEventTypes.et_COMBO_SELECT);
                EventFilter oFilterClose = oFilters.Add(BoEventTypes.et_FORM_CLOSE);

                foreach (string formType in FormConfigurations.Keys)
                {
                    oFilterCfl.AddEx(formType);
                    oFilterVal.AddEx(formType);
                    oFilterCombo.AddEx(formType);
                    oFilterClose.AddEx(formType);
                }

                sboApplication.SetFilter(oFilters);
            }
            catch (Exception ex)
            {
                Log("Avertissement SetApplicationFilters : " + ex.Message);
            }
        }

        /// <summary>
        /// Gestion des événements du cycle de vie de SAP Business One (fermeture, changement de société)
        /// </summary>
        public static void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                case BoAppEventTypes.aet_CompanyChanged:
                case BoAppEventTypes.aet_LanguageChanged:
                case BoAppEventTypes.aet_ServerTerminition:
                    ResetCaches();
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        /// <summary>
        /// Gestion des événements UI sur les formulaires
        /// </summary>
        public static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal == null || pVal.BeforeAction)
                    return;

                // Nettoyer les entrées de déduplication lorsqu'un formulaire se ferme
                if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
                {
                    ClearFormRowTracking(FormUID);
                    return;
                }

                if (string.IsNullOrEmpty(pVal.FormTypeEx) || !FormConfigurations.TryGetValue(pVal.FormTypeEx, out FormConfig config))
                    return;

                if (pVal.ItemUID != config.MatrixUID || pVal.Row <= 0)
                    return;

                bool isItemColumn = (pVal.ColUID == config.ItemCodeCol || pVal.ColUID == "3");
                if (!isItemColumn)
                    return;

                bool isTriggerEvent =
                    pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST ||
                    (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemChanged) ||
                    pVal.EventType == BoEventTypes.et_COMBO_SELECT;

                if (!isTriggerEvent)
                    return;

                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbouiCOM.EditText oEditItem = null;

                try
                {
                    string itemCode = string.Empty;

                    // Si c'est un ChooseFromList, récupérer l'ItemCode depuis SelectedObjects sans toucher la matrice si possible
                    if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        try
                        {
                            IChooseFromListEvent oCFLEvento = (IChooseFromListEvent)pVal;
                            DataTable oDataTable = oCFLEvento.SelectedObjects;
                            if (oDataTable != null && oDataTable.Rows.Count > 0)
                            {
                                itemCode = Convert.ToString(oDataTable.GetValue("ItemCode", 0))?.Trim() ?? string.Empty;
                                if (!string.IsNullOrEmpty(itemCode))
                                {
                                    _lastSelectedItemCode = itemCode;
                                    _lastSelectedRow = pVal.Row;
                                }
                            }
                        }
                        catch
                        {
                            // Fallback sur la lecture de la cellule
                        }
                    }

                    oForm = sboApplication.Forms.Item(FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(config.MatrixUID).Specific;

                    if (pVal.Row > oMatrix.RowCount)
                        return;

                    if (string.IsNullOrEmpty(itemCode))
                    {
                        try
                        {
                            oEditItem = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific(config.ItemCodeCol, pVal.Row);
                            itemCode = oEditItem?.Value != null ? oEditItem.Value.Trim() : string.Empty;
                        }
                        catch
                        {
                            // Ignorer
                        }
                    }

                    if (string.IsNullOrEmpty(itemCode) && _lastSelectedRow == pVal.Row && !string.IsNullOrEmpty(_lastSelectedItemCode))
                    {
                        itemCode = _lastSelectedItemCode;
                    }

                    if (string.IsNullOrEmpty(itemCode))
                        return;

                    // Déduplication : évite d'exécuter 2 fois le traitement sur la même ligne quand et_VALIDATE suit immédiatement et_CHOOSE_FROM_LIST
                    string formRowKey = $"{FormUID}|{pVal.Row}";
                    if (_lastAppliedByFormRow.TryGetValue(formRowKey, out string alreadyAppliedItem) &&
                        string.Equals(alreadyAppliedItem, itemCode, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    ItemAnalyticalData analyticalData = GetAnalyticalDataFromItem(itemCode);

                    if (analyticalData == null || !analyticalData.HasAnyValue)
                    {
                        string msg = $"SAP_Analytix : Aucun code analytique défini sur l'article '{itemCode}'.";
                        sboApplication.StatusBar.SetText(msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        _lastAppliedByFormRow[formRowKey] = itemCode;
                        return;
                    }

                    // Affectation sur tous les axes analytiques concernés (1 à 5) et/ou les colonnes UDF personnalisées
                    bool applied = ApplyAnalyticalDataToRow(oForm, oMatrix, config, pVal.Row, analyticalData, pVal.EventType);
                    if (applied)
                    {
                        _lastAppliedByFormRow[formRowKey] = itemCode;
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur SBO_Application_ItemEvent (Form={pVal.FormTypeEx}, Row={pVal.Row}) : {ex}");
                    sboApplication.StatusBar.SetText(
                        $"SAP_Analytix - Erreur ligne {pVal.Row} : {ex.Message}",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    ReleaseCom(ref oEditItem);
                    ReleaseCom(ref oMatrix);
                    ReleaseCom(ref oForm);
                }
            }
            catch (Exception ex)
            {
                Log("Erreur globale ItemEvent : " + ex.Message);
                sboApplication?.StatusBar?.SetText("Erreur Add-on : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private static void ClearFormRowTracking(string formUID)
        {
            if (string.IsNullOrEmpty(formUID) || _lastAppliedByFormRow.Count == 0)
                return;

            string prefix = formUID + "|";
            var toRemove = new List<string>();
            foreach (string key in _lastAppliedByFormRow.Keys)
            {
                if (key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    toRemove.Add(key);
            }
            foreach (string key in toRemove)
            {
                _lastAppliedByFormRow.Remove(key);
            }
        }

        /// <summary>
        /// Applique les codes analytiques de l'article sur les colonnes correspondantes (Axes 1 à 5 + UDF) de la ligne.
        /// Utilise un cache de colonnes par FormTypeEx et Freeze(true/false) pour éviter tout ralentissement d'affichage.
        /// </summary>
        public static bool ApplyAnalyticalDataToRow(
            SAPbouiCOM.Form oForm,
            SAPbouiCOM.Matrix oMatrix,
            FormConfig config,
            int row,
            ItemAnalyticalData data,
            BoEventTypes eventType)
        {
            List<string> appliedSummary = new List<string>();
            bool anyColumnFound = false;

            // Récupérer ou construire une seule fois la carte des colonnes analytiques pour ce type de formulaire
            MatrixColumnMap colMap = GetOrBuildMatrixColumnMap(oMatrix, config);

            List<string> BuildCandidates(string primaryVal, bool includeGlobalResolved)
            {
                var list = new List<string>();
                void AddUnique(string v)
                {
                    if (!string.IsNullOrWhiteSpace(v) && !list.Contains(v))
                        list.Add(v);
                }

                AddUnique(primaryVal);
                if (!string.IsNullOrEmpty(primaryVal) && data.ValueAliases.TryGetValue(primaryVal, out string[] aliases))
                {
                    foreach (string alias in aliases)
                        AddUnique(alias);
                }

                if (includeGlobalResolved)
                {
                    AddUnique(data.ResolvedOcrCode);
                    AddUnique(data.ResolvedOcrName);
                }

                return list;
            }

            // Préparer la liste des écritures à effectuer avant de geler le formulaire
            var plannedWrites = new List<Tuple<string, string, List<string>>>();

            // 1. Colonnes d'axes analytiques standard SAP (Dimensions 1 à 5)
            foreach (KeyValuePair<int, string> kvp in data.DimensionValues)
            {
                int dimCode = kvp.Key;
                string costCenter = kvp.Value;
                if (string.IsNullOrEmpty(costCenter))
                    continue;

                if (colMap.DimensionCols.TryGetValue(dimCode, out string dimColUID) && !string.IsNullOrEmpty(dimColUID))
                {
                    anyColumnFound = true;
                    plannedWrites.Add(Tuple.Create(dimColUID, $"Axe {dimCode}: '{costCenter}' (col {dimColUID})", BuildCandidates(costCenter, includeGlobalResolved: false)));
                }
            }

            // 2. Champs UDF spécifiques (ex: U_DirComm, U_Maintenance, U_Support) s'ils existent sur la ligne
            foreach (KeyValuePair<string, string> udfKvp in data.AllUdfValues)
            {
                if (string.Equals(udfKvp.Key, "U_Analytique", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (colMap.EditableUdfCols.Contains(udfKvp.Key))
                {
                    anyColumnFound = true;
                    plannedWrites.Add(Tuple.Create(udfKvp.Key, $"{udfKvp.Key}: '{udfKvp.Value}'", BuildCandidates(udfKvp.Value, includeGlobalResolved: false)));
                }
            }

            // 3. Colonne UDF principale (ex: U_DocAnalytique)
            if (!string.IsNullOrEmpty(data.CustomUdfValue))
            {
                List<string> customCandidates = BuildCandidates(data.CustomUdfValue, includeGlobalResolved: true);
                if (!string.IsNullOrEmpty(colMap.PrimaryUdfCol))
                {
                    anyColumnFound = true;
                    plannedWrites.Add(Tuple.Create(colMap.PrimaryUdfCol, $"{colMap.PrimaryUdfCol}: '{data.CustomUdfValue}'", customCandidates));
                }
                else if (!anyColumnFound)
                {
                    // Repli générique sur la première colonne d'axe disponible
                    for (int dim = 1; dim <= 5; dim++)
                    {
                        if (colMap.DimensionCols.TryGetValue(dim, out string fallbackCol) && !string.IsNullOrEmpty(fallbackCol))
                        {
                            anyColumnFound = true;
                            plannedWrites.Add(Tuple.Create(fallbackCol, $"{fallbackCol}: '{data.CustomUdfValue}'", customCandidates));
                            break;
                        }
                    }
                }
            }

            if (!anyColumnFound)
            {
                // Invalider le cache au cas où l'utilisateur vient d'activer la colonne dans Paramètres de formulaire
                _matrixColumnCache.Remove(config.FormType);

                string msg = $"SAP_Analytix : Colonnes analytiques introuvables ou masquées dans [{config.FormName}]. Activez les axes (1 à 5) ou '{config.CostCenterCol}' dans Paramètres de formulaire.";
                Log(msg);
                sboApplication.StatusBar.SetText(msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }

            if (plannedWrites.Count == 0)
                return false;

            // Geler le formulaire une seule fois pour toutes les écritures de cellules afin d'éviter les recalculs multiples
            bool frozen = false;
            try
            {
                if (oForm != null)
                {
                    oForm.Freeze(true);
                    frozen = true;
                }

                bool anyWriteFailed = false;
                foreach (var write in plannedWrites)
                {
                    if (WriteValueToMatrixCell(oMatrix, write.Item1, row, write.Item3))
                    {
                        appliedSummary.Add(write.Item2);
                    }
                    else
                    {
                        anyWriteFailed = true;
                    }
                }

                if (anyWriteFailed)
                {
                    // L'utilisateur a peut-être masqué/verrouillé une colonne entre-temps
                    _matrixColumnCache.Remove(config.FormType);
                }
            }
            finally
            {
                if (frozen && oForm != null)
                {
                    try { oForm.Freeze(false); } catch { }
                }
            }

            if (appliedSummary.Count > 0)
            {
                string successMsg = $"[{config.FormName}] Code analytique appliqué ligne {row} -> {string.Join(", ", appliedSummary)}";
                sboApplication.StatusBar.SetText(successMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Scanne les colonnes de la matrice UNE SEULE FOIS par type de formulaire (1 passe unique, sans exceptions COM répétées).
        /// </summary>
        private static MatrixColumnMap GetOrBuildMatrixColumnMap(SAPbouiCOM.Matrix oMatrix, FormConfig config)
        {
            if (_matrixColumnCache.TryGetValue(config.FormType, out MatrixColumnMap cached))
                return cached;

            MatrixColumnMap map = new MatrixColumnMap();
            SAPbouiCOM.Columns oCols = null;

            try
            {
                oCols = oMatrix.Columns;
                int count = oCols.Count;

                // Index des UIDs connus par dimension pour correspondance O(1)
                var knownUidToDim = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in DimensionColumnCandidates)
                {
                    foreach (string uid in kv.Value)
                    {
                        if (!knownUidToDim.ContainsKey(uid))
                            knownUidToDim[uid] = kv.Key;
                    }
                }

                string preferredFound = string.Empty;
                string candidateUdfFound = string.Empty;
                string genericUdfFound = string.Empty;

                for (int i = 0; i < count; i++)
                {
                    SAPbouiCOM.Column col = null;
                    SAPbouiCOM.DataBind db = null;
                    try
                    {
                        col = oCols.Item(i);
                        if (col == null || !col.Visible || !col.Editable)
                            continue;

                        string uid = col.UniqueID ?? string.Empty;
                        if (string.IsNullOrEmpty(uid))
                            continue;

                        // 1. Vérifier si c'est une colonne UDF (U_*)
                        if (uid.StartsWith("U_", StringComparison.OrdinalIgnoreCase))
                        {
                            map.EditableUdfCols.Add(uid);

                            if (string.Equals(uid, config.CostCenterCol, StringComparison.OrdinalIgnoreCase))
                            {
                                preferredFound = uid;
                            }
                            else if (string.IsNullOrEmpty(candidateUdfFound))
                            {
                                foreach (string cand in CandidateCustomUdfCols)
                                {
                                    if (string.Equals(uid, cand, StringComparison.OrdinalIgnoreCase))
                                    {
                                        candidateUdfFound = uid;
                                        break;
                                    }
                                }
                            }

                            if (string.IsNullOrEmpty(genericUdfFound) &&
                                (uid.IndexOf("Analyt", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 uid.IndexOf("Cout", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 uid.IndexOf("Cost", StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                genericUdfFound = uid;
                            }
                        }

                        // 2. Vérifier si c'est l'une des 5 colonnes d'axes analytiques par UID connu
                        if (knownUidToDim.TryGetValue(uid, out int dimFromUid))
                        {
                            if (!map.DimensionCols.ContainsKey(dimFromUid))
                                map.DimensionCols[dimFromUid] = uid;
                            continue;
                        }

                        // 3. Sinon vérifier DataBind.Alias (OcrCode, OcrCode2..5) seulement si tous les 5 axes ne sont pas encore trouvés
                        if (map.DimensionCols.Count < 5 && !uid.StartsWith("U_", StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                db = col.DataBind;
                                string alias = db?.Alias ?? string.Empty;
                                if (!string.IsNullOrEmpty(alias) && knownUidToDim.TryGetValue(alias, out int dimFromAlias))
                                {
                                    if (!map.DimensionCols.ContainsKey(dimFromAlias))
                                        map.DimensionCols[dimFromAlias] = uid;
                                }
                            }
                            catch
                            {
                                // Ignorer
                            }
                        }
                    }
                    catch
                    {
                        // Ignorer
                    }
                    finally
                    {
                        ReleaseCom(ref db);
                        ReleaseCom(ref col);
                    }
                }

                map.PrimaryUdfCol = !string.IsNullOrEmpty(preferredFound)
                    ? preferredFound
                    : (!string.IsNullOrEmpty(candidateUdfFound) ? candidateUdfFound : genericUdfFound);

                if (map.DimensionCols.Count > 0 || !string.IsNullOrEmpty(map.PrimaryUdfCol) || map.EditableUdfCols.Count > 0)
                {
                    _matrixColumnCache[config.FormType] = map;
                }
            }
            finally
            {
                ReleaseCom(ref oCols);
            }

            return map;
        }

        private static bool WriteValueToMatrixCell(SAPbouiCOM.Matrix oMatrix, string targetColUID, int row, List<string> candidateValues)
        {
            if (candidateValues == null || candidateValues.Count == 0)
                return false;

            object cellObj = null;
            try
            {
                cellObj = oMatrix.GetCellSpecific(targetColUID, row);
                if (cellObj == null)
                    return false;

                // Cas 1 : Champ texte (EditText)
                if (cellObj is SAPbouiCOM.EditText oEditCostCenter)
                {
                    string currentValue = oEditCostCenter.Value != null ? oEditCostCenter.Value.Trim() : string.Empty;
                    foreach (string val in candidateValues)
                    {
                        if (string.Equals(currentValue, val, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }

                    foreach (string val in candidateValues)
                    {
                        try
                        {
                            oEditCostCenter.Value = val;
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Log($"Tentative EditText col={targetColUID} val='{val}' échouée : {ex.Message}");
                        }
                    }
                    return false;
                }
                // Cas 2 : Liste déroulante (ComboBox)
                else if (cellObj is SAPbouiCOM.ComboBox oComboCostCenter)
                {
                    string currentValue = oComboCostCenter.Selected?.Value != null ? oComboCostCenter.Selected.Value.Trim() : string.Empty;
                    string currentDesc = oComboCostCenter.Selected?.Description != null ? oComboCostCenter.Selected.Description.Trim() : string.Empty;
                    foreach (string val in candidateValues)
                    {
                        string normVal = NormalizeForComparison(val);
                        if ((!string.IsNullOrEmpty(currentValue) && NormalizeForComparison(currentValue) == normVal) ||
                            (!string.IsNullOrEmpty(currentDesc) && NormalizeForComparison(currentDesc) == normVal))
                        {
                            return true;
                        }
                    }

                    // A. Recherche intelligente dans ValidValues si présent (insensible à la casse, aux espaces et aux accents)
                    SAPbouiCOM.ValidValues validValues = null;
                    try
                    {
                        validValues = oComboCostCenter.ValidValues;
                        if (validValues != null && validValues.Count > 0)
                        {
                            int count = validValues.Count;
                            foreach (string candidate in candidateValues)
                            {
                                string normCandidate = NormalizeForComparison(candidate);

                                for (int i = 0; i < count; i++)
                                {
                                    SAPbouiCOM.ValidValue vv = null;
                                    try
                                    {
                                        vv = validValues.Item(i);
                                        string vCode = vv?.Value?.Trim() ?? string.Empty;
                                        string vDesc = vv?.Description?.Trim() ?? string.Empty;

                                        if ((!string.IsNullOrEmpty(vCode) && NormalizeForComparison(vCode) == normCandidate) ||
                                            (!string.IsNullOrEmpty(vDesc) && NormalizeForComparison(vDesc) == normCandidate))
                                        {
                                            try
                                            {
                                                oComboCostCenter.Select(vCode, BoSearchKey.psk_ByValue);
                                            }
                                            catch
                                            {
                                                oComboCostCenter.Select(i, BoSearchKey.psk_Index);
                                            }
                                            return true;
                                        }
                                    }
                                    finally
                                    {
                                        ReleaseCom(ref vv);
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Fallback sur Select direct
                    }
                    finally
                    {
                        ReleaseCom(ref validValues);
                    }

                    // B. Sélection directe par valeur puis par description
                    foreach (string val in candidateValues)
                    {
                        try
                        {
                            oComboCostCenter.Select(val, BoSearchKey.psk_ByValue);
                            return true;
                        }
                        catch
                        {
                            try
                            {
                                oComboCostCenter.Select(val, BoSearchKey.psk_ByDescription);
                                return true;
                            }
                            catch
                            {
                                // Essayer le candidat suivant
                            }
                        }
                    }

                    return false;
                }

                return false;
            }
            catch (Exception ex)
            {
                Log($"Erreur écriture cellule col={targetColUID}, row={row} : {ex.Message}");
                return false;
            }
            finally
            {
                ReleaseCom(ref cellObj);
            }
        }

        /// <summary>
        /// Normalise une chaîne (majuscules, sans accents, espaces normalisés) pour comparer
        /// "Direction commerciale" avec "DIRECTION COMMERCIALE" ou "Direction Commerciale".
        /// </summary>
        public static string NormalizeForComparison(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            string normalized = input.Trim().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalized.Length);
            bool previousWasSpace = false;
            foreach (char c in normalized)
            {
                UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(c);
                if (uc == UnicodeCategory.NonSpacingMark)
                    continue;

                if (char.IsWhiteSpace(c))
                {
                    if (!previousWasSpace)
                    {
                        sb.Append(' ');
                        previousWasSpace = true;
                    }
                }
                else
                {
                    sb.Append(char.ToUpperInvariant(c));
                    previousWasSpace = false;
                }
            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        /// <summary>
        /// Charge en mémoire UNE SEULE FOIS les tables OOCR et OPRC afin d'éviter des requêtes SQL
        /// répétées à chaque sélection d'article.
        /// </summary>
        public static void EnsureCostCenterCacheLoaded()
        {
            if (_oocrCacheLoaded || oCompany == null || !oCompany.Connected)
                return;

            _oocrCacheLoaded = true;
            Recordset oRec = null;
            SAPbobsCOM.Fields fields = null;

            try
            {
                oRec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // 1. Charger OOCR en une seule requête
                oRec.DoQuery("SELECT \"OcrCode\", \"OcrName\", \"DimCode\" FROM \"OOCR\"");
                while (!oRec.EoF)
                {
                    fields = oRec.Fields;
                    string code = ReadFieldString(fields, "OcrCode");
                    string name = ReadFieldString(fields, "OcrName");
                    string dimStr = ReadFieldString(fields, "DimCode");
                    int.TryParse(dimStr, out int dim);
                    AddCostCenterToCache(code, name, dim);
                    ReleaseCom(ref fields);
                    oRec.MoveNext();
                }

                // 2. Compléter avec OPRC (au cas où un centre de coût n'est pas dans OOCR)
                try
                {
                    oRec.DoQuery("SELECT \"PrcCode\", \"PrcName\", \"DimCode\" FROM \"OPRC\"");
                    while (!oRec.EoF)
                    {
                        fields = oRec.Fields;
                        string code = ReadFieldString(fields, "PrcCode");
                        string name = ReadFieldString(fields, "PrcName");
                        string dimStr = ReadFieldString(fields, "DimCode");
                        int.TryParse(dimStr, out int dim);
                        AddCostCenterToCache(code, name, dim);
                        ReleaseCom(ref fields);
                        oRec.MoveNext();
                    }
                }
                catch
                {
                    // Ignorer si OPRC non disponible
                }
            }
            catch
            {
                // En cas d'erreur au préchargement, on autorisera le fallback SQL unitaire
                _oocrCacheLoaded = false;
            }
            finally
            {
                ReleaseCom(ref fields);
                ReleaseCom(ref oRec);
            }
        }

        private static void AddCostCenterToCache(string code, string name, int dim)
        {
            if (string.IsNullOrEmpty(code) && string.IsNullOrEmpty(name))
                return;

            var entry = new CostCenterCacheEntry
            {
                Code = !string.IsNullOrEmpty(code) ? code : name,
                Name = !string.IsNullOrEmpty(name) ? name : code,
                DimCode = (dim >= 1 && dim <= 5) ? dim : 0
            };

            string normCode = NormalizeForComparison(entry.Code);
            if (!string.IsNullOrEmpty(normCode) && !_oocrCache.ContainsKey(normCode))
                _oocrCache[normCode] = entry;

            string normName = NormalizeForComparison(entry.Name);
            if (!string.IsNullOrEmpty(normName) && !_oocrCache.ContainsKey(normName))
                _oocrCache[normName] = entry;
        }

        /// <summary>
        /// Récupère toutes les données analytiques d'un article (OITM) avec requête SQL ciblée (uniquement les colonnes analytiques)
        /// et cache court (60s) pour une réponse quasi instantanée.
        /// </summary>
        public static ItemAnalyticalData GetAnalyticalDataFromItem(string itemCode)
        {
            if (oCompany == null || !oCompany.Connected || string.IsNullOrEmpty(itemCode))
                return new ItemAnalyticalData();

            // Vérifier le cache article (valide 60 secondes)
            if (_itemCache.TryGetValue(itemCode, out CachedItemEntry cached) &&
                (DateTime.UtcNow - cached.LoadedAtUtc) < ItemCacheTtl)
            {
                return cached.Data;
            }

            ItemAnalyticalData result = new ItemAnalyticalData();
            Recordset oRec = null;
            SAPbobsCOM.Fields fields = null;

            try
            {
                oRec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string safeItemCode = itemCode.Replace("'", "''");
                string query;

                if (_oitmAnalyticalColumns != null && _oitmAnalyticalColumns.Count > 0)
                {
                    // Requête légère ciblée (uniquement les 5-8 colonnes analytiques au lieu des 200+ colonnes de OITM)
                    var colsSql = new List<string>(_oitmAnalyticalColumns.Count);
                    foreach (string c in _oitmAnalyticalColumns)
                        colsSql.Add($"\"{c}\"");
                    query = $"SELECT {string.Join(", ", colsSql)} FROM \"OITM\" WHERE \"ItemCode\" = '{safeItemCode}'";
                }
                else
                {
                    // Première exécution : découvrir automatiquement quelles colonnes analytiques existent dans OITM
                    query = $"SELECT * FROM \"OITM\" WHERE \"ItemCode\" = '{safeItemCode}'";
                }

                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    fields = oRec.Fields;

                    // Si c'est la première exécution, mémoriser la liste exacte des colonnes analytiques présentes dans OITM
                    if (_oitmAnalyticalColumns == null)
                    {
                        var discovered = new List<string>();
                        int totalFields = fields.Count;
                        for (int i = 0; i < totalFields; i++)
                        {
                            SAPbobsCOM.Field f = null;
                            try
                            {
                                f = fields.Item(i);
                                string name = f?.Name ?? string.Empty;
                                if (IsTrackedOitmColumn(name))
                                    discovered.Add(name);
                            }
                            finally
                            {
                                ReleaseCom(ref f);
                            }
                        }
                        if (discovered.Count > 0)
                            _oitmAnalyticalColumns = discovered;
                    }

                    // 1. Lire les 5 axes analytiques standard sur la fiche article (OcrCode, OcrCode2, OcrCode3, OcrCode4, OcrCode5)
                    for (int dim = 1; dim <= 5; dim++)
                    {
                        string ocrFieldName = dim == 1 ? "OcrCode" : $"OcrCode{dim}";
                        if (_oitmAnalyticalColumns != null && !_oitmAnalyticalColumns.Exists(c => string.Equals(c, ocrFieldName, StringComparison.OrdinalIgnoreCase)))
                            continue;

                        string val = ReadFieldString(fields, ocrFieldName);
                        if (!string.IsNullOrEmpty(val))
                        {
                            ResolveCostCenterInSap(val, out string resolvedCode, out string resolvedName, out _);
                            string effectiveCode = !string.IsNullOrEmpty(resolvedCode) ? resolvedCode : val;
                            result.DimensionValues[dim] = effectiveCode;
                            RecordValueAliases(result, effectiveCode, val, resolvedCode, resolvedName);
                        }
                    }

                    // 2. Lire les champs UDF explicites (U_Analytique, U_DocAnalytique, U_CodeAnalytique, U_CentreCout)
                    string[] udfCandidates = new[] { "U_Analytique", "U_DocAnalytique", "U_CodeAnalytique", "U_CentreCout" };
                    foreach (string udfName in udfCandidates)
                    {
                        if (_oitmAnalyticalColumns != null && !_oitmAnalyticalColumns.Exists(c => string.Equals(c, udfName, StringComparison.OrdinalIgnoreCase)))
                            continue;

                        string val = ReadFieldString(fields, udfName);
                        if (!string.IsNullOrEmpty(val))
                        {
                            result.AllUdfValues[udfName] = val;
                            RegisterUdfAnalyticalValue(result, val);
                        }
                    }

                    // 3. Lire les autres champs UDF analytiques découverts dans OITM
                    if (_oitmAnalyticalColumns != null)
                    {
                        foreach (string colName in _oitmAnalyticalColumns)
                        {
                            if (!colName.StartsWith("U_", StringComparison.OrdinalIgnoreCase) || result.AllUdfValues.ContainsKey(colName))
                                continue;

                            string strVal = ReadFieldString(fields, colName);
                            if (!string.IsNullOrEmpty(strVal))
                            {
                                result.AllUdfValues[colName] = strVal;
                                RegisterUdfAnalyticalValue(result, strVal);
                            }
                        }
                    }
                }

                _itemCache[itemCode] = new CachedItemEntry
                {
                    LoadedAtUtc = DateTime.UtcNow,
                    Data = result
                };
            }
            catch (Exception ex)
            {
                Log("Erreur GetAnalyticalDataFromItem : " + ex.Message);
                sboApplication?.StatusBar?.SetText("Erreur lecture fiche article : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                ReleaseCom(ref fields);
                ReleaseCom(ref oRec);
            }

            return result;
        }

        private static bool IsTrackedOitmColumn(string name)
        {
            if (string.IsNullOrEmpty(name))
                return false;

            if (string.Equals(name, "OcrCode", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "OcrCode2", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "OcrCode3", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "OcrCode4", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "OcrCode5", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return name.StartsWith("U_", StringComparison.OrdinalIgnoreCase) && IsAnalyticalUdfFieldName(name);
        }

        private static bool IsAnalyticalUdfFieldName(string name)
        {
            return name.IndexOf("Analyt", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Cout", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Cost", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Axe", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Ocr", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("DirComm", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Direction", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Maintenance", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Support", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void RecordValueAliases(ItemAnalyticalData result, string key, string rawVal, string resolvedCode, string resolvedName)
        {
            var aliases = new List<string>();
            void Add(string s)
            {
                if (!string.IsNullOrWhiteSpace(s) && !aliases.Contains(s))
                    aliases.Add(s);
            }
            Add(rawVal);
            Add(resolvedCode);
            Add(resolvedName);
            if (aliases.Count > 0 && !string.IsNullOrEmpty(key))
            {
                result.ValueAliases[key] = aliases.ToArray();
            }
        }

        private static void RegisterUdfAnalyticalValue(ItemAnalyticalData result, string rawVal)
        {
            if (string.IsNullOrEmpty(rawVal))
                return;

            if (string.IsNullOrEmpty(result.CustomUdfValue))
                result.CustomUdfValue = rawVal;

            int dimCode = ResolveCostCenterInSap(rawVal, out string resolvedCode, out string resolvedName, out _);
            if (!string.IsNullOrEmpty(resolvedCode) && string.IsNullOrEmpty(result.ResolvedOcrCode))
                result.ResolvedOcrCode = resolvedCode;
            if (!string.IsNullOrEmpty(resolvedName) && string.IsNullOrEmpty(result.ResolvedOcrName))
                result.ResolvedOcrName = resolvedName;

            string effectiveCode = !string.IsNullOrEmpty(resolvedCode) ? resolvedCode : rawVal;
            RecordValueAliases(result, rawVal, rawVal, resolvedCode, resolvedName);
            RecordValueAliases(result, effectiveCode, rawVal, resolvedCode, resolvedName);

            if (dimCode >= 1 && dimCode <= 5 && !result.DimensionValues.ContainsKey(dimCode))
            {
                result.DimensionValues[dimCode] = effectiveCode;
            }
        }

        /// <summary>
        /// Résout un centre de coût (par Code ou par Nom) en priorité depuis le cache mémoire OOCR/OPRC (0 ms),
        /// puis en fallback SQL si absent du cache.
        /// </summary>
        public static int ResolveCostCenterInSap(string codeOrName, out string resolvedCode, out string resolvedName, out int resolvedDim)
        {
            resolvedCode = string.Empty;
            resolvedName = string.Empty;
            resolvedDim = 0;

            if (oCompany == null || !oCompany.Connected || string.IsNullOrEmpty(codeOrName))
                return 0;

            // 1. Lookup instantané dans le cache mémoire OOCR/OPRC
            string normKey = NormalizeForComparison(codeOrName);
            if (_oocrCache.TryGetValue(normKey, out CostCenterCacheEntry cachedEntry))
            {
                resolvedCode = cachedEntry.Code;
                resolvedName = cachedEntry.Name;
                resolvedDim = cachedEntry.DimCode;
                return resolvedDim;
            }

            // Si le cache global a déjà été chargé et ne contient pas ce code, ne pas re-requêter SQL inutilement
            if (_oocrCacheLoaded && _oocrCache.Count > 0)
                return 0;

            // 2. Fallback SQL unitaire (si le préchargement n'a pas encore tourné)
            Recordset oRec = null;
            SAPbobsCOM.Fields fields = null;

            try
            {
                oRec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string safeVal = codeOrName.Replace("'", "''");

                string query = $"SELECT \"OcrCode\", \"OcrName\", \"DimCode\" FROM \"OOCR\" WHERE UPPER(\"OcrCode\") = UPPER('{safeVal}') OR UPPER(\"OcrName\") = UPPER('{safeVal}')";
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    fields = oRec.Fields;
                    resolvedCode = ReadFieldString(fields, "OcrCode");
                    resolvedName = ReadFieldString(fields, "OcrName");
                    string dimStr = ReadFieldString(fields, "DimCode");
                    if (int.TryParse(dimStr, out int dim) && dim >= 1 && dim <= 5)
                    {
                        resolvedDim = dim;
                    }
                    if (string.IsNullOrEmpty(resolvedCode))
                        resolvedCode = codeOrName;

                    AddCostCenterToCache(resolvedCode, resolvedName, resolvedDim);
                    return resolvedDim;
                }

                ReleaseCom(ref fields);
                string queryOprc = $"SELECT \"PrcCode\", \"PrcName\", \"DimCode\" FROM \"OPRC\" WHERE UPPER(\"PrcCode\") = UPPER('{safeVal}') OR UPPER(\"PrcName\") = UPPER('{safeVal}')";
                oRec.DoQuery(queryOprc);

                if (!oRec.EoF)
                {
                    fields = oRec.Fields;
                    resolvedCode = ReadFieldString(fields, "PrcCode");
                    resolvedName = ReadFieldString(fields, "PrcName");
                    string dimStr = ReadFieldString(fields, "DimCode");
                    if (int.TryParse(dimStr, out int dim) && dim >= 1 && dim <= 5)
                    {
                        resolvedDim = dim;
                    }
                    if (string.IsNullOrEmpty(resolvedCode))
                        resolvedCode = codeOrName;

                    AddCostCenterToCache(resolvedCode, resolvedName, resolvedDim);
                    return resolvedDim;
                }
            }
            catch
            {
                // Ignorer si les tables ne sont pas accessibles
            }
            finally
            {
                ReleaseCom(ref fields);
                ReleaseCom(ref oRec);
            }

            return 0;
        }

        public static int GetDimensionForCostCenter(string costCenterCode)
        {
            return ResolveCostCenterInSap(costCenterCode, out _, out _, out _);
        }

        private static string ReadFieldString(SAPbobsCOM.Fields fields, string fieldName)
        {
            SAPbobsCOM.Field field = null;
            try
            {
                field = fields.Item(fieldName);
                object val = field?.Value;
                if (val != null && !(val is DBNull))
                {
                    return val.ToString().Trim();
                }
            }
            catch
            {
                // Le champ n'existe pas
            }
            finally
            {
                ReleaseCom(ref field);
            }
            return string.Empty;
        }

        private static void ReleaseCom<T>(ref T comObj) where T : class
        {
            if (comObj != null)
            {
                try
                {
                    if (Marshal.IsComObject(comObj))
                    {
                        Marshal.ReleaseComObject(comObj);
                    }
                }
                catch
                {
                    // Ignorer en mode test unitaire
                }
                comObj = null;
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
                    ReleaseCom(ref oCompany);
                }

                ReleaseCom(ref sboApplication);

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
