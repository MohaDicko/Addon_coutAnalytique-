using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SAP_Analytix
{
    /// <summary>
    /// Suite d'auto-tests simulant les objets COM de SAP Business One (UI API & DI API)
    /// pour vérifier hors-ligne tous les scénarios métier de SAP_Analytix (incluant les 5 axes analytiques).
    /// </summary>
    public static class SelfTestRunner
    {
        private class DynamicComMock : RealProxy, System.Runtime.Remoting.IRemotingTypeInfo
        {
            private readonly Type _interfaceType;
            private readonly Func<string, object[], object> _handler;

            public DynamicComMock(Type interfaceType, Func<string, object[], object> handler)
                : base(interfaceType)
            {
                _interfaceType = interfaceType;
                _handler = handler;
            }

            public string TypeName
            {
                get => _interfaceType.FullName;
                set { }
            }

            public bool CanCastTo(Type fromType, object o)
            {
                return fromType.IsAssignableFrom(_interfaceType) ||
                       (fromType == typeof(SAPbouiCOM.ItemEvent) && _interfaceType == typeof(SAPbouiCOM.IChooseFromListEvent)) ||
                       (fromType == typeof(SAPbouiCOM.IChooseFromListEvent) && _interfaceType == typeof(SAPbouiCOM.ItemEvent));
            }

            public override System.Runtime.Remoting.Messaging.IMessage Invoke(System.Runtime.Remoting.Messaging.IMessage msg)
            {
                IMethodCallMessage call = (IMethodCallMessage)msg;
                object result = _handler(call.MethodName, call.InArgs);
                return new ReturnMessage(result, null, 0, call.LogicalCallContext, call);
            }

            public static T Create<T>(Func<string, object[], object> handler)
            {
                return (T)new DynamicComMock(typeof(T), handler).GetTransparentProxy();
            }
        }

        public static int RunAllTests()
        {
            var sb = new System.Text.StringBuilder();
            void LogLine(string line)
            {
                Console.WriteLine(line);
                sb.AppendLine(line);
            }

            LogLine("============================================================");
            LogLine("  EXECUTION DES TESTS AUTOMATISES SAP_ANALYTIX (5 AXES)     ");
            LogLine("============================================================");

            int passed = 0;
            int failed = 0;

            void AssertTest(string name, Func<bool> test)
            {
                try
                {
                    bool ok = test();
                    if (ok)
                    {
                        LogLine($"[OK]   {name}");
                        passed++;
                    }
                    else
                    {
                        LogLine($"[FAIL] {name}");
                        failed++;
                    }
                }
                catch (Exception ex)
                {
                    LogLine($"[ERR]  {name} -> {ex.Message}");
                    failed++;
                }
            }

            // Test 1 : Vérifier que les 6 formulaires principaux (codes UI API) sont enregistrés
            string[] requiredForms = new[] { "139", "142", "143", "140", "133", "141" };
            foreach (string formType in requiredForms)
            {
                string ft = formType;
                AssertTest($"Formulaire SAP FormTypeEx='{ft}' reconnu et actif (et_VALIDATE + EditText)", () =>
                    RunScenario(ft, BoEventTypes.et_VALIDATE, "ART-001", "CC-VENTES", "U_DocAnalytique", useComboBox: false));
            }

            // Test 2 : Vérifier l'événement et_CHOOSE_FROM_LIST (sélection via la loupe / Tab)
            AssertTest("Sélection d'article via la loupe (et_CHOOSE_FROM_LIST) sur Commande Fournisseur (142)", () =>
                RunScenario("142", BoEventTypes.et_CHOOSE_FROM_LIST, "ART-ACHAT", "CC-ACHATS", "U_DocAnalytique", useComboBox: false));

            // Test 3 : Vérifier le support d'une colonne ComboBox (UDF avec liste de valeurs autorisées)
            AssertTest("Support d'une colonne U_DocAnalytique de type ComboBox (liste déroulante)", () =>
                RunScenario("133", BoEventTypes.et_VALIDATE, "ART-COMBO", "CC-LABO", "U_DocAnalytique", useComboBox: true));

            // Test 4 : Vérifier les 5 Axes analytiques SAP (Axes 1, 2, et surtout Axes 3=Direction commerciale, 4=Maintenance, 5=Support)
            AssertTest("Axe 1 (OcrCode / col '2004') rempli automatiquement", () =>
                RunDimensionScenario(dimCode: 1, matrixColUID: "2004", matrixColAlias: "OcrCode", costCenter: "CC-AXE1"));

            AssertTest("Axe 2 (OcrCode2 / col '2005') rempli automatiquement", () =>
                RunDimensionScenario(dimCode: 2, matrixColUID: "2005", matrixColAlias: "OcrCode2", costCenter: "CC-AXE2"));

            AssertTest("Axe 3 - Direction commerciale (OcrCode3 / col '2006' ou '10001068') rempli automatiquement", () =>
                RunDimensionScenario(dimCode: 3, matrixColUID: "10001068", matrixColAlias: "OcrCode3", costCenter: "DIR-COMM"));

            AssertTest("Axe 4 - Maintenance (OcrCode4 / col '2007' ou '10001070') rempli automatiquement", () =>
                RunDimensionScenario(dimCode: 4, matrixColUID: "10001070", matrixColAlias: "OcrCode4", costCenter: "MAINTENANCE"));

            AssertTest("Axe 5 - Support (OcrCode5 / col '2008' ou '10001072') rempli automatiquement", () =>
                RunDimensionScenario(dimCode: 5, matrixColUID: "10001072", matrixColAlias: "OcrCode5", costCenter: "SUPPORT"));

            // Test 5 : Vérifier le routage intelligent d'un UDF U_Analytique vers l'Axe 3/4/5 via la table OOCR (DimCode)
            AssertTest("Routage automatique de U_Analytique vers la colonne Axe 3 (Direction commerciale) via OOCR.DimCode=3", () =>
                RunUdfToDimensionRoutingScenario(udfCostCenter: "CC-DIR-COMM", oocrDimCode: 3, matrixColUID: "2006", matrixColAlias: "OcrCode3"));

            // Test 6 : Résolution par Nom/Libellé dans OITM.U_Analytique vers le OcrCode et l'Axe correspondant via OOCR
            AssertTest("Résolution par libellé 'Direction commerciale' dans U_Analytique -> OcrCode 'CC_DIR_COM' sur Axe 3 (col '2006')", () =>
                RunLabelResolutionScenario("Direction commerciale", "CC_DIR_COM", dimCode: 3, matrixColUID: "2006", matrixColAlias: "OcrCode3"));

            AssertTest("Résolution par libellé 'Maintenance' dans U_Analytique -> OcrCode 'CC_MAINT' sur Axe 4 (col '2007')", () =>
                RunLabelResolutionScenario("Maintenance", "CC_MAINT", dimCode: 4, matrixColUID: "2007", matrixColAlias: "OcrCode4"));

            AssertTest("Résolution par libellé 'Support' dans U_Analytique -> OcrCode 'CC_SUPP' sur Axe 5 (col '2008')", () =>
                RunLabelResolutionScenario("Support", "CC_SUPP", dimCode: 5, matrixColUID: "2008", matrixColAlias: "OcrCode5"));

            // Test 7 : Sélection dans un ComboBox (ValidValues) insensible à la casse/accents entre OITM et ValidValues
            AssertTest("ComboBox ValidValues : 'Direction commerciale' sélectionne Value='CC_DIR_COM' (Description='DIRECTION COMMERCIALE')", () =>
            {
                string lastStatusText = string.Empty;
                BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;

                var validValues = new List<Tuple<string, string>>
                {
                    Tuple.Create("CC_GEN", "GÉNÉRAL"),
                    Tuple.Create("CC_DIR_COM", "DIRECTION COMMERCIALE"),
                    Tuple.Create("CC_MAINT", "MAINTENANCE"),
                    Tuple.Create("CC_SUPP", "SUPPORT")
                };

                SetupMockEnvironment(
                    itemDbValues: new Dictionary<string, string> { { "U_Analytique", "Direction commerciale" } },
                    oocrEntries: null,
                    matrixColName: "U_DocAnalytique",
                    matrixColAlias: "U_DocAnalytique",
                    itemCellValue: "ART-COMBO-DIR",
                    useComboBox: true,
                    comboValidValues: validValues,
                    onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                    out var getWrittenValue);

                ItemEvent evt = CreateMockItemEvent("139", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
                Program.SBO_Application_ItemEvent("F_COMBO_VV", ref evt, out bool bubble);

                return bubble &&
                       getWrittenValue() == "CC_DIR_COM" &&
                       lastMsgType == BoStatusBarMessageType.smt_Success;
            });

            // Test 8 : Vérifier le message d'avertissement quand l'article n'a pas de code analytique
            AssertTest("Avertissement dans la barre d'état si l'article n'a aucun code analytique", () =>
            {
                string lastStatusText = string.Empty;
                BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;

                SetupMockEnvironment(
                    itemDbValues: new Dictionary<string, string> { { "U_Analytique", "" } },
                    oocrEntries: null,
                    matrixColName: "U_DocAnalytique",
                    matrixColAlias: "U_DocAnalytique",
                    itemCellValue: "ART-VIDE",
                    useComboBox: false,
                    comboValidValues: null,
                    onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                    out var getWrittenValue);

                ItemEvent evt = CreateMockItemEvent("139", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
                Program.SBO_Application_ItemEvent("F_1", ref evt, out bool bubble);

                return bubble &&
                       string.IsNullOrEmpty(getWrittenValue()) &&
                       lastMsgType == BoStatusBarMessageType.smt_Warning &&
                       lastStatusText.Contains("ART-VIDE");
            });

            // Test 9 : Performance / Déduplication — et_VALIDATE après et_CHOOSE_FROM_LIST ne ré-exécute pas le traitement
            AssertTest("Performance : déduplication (et_VALIDATE après et_CHOOSE_FROM_LIST sur la même ligne = 0 recalcul)", () =>
            {
                int statusCalls = 0;
                SetupMockEnvironment(
                    itemDbValues: new Dictionary<string, string> { { "OcrCode3", "DIR-COMM" } },
                    oocrEntries: new List<Tuple<string, string, int>> { Tuple.Create("DIR-COMM", "Direction commerciale", 3) },
                    matrixColName: "2006",
                    matrixColAlias: "OcrCode3",
                    itemCellValue: "ART-PERF",
                    useComboBox: false,
                    comboValidValues: null,
                    onStatusText: (txt, t) => { statusCalls++; },
                    out var getWrittenValue);

                ItemEvent evtCfl = CreateMockItemEvent("142", "38", "1", 1, BoEventTypes.et_CHOOSE_FROM_LIST, itemChanged: true, cflItemCode: "ART-PERF");
                Program.SBO_Application_ItemEvent("F_PERF", ref evtCfl, out bool bubble1);

                ItemEvent evtVal = CreateMockItemEvent("142", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
                Program.SBO_Application_ItemEvent("F_PERF", ref evtVal, out bool bubble2);

                return bubble1 && bubble2 && getWrittenValue() == "DIR-COMM" && statusCalls == 1;
            });

            LogLine("------------------------------------------------------------");
            LogLine($"Résultat : {passed} test(s) réussi(s), {failed} échec(s).");
            LogLine("============================================================");

            string reportPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "selftest_results.txt");
            System.IO.File.WriteAllText(reportPath, sb.ToString(), System.Text.Encoding.UTF8);

            return failed == 0 ? 0 : 1;
        }

        private static bool RunScenario(string formTypeEx, BoEventTypes eventType, string itemCode, string expectedCostCenter, string activeColName, bool useComboBox)
        {
            string lastStatusText = string.Empty;
            BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;

            SetupMockEnvironment(
                itemDbValues: new Dictionary<string, string> { { "U_Analytique", expectedCostCenter } },
                oocrEntries: null,
                matrixColName: activeColName,
                matrixColAlias: activeColName,
                itemCellValue: eventType == BoEventTypes.et_CHOOSE_FROM_LIST ? "" : itemCode,
                useComboBox: useComboBox,
                comboValidValues: null,
                onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                out var getWrittenValue);

            ItemEvent evt = CreateMockItemEvent(
                formTypeEx,
                "38",
                "1",
                1,
                eventType,
                itemChanged: true,
                cflItemCode: eventType == BoEventTypes.et_CHOOSE_FROM_LIST ? itemCode : null);

            Program.SBO_Application_ItemEvent("F_TEST", ref evt, out bool bubble);

            string actual = getWrittenValue();
            return bubble &&
                   actual == expectedCostCenter &&
                   lastMsgType == BoStatusBarMessageType.smt_Success &&
                   lastStatusText.Contains(expectedCostCenter);
        }

        private static bool RunDimensionScenario(int dimCode, string matrixColUID, string matrixColAlias, string costCenter)
        {
            string lastStatusText = string.Empty;
            BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;
            string itemField = dimCode == 1 ? "OcrCode" : $"OcrCode{dimCode}";

            SetupMockEnvironment(
                itemDbValues: new Dictionary<string, string> { { itemField, costCenter } },
                oocrEntries: new List<Tuple<string, string, int>> { Tuple.Create(costCenter, costCenter, dimCode) },
                matrixColName: matrixColUID,
                matrixColAlias: matrixColAlias,
                itemCellValue: $"ART-AXE{dimCode}",
                useComboBox: false,
                comboValidValues: null,
                onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                out var getWrittenValue);

            ItemEvent evt = CreateMockItemEvent("139", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
            Program.SBO_Application_ItemEvent("F_DIM", ref evt, out bool bubble);

            return bubble &&
                   getWrittenValue() == costCenter &&
                   lastMsgType == BoStatusBarMessageType.smt_Success &&
                   lastStatusText.Contains(costCenter);
        }

        private static bool RunUdfToDimensionRoutingScenario(string udfCostCenter, int oocrDimCode, string matrixColUID, string matrixColAlias)
        {
            string lastStatusText = string.Empty;
            BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;

            SetupMockEnvironment(
                itemDbValues: new Dictionary<string, string> { { "U_Analytique", udfCostCenter } },
                oocrEntries: new List<Tuple<string, string, int>> { Tuple.Create(udfCostCenter, udfCostCenter, oocrDimCode) },
                matrixColName: matrixColUID,
                matrixColAlias: matrixColAlias,
                itemCellValue: "ART-UDF-DIM3",
                useComboBox: false,
                comboValidValues: null,
                onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                out var getWrittenValue);

            ItemEvent evt = CreateMockItemEvent("142", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
            Program.SBO_Application_ItemEvent("F_UDF_DIM", ref evt, out bool bubble);

            return bubble &&
                   getWrittenValue() == udfCostCenter &&
                   lastMsgType == BoStatusBarMessageType.smt_Success &&
                   lastStatusText.Contains(udfCostCenter);
        }

        private static bool RunLabelResolutionScenario(string labelInItem, string expectedOcrCode, int dimCode, string matrixColUID, string matrixColAlias)
        {
            string lastStatusText = string.Empty;
            BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;

            SetupMockEnvironment(
                itemDbValues: new Dictionary<string, string> { { "U_Analytique", labelInItem } },
                oocrEntries: new List<Tuple<string, string, int>> { Tuple.Create(expectedOcrCode, labelInItem, dimCode) },
                matrixColName: matrixColUID,
                matrixColAlias: matrixColAlias,
                itemCellValue: $"ART-LBL-{dimCode}",
                useComboBox: false,
                comboValidValues: null,
                onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                out var getWrittenValue);

            ItemEvent evt = CreateMockItemEvent("133", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
            Program.SBO_Application_ItemEvent("F_LBL", ref evt, out bool bubble);

            return bubble &&
                   getWrittenValue() == expectedOcrCode &&
                   lastMsgType == BoStatusBarMessageType.smt_Success &&
                   lastStatusText.Contains(expectedOcrCode);
        }

        private static void SetupMockEnvironment(
            Dictionary<string, string> itemDbValues,
            List<Tuple<string, string, int>> oocrEntries,
            string matrixColName,
            string matrixColAlias,
            string itemCellValue,
            bool useComboBox,
            List<Tuple<string, string>> comboValidValues,
            Action<string, BoStatusBarMessageType> onStatusText,
            out Func<string> getWrittenValue)
        {
            Program.ResetCaches();
            string writtenValue = string.Empty;
            getWrittenValue = () => writtenValue;

            var fieldNames = new List<string>(itemDbValues.Keys);
            var mockItemFields = DynamicComMock.Create<SAPbobsCOM.Fields>((method, args) =>
            {
                if (method == "get_Count") return fieldNames.Count;
                if (method == "Item")
                {
                    string key = args[0] is int idx ? fieldNames[idx] : Convert.ToString(args[0]);
                    if (!itemDbValues.TryGetValue(key, out string val))
                        throw new Exception("Field not found");
                    return DynamicComMock.Create<SAPbobsCOM.Field>((m, a) =>
                    {
                        if (m == "get_Name") return key;
                        if (m == "get_Value") return val;
                        return null;
                    });
                }
                return null;
            });

            Tuple<string, string, int> FindMatchingOocr(string query)
            {
                if (oocrEntries == null || string.IsNullOrEmpty(query))
                    return null;
                foreach (var entry in oocrEntries)
                {
                    if (query.IndexOf($"'{entry.Item1}'", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        query.IndexOf($"'{entry.Item2}'", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return entry;
                    }
                }
                return null;
            }

            Program.oCompany = DynamicComMock.Create<SAPbobsCOM.Company>((method, args) =>
            {
                if (method == "get_Connected") return true;
                if (method == "GetBusinessObject")
                {
                    string lastQuery = string.Empty;
                    return DynamicComMock.Create<SAPbobsCOM.Recordset>((rm, ra) =>
                    {
                        if (rm == "DoQuery") { lastQuery = Convert.ToString(ra[0]); return null; }
                        if (rm == "get_EoF")
                        {
                            if (lastQuery.Contains("\"OOCR\""))
                                return FindMatchingOocr(lastQuery) == null;
                            if (lastQuery.Contains("\"OPRC\""))
                                return true;
                            return false;
                        }
                        if (rm == "get_Fields")
                        {
                            if (lastQuery.Contains("\"OOCR\""))
                            {
                                var match = FindMatchingOocr(lastQuery);
                                return DynamicComMock.Create<SAPbobsCOM.Fields>((fm, fa) =>
                                {
                                    if (fm == "Item")
                                    {
                                        string colKey = Convert.ToString(fa[0]);
                                        object colVal = null;
                                        if (match != null)
                                        {
                                            if (colKey == "OcrCode") colVal = match.Item1;
                                            else if (colKey == "OcrName") colVal = match.Item2;
                                            else if (colKey == "DimCode") colVal = match.Item3;
                                        }
                                        return DynamicComMock.Create<SAPbobsCOM.Field>((flm, fla) =>
                                        {
                                            if (flm == "get_Value") return colVal;
                                            return null;
                                        });
                                    }
                                    return null;
                                });
                            }
                            return mockItemFields;
                        }
                        return null;
                    });
                }
                return null;
            });

            var mockDataBind = DynamicComMock.Create<SAPbouiCOM.DataBind>((method, args) =>
            {
                if (method == "get_Alias") return matrixColAlias;
                return null;
            });

            var mockCol = DynamicComMock.Create<SAPbouiCOM.Column>((method, args) =>
            {
                if (method == "get_UniqueID") return matrixColName;
                if (method == "get_Visible") return true;
                if (method == "get_Editable") return true;
                if (method == "get_DataBind") return mockDataBind;
                return null;
            });

            var mockColumns = DynamicComMock.Create<SAPbouiCOM.Columns>((method, args) =>
            {
                if (method == "get_Count") return 1;
                if (method == "Item")
                {
                    if (args[0] is int i && i == 0) return mockCol;
                    if (Convert.ToString(args[0]) == matrixColName) return mockCol;
                    throw new Exception("Column not found");
                }
                return null;
            });

            var mockItemCell = DynamicComMock.Create<SAPbouiCOM.EditText>((method, args) =>
            {
                if (method == "get_Value") return itemCellValue;
                return null;
            });

            object mockTargetCell;
            if (!useComboBox)
            {
                mockTargetCell = DynamicComMock.Create<SAPbouiCOM.EditText>((method, args) =>
                {
                    if (method == "get_Value") return writtenValue;
                    if (method == "set_Value") { writtenValue = Convert.ToString(args[0]); return null; }
                    return null;
                });
            }
            else
            {
                SAPbouiCOM.ValidValues mockValidValues = null;
                if (comboValidValues != null)
                {
                    mockValidValues = DynamicComMock.Create<SAPbouiCOM.ValidValues>((vvm, vva) =>
                    {
                        if (vvm == "get_Count") return comboValidValues.Count;
                        if (vvm == "Item")
                        {
                            int idx = Convert.ToInt32(vva[0]);
                            var entry = comboValidValues[idx];
                            return DynamicComMock.Create<SAPbouiCOM.ValidValue>((vm, va) =>
                            {
                                if (vm == "get_Value") return entry.Item1;
                                if (vm == "get_Description") return entry.Item2;
                                return null;
                            });
                        }
                        return null;
                    });
                }

                mockTargetCell = DynamicComMock.Create<SAPbouiCOM.ComboBox>((method, args) =>
                {
                    if (method == "get_Selected") return null;
                    if (method == "get_ValidValues") return mockValidValues;
                    if (method == "Select")
                    {
                        BoSearchKey keyType = (BoSearchKey)args[1];
                        if (keyType == BoSearchKey.psk_Index && comboValidValues != null)
                        {
                            int idx = Convert.ToInt32(args[0]);
                            writtenValue = comboValidValues[idx].Item1;
                        }
                        else
                        {
                            writtenValue = Convert.ToString(args[0]);
                        }
                        return null;
                    }
                    return null;
                });
            }

            var mockMatrix = DynamicComMock.Create<SAPbouiCOM.Matrix>((method, args) =>
            {
                if (method == "get_RowCount") return 5;
                if (method == "get_Columns") return mockColumns;
                if (method == "GetCellSpecific")
                {
                    string col = Convert.ToString(args[0]);
                    if (col == "1") return mockItemCell;
                    if (col == matrixColName) return mockTargetCell;
                }
                return null;
            });

            var mockFormItem = DynamicComMock.Create<SAPbouiCOM.Item>((method, args) =>
            {
                if (method == "get_Specific") return mockMatrix;
                return null;
            });

            var mockFormItems = DynamicComMock.Create<SAPbouiCOM.Items>((method, args) =>
            {
                if (method == "Item") return mockFormItem;
                return null;
            });

            var mockForm = DynamicComMock.Create<SAPbouiCOM.Form>((method, args) =>
            {
                if (method == "get_Items") return mockFormItems;
                if (method == "Freeze") return null;
                return null;
            });

            var mockForms = DynamicComMock.Create<SAPbouiCOM.Forms>((method, args) =>
            {
                if (method == "Item") return mockForm;
                return null;
            });

            var mockStatusBar = DynamicComMock.Create<SAPbouiCOM.StatusBar>((method, args) =>
            {
                if (method == "SetText")
                {
                    onStatusText?.Invoke(Convert.ToString(args[0]), (BoStatusBarMessageType)args[2]);
                }
                return null;
            });

            Program.sboApplication = DynamicComMock.Create<SAPbouiCOM.Application>((method, args) =>
            {
                if (method == "get_Forms") return mockForms;
                if (method == "get_StatusBar") return mockStatusBar;
                return null;
            });
        }

        private static ItemEvent CreateMockItemEvent(
            string formTypeEx,
            string itemUID,
            string colUID,
            int row,
            BoEventTypes eventType,
            bool itemChanged,
            string cflItemCode)
        {
            var mockDataTable = cflItemCode == null ? null : DynamicComMock.Create<SAPbouiCOM.DataTable>((m, a) =>
            {
                if (m == "get_Rows")
                {
                    return DynamicComMock.Create<SAPbouiCOM.DataRows>((rm, ra) =>
                    {
                        if (rm == "get_Count") return 1;
                        return null;
                    });
                }
                if (m == "GetValue") return cflItemCode;
                return null;
            });

            return DynamicComMock.Create<SAPbouiCOM.IChooseFromListEvent>((method, args) =>
            {
                switch (method)
                {
                    case "get_BeforeAction":
                    case "get_Before_Action":
                        return false;
                    case "get_FormTypeEx":
                        return formTypeEx;
                    case "get_ItemUID":
                        return itemUID;
                    case "get_ColUID":
                        return colUID;
                    case "get_Row":
                        return row;
                    case "get_EventType":
                        return eventType;
                    case "get_ItemChanged":
                        return itemChanged;
                    case "get_SelectedObjects":
                        return mockDataTable;
                    default:
                        return null;
                }
            }) as ItemEvent;
        }
    }
}
