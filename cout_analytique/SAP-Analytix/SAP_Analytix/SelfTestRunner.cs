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
    /// pour vérifier hors-ligne tous les scénarios métier de SAP_Analytix.
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
            LogLine("  EXECUTION DES TESTS AUTOMATISES SAP_ANALYTIX (SIMULATION) ");
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

            // Test 4 : Vérifier le repli sur la colonne standard SAP Règle de répartition ("110000109")
            AssertTest("Repli automatique sur la colonne standard SAP Centre de coûts ('110000109')", () =>
                RunScenario("139", BoEventTypes.et_VALIDATE, "ART-STD", "CC-STD-01", "110000109", useComboBox: false));

            // Test 5 : Vérifier l'événement et_LOST_FOCUS après saisie
            AssertTest("Déclenchement sur sortie de cellule (et_LOST_FOCUS) sur Livraison (140)", () =>
                RunScenario("140", BoEventTypes.et_LOST_FOCUS, "ART-LVR", "CC-LOGISTIQUE", "U_DocAnalytique", useComboBox: false));

            // Test 6 : Vérifier le message d'avertissement quand l'article n'a pas de code analytique
            AssertTest("Avertissement dans la barre d'état si l'article n'a pas de U_Analytique", () =>
            {
                string lastStatusText = string.Empty;
                BoStatusBarMessageType lastMsgType = BoStatusBarMessageType.smt_None;

                SetupMockEnvironment(
                    itemDbValues: new Dictionary<string, string> { { "U_Analytique", "" } },
                    matrixColName: "U_DocAnalytique",
                    itemCellValue: "ART-VIDE",
                    useComboBox: false,
                    onStatusText: (txt, t) => { lastStatusText = txt; lastMsgType = t; },
                    out var getWrittenValue);

                ItemEvent evt = CreateMockItemEvent("139", "38", "1", 1, BoEventTypes.et_VALIDATE, itemChanged: true, cflItemCode: null);
                Program.SBO_Application_ItemEvent("F_1", ref evt, out bool bubble);

                return bubble &&
                       string.IsNullOrEmpty(getWrittenValue()) &&
                       lastMsgType == BoStatusBarMessageType.smt_Warning &&
                       lastStatusText.Contains("ART-VIDE");
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
                matrixColName: activeColName,
                itemCellValue: eventType == BoEventTypes.et_CHOOSE_FROM_LIST ? "" : itemCode,
                useComboBox: useComboBox,
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

        private static void SetupMockEnvironment(
            Dictionary<string, string> itemDbValues,
            string matrixColName,
            string itemCellValue,
            bool useComboBox,
            Action<string, BoStatusBarMessageType> onStatusText,
            out Func<string> getWrittenValue)
        {
            string writtenValue = string.Empty;
            getWrittenValue = () => writtenValue;

            // Mock DI API (oCompany + Recordset)
            var fieldNames = new List<string>(itemDbValues.Keys);
            var mockFields = DynamicComMock.Create<SAPbobsCOM.Fields>((method, args) =>
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

            var mockRecordset = DynamicComMock.Create<SAPbobsCOM.Recordset>((method, args) =>
            {
                if (method == "DoQuery") return null;
                if (method == "get_EoF") return false;
                if (method == "get_Fields") return mockFields;
                return null;
            });

            Program.oCompany = DynamicComMock.Create<SAPbobsCOM.Company>((method, args) =>
            {
                if (method == "get_Connected") return true;
                if (method == "GetBusinessObject") return mockRecordset;
                return null;
            });

            // Mock UI API (Matrix Columns, Cells, Form, Application)
            var mockCol = DynamicComMock.Create<SAPbouiCOM.Column>((method, args) =>
            {
                if (method == "get_UniqueID") return matrixColName;
                if (method == "get_Visible") return true;
                if (method == "get_Editable") return true;
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
                mockTargetCell = DynamicComMock.Create<SAPbouiCOM.ComboBox>((method, args) =>
                {
                    if (method == "get_Selected") return null;
                    if (method == "Select") { writtenValue = Convert.ToString(args[0]); return null; }
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
