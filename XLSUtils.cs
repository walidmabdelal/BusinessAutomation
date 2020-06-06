/*
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BaseLineD
{

    static class XLS
    {
        private static Application engXLS, BLXLS;
        private static Range tmpRange, provColRange, engColRange, blColRange;
        private static Workbook engWorkbook, provWorkbook, blWorkbook;
        private static Worksheet engSheet, provSheet, blSheet, intrmSheet;
        private static int BLastLine;
        private static string search_value, eng_children_set, prov_children_set, prev_search_value;
        private static readonly string[] BlHeaderNetwork = new string[] {"Network", "Network Code", "Zone_No", "Zone", "Cluster_No", "Cluster Description",
            "Segment_No", "Segment Description"};
        private static readonly string[] BlHeaderFunction = new string[] {"Discipline", "Discipline Code", "System", "System Code",
            "S_System", "S_System Code", "Element", "Element Code", "Component", "Component Code", "Item", "Item Code"};
        private static readonly string[] BlHeaderInformation = new string[] { "Type", "Type_Code", "Info_Class", "Info_Class Code" };
        private static readonly string[] BlStatus = new string[] { "Removed", "Added", "Modified Description", "Hierarchical Change"};
        private static List<string> prov_children = new List<string>();
        private static List<string> eng_children = new List<string>();


        internal static (Excel.Workbook, Excel.Workbook, Excel.Workbook) XlsInitialize(string idDBPath, string provPath)
        {
            try
            {
                // Create an instance of Microsoft Excel and make it invisible 
                engXLS = new Application();
                BLXLS = new Application();
                engXLS.Visible = false;
                BLXLS.Visible = true;
                BLXLS.WindowState = XlWindowState.xlMaximized;

                // open a Workbook and get the active Worksheet 
                engWorkbook = engXLS.Workbooks.Open(idDBPath);
                provWorkbook = engXLS.Workbooks.Open(provPath);

                // initialize the new baseline xls
                blWorkbook = BLXLS.Workbooks.Add();
            }
            catch
            {
                throw;
            }
            return (engWorkbook, provWorkbook, blWorkbook);
        }


        internal static Worksheet InitializeTab(string VirtualIem, string ClassStatus)
        {
            Worksheet BLInterimSheet = null;

            string tab_names = VirtualIem + " " + ClassStatus;
            string classifier = "";
            int arr_start, arr_end, j = 0;

            try
            {
                BLInterimSheet = Global.BLXls.Worksheets[tab_names];
            }
            catch
            {
                if (Global.BLXls.Sheets.Count == 1)
                {
                    Global.BLXls.Worksheets[1].Name = tab_names;
                    BLInterimSheet = Global.BLXls.Worksheets[tab_names];
                }
                else
                {
                    BLInterimSheet = Global.BLXls.Worksheets.Add(After: Global.BLXls.Sheets[Global.BLXls.Sheets.Count]);
                    BLInterimSheet.Name = tab_names;
                }   
            }

            switch (VirtualIem)
            {
                case "Network":
                    classifier = "GEO";
                    arr_start = 0;
                    arr_end = 1;
                    break;
                case "Zone":
                    classifier = "GEO";
                    arr_start = 2;
                    arr_end = 3;
                    break;
                case "Cluster":
                    classifier = "GEO";
                    arr_start = 2;
                    arr_end = 5;
                    break;
                case "Segment":
                    classifier = "GEO";
                    arr_start = 2;
                    arr_end = 7;
                    break;
                case "ID1":
                    classifier = "Func";
                    arr_start = 0;
                    arr_end = 1;
                    break;
                case "ID2":
                    classifier = "Func";
                    arr_start = 0;
                    arr_end = 3;
                    break;
                case "ID3":
                    classifier = "Func";
                    arr_start = 0;
                    arr_end = 5;
                    break;
                case "ID4":
                    classifier = "Func";
                    arr_start = 0;
                    arr_end = 7;
                    break;
                case "ID5":
                    classifier = "Func";
                    arr_start = 0;
                    arr_end = 9;
                    break;
                case "ID6":
                    classifier = "Func";
                    arr_start = 0;
                    arr_end = 11;
                    break;
                case "Information":
                    classifier = "Info";
                    arr_start = 0;
                    arr_end = 3;
                    break;
                default:
                    arr_start = 0;
                    arr_end = 0;
                    break;
            }

            switch (classifier)
            {
                case "GEO":
                    for (int i1 = arr_start; i1 <= arr_end; i1++)
                    {
                        BLInterimSheet.Cells[1, j + 1].Value = BlHeaderNetwork[i1];
                        BLInterimSheet.Cells[1, j + 1].Font.Bold = true;
                        BLInterimSheet.Cells[1, j + 1].Interior.ColorIndex = 33;
                        BLInterimSheet.Cells[1, j + 1].Font.ColorIndex = 10;
                        BLInterimSheet.Cells[1, j + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        j += 1;
                    }
                    break;
                case "Func":
                    for (int i1 = arr_start; i1 <= arr_end; i1++)
                    {
                        BLInterimSheet.Cells[1, j + 1].Value = BlHeaderFunction[i1];
                        BLInterimSheet.Cells[1, j + 1].Font.Bold = true;
                        BLInterimSheet.Cells[1, j + 1].Interior.ColorIndex = 33;
                        BLInterimSheet.Cells[1, j + 1].Font.ColorIndex = 10;
                        BLInterimSheet.Cells[1, j + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        j += 1;
                    }
                    break;
                case "Info":
                    for (int i1 = arr_start; i1 <= arr_end; i1++)
                    {
                        BLInterimSheet.Cells[1, j + 1].Value = BlHeaderInformation[i1];
                        BLInterimSheet.Cells[1, j + 1].Font.Bold = true;
                        BLInterimSheet.Cells[1, j + 1].Interior.ColorIndex = 33;
                        BLInterimSheet.Cells[1, j + 1].Font.ColorIndex = 10;
                        BLInterimSheet.Cells[1, j + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        j += 1;
                    }
                    break;
            }
            return BLInterimSheet;

            /*
            for (int i = 0; i < Global.TabNames.Length; i++)
            {
                string s = Global.TabNames[i];
                // Naming sheets
                Worksheet nSheet = blWorkbook.Worksheets[Array.IndexOf(Global.TabNames, s) + 1];
                nSheet.Name = s;

                // Populate col names
                for (int i1 = 0; i1 < BlHeader.Length; i1++)
                {
                    string st = BlHeader[i1];
                    nSheet.Cells[1, Array.IndexOf(BlHeader, st) + 1] = st;
                }
                // Header formatting
                Range rng1 = nSheet.get_Range("A1", "E1");
                rng1.Font.Bold = true;
                rng1.Interior.ColorIndex = 33;
                rng1.Font.ColorIndex = 3;
                rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                blWorkbook.Worksheets.Add(After: blWorkbook.Sheets[blWorkbook.Sheets.Count]);
            }
            */
        }

/*
        internal static Range XlsFind(string sString, Range searchRng)
        {
            Range result = searchRng.Find(
                What: sString,
                LookIn: XlFindLookIn.xlValues,
                LookAt: XlLookAt.xlPart,
                SearchOrder: XlSearchOrder.xlByRows,
                SearchDirection: XlSearchDirection.xlNext
            );
            return result;
        }


        private static bool Starts(string search_string, string target_String)
        {
            return search_string.StartsWith(target_String);
        }


        internal static void GetChildren(string ParentCode, string searchCol, 
            int engChildrenCol, int provChildrenCol, char controller, char modifier)
        {
            string childrenSet = "";
            string newchild, first_addrss, this_addrss;
            int xlsLastLine;
            switch (controller)
            {
                case 'E':
                    tmpRange = engSheet.UsedRange.Columns[searchCol];
                    xlsLastLine = tmpRange.Rows.Count;
                    Range eng_result = XlsFind(ParentCode, tmpRange);
                    if (eng_result == null)
                    {
                        break;
                    }
                    newchild = engSheet.Cells[eng_result.Row, engChildrenCol].Value;
                    if (((modifier == 'Z') & (Starts(newchild, Char.ToString('C')))) |
                        ((modifier == 'C') & (Starts(newchild, Char.ToString('S')))) |
                        ((Starts(search_value, "V") & (Starts(newchild, Char.ToString('W'))))) |
                        ((Starts(search_value, "W") & (Starts(newchild, Char.ToString('U'))))))
                    {
                        childrenSet = childrenSet + newchild;
                        eng_children.Add(item: newchild);
                    }    
                    first_addrss = eng_result.Address;
                    this_addrss = first_addrss;
                    Range next_eng_result = eng_result;
                    do
                    {
                        next_eng_result = tmpRange.FindNext(next_eng_result);
                        this_addrss = next_eng_result.Address;
                        newchild = engSheet.Cells[next_eng_result.Row, engChildrenCol].Value;
                        if (((modifier == 'Z') & (Starts(newchild, Char.ToString('C')))) |
                            ((modifier == 'C') & (Starts(newchild, Char.ToString('S')))) |
                            ((Starts(search_value, "V") & (Starts(newchild, Char.ToString('W'))))) |
                            ((Starts(search_value, "W") & (Starts(newchild, Char.ToString('U'))))))
                        {
                            if (!childrenSet.Contains(newchild))
                            {
                                childrenSet = childrenSet + ", " + newchild;
                                eng_children.Add(item: newchild);
                            }
                        }
                    } while (this_addrss != first_addrss);
                    break;

                case 'P':
                    if (modifier == 'Z')
                    {
                        tmpRange = intrmSheet.UsedRange.Columns[searchCol];
                    }
                    else
                    {
                        tmpRange = provSheet.UsedRange.Columns[searchCol];
                    }
                    xlsLastLine = tmpRange.Rows.Count;
                    Range prov_result = XlsFind(ParentCode, tmpRange);
                    Worksheet interiSheet2 = Global.Provwrkbk.Worksheets["Virtual Item Members"];
                    if (prov_result == null)
                    {
                        break;
                    }
                    newchild = interiSheet2.Cells[prov_result.Row, provChildrenCol].Value;

                    if (((modifier == 'Z') & (Starts(newchild, Char.ToString('C')))) |
                        ((modifier == 'C') & (Starts(newchild, Char.ToString('S')))) |
                        ((Starts(search_value, "V") & (Starts(newchild, Char.ToString('W'))))) |
                        ((Starts(search_value, "W") & (Starts(newchild, Char.ToString('U'))))))
                    {
                        childrenSet = childrenSet + newchild;
                        prov_children.Add(item: newchild);
                    }

                    first_addrss = prov_result.Address;
                    this_addrss = first_addrss;
                    Range next_prov_result = prov_result;

                    do
                    {
                        next_prov_result = tmpRange.FindNext(next_prov_result);
                        this_addrss = next_prov_result.Address;
                        newchild = interiSheet2.Cells[next_prov_result.Row, provChildrenCol].Value;
                        if (((modifier == 'Z') & (Starts(newchild, Char.ToString('C')))) |
                            ((modifier == 'C') & (Starts(newchild, Char.ToString('S')))) |
                            ((Starts(search_value, "V") & (Starts(newchild, Char.ToString('W'))))) |
                            ((Starts(search_value, "W") & (Starts(newchild, Char.ToString('U'))))))
                        {
                            if (!childrenSet.Contains(newchild))
                            {
                                childrenSet = childrenSet + ", " + newchild;
                                prov_children.Add(item: newchild);
                            }
                        }
                    } while (this_addrss != first_addrss);
                    break;
            }
        }


        internal static Boolean DetectDescriptionChange(int s_rowvalue, int r_rowvalue, int eng_col, int prov_col)
        {
            string eng_description = engSheet.Cells[s_rowvalue, eng_col].Value;
            string prov_description = provSheet.Cells[r_rowvalue, prov_col].Value;

            if(prov_description == eng_description)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        internal static Boolean EqualHierarchy(List<string> hierarchy1, List<string> hierarchy2)
        {
            bool equal = hierarchy1.All(x => hierarchy2.Contains(x) &&
                          hierarchy1.Count(a => a == x) == hierarchy2.Count(b => b == x));

            return equal ? true : false;
        }


        private static void ExcuteChanges(int BLastLine, Range rngCol, string prov_search_col, string eng_search_col, 
            int eng_dscr_col, int prov_dscr_col, int eng_children_col, int prov_children_col, char modifier)
        {
            int result_row, eng_search_row;
            provColRange = provSheet.UsedRange.Columns[prov_search_col];
            engColRange = engSheet.UsedRange.Columns[eng_search_col];
            blColRange = blSheet.UsedRange.Columns[1];

            foreach (Range r in rngCol.Cells)
            {
                search_value = r.Value;
                Debug.WriteLine("EngSearchValue = " + search_value);
                if ((Starts(search_value, "Zone")) | (Starts(search_value, "Cluster")) |
                    (Starts(search_value, "Segment")) | ((modifier == 'S') & 
                    !((Starts(r.Value, char.ToString('U'))) |
                    (Starts(r.Value, char.ToString('S'))))) | ((modifier == 'C') &
                    !((Starts(r.Value, char.ToString('W'))) |
                    (Starts(r.Value, char.ToString('C'))))) | ((modifier == 'Z') &
                    !((Starts(r.Value, char.ToString('V'))) |
                    (Starts(r.Value, char.ToString('Z'))))))
                {
                    continue;
                }
                if (search_value == prev_search_value)
                {
                    continue;
                }

                eng_search_row = r.Row;

                Range resultRange = XlsFind(search_value, provColRange);

                result_row = resultRange != null ? resultRange.Row : 1;

                // Cleanup of children sets
                prov_children.Clear();
                eng_children.Clear();

                /*
                // Get hierarchical dependencies
                GetChildren(r.Value, searchCol: eng_search_col, engChildrenCol: eng_children_col,
                    provChildrenCol: prov_children_col, controller: 'E', modifier: modifier);
                eng_children_set = String.Join(", ", eng_children.ToArray());
                Debug.WriteLine("EngChildrenSet = " + eng_children_set);

                if ((modifier != 'S') & (modifier != 'C') & (modifier != 'Z'))
                {
                    GetChildren(r.Value, searchCol: prov_search_col, engChildrenCol: eng_children_col,
                        provChildrenCol: prov_children_col, controller: 'P', modifier: modifier);

                    prov_children_set = String.Join(", ", prov_children.ToArray());
                    Debug.WriteLine("ProvChildrenSet = " + prov_children_set);
                }

                if (modifier == 'Z')
                {
                    intrmSheet = Global.Provwrkbk.Worksheets["Virtual Item Members"];

                    GetChildren(r.Value, searchCol: "A:A", engChildrenCol: eng_children_col,
                        provChildrenCol: prov_children_col, controller: 'P', modifier: modifier);

                    prov_children_set = String.Join(", ", prov_children.ToArray());
                    Debug.WriteLine("ProvChildrenSet = " + prov_children_set);
                }

                if (modifier == 'C')
                {
                    GetChildren(r.Value, searchCol: "G:G", engChildrenCol: eng_children_col,
                        provChildrenCol: prov_children_col, controller: 'P', modifier: modifier);

                    prov_children_set = String.Join(", ", prov_children.ToArray());
                    Debug.WriteLine("ProvChildrenSet = " + prov_children_set);
                }
                */
/*
                switch (resultRange)
                {
                    case null:
                        if (eng_search_row != 1)
                        {
                            if (Starts(r.Value, "N"))
                            {
                                Worksheet temp = Global.BLXls.Worksheets["Network"];
                                temp.Activate();
                                int TempLine = temp.UsedRange.Rows.Count;
                                blSheet.Cells[TempLine + 1, 1].Value = r.Value;
                                blSheet.Cells[TempLine + 1, 2].Value = "N/A";
                                blSheet.Cells[TempLine + 1, 3].Value = engSheet.Cells[eng_search_row, eng_dscr_col];
                                blSheet.Cells[TempLine + 1, 4].Value = BlStatus[1];
                            }
                            else
                            {
                                blSheet.Activate();
                                blSheet.Cells[BLastLine + 1, 1].Value = search_value;
                                blSheet.Cells[BLastLine + 1, 2].Value = "N/A";
                                blSheet.Cells[BLastLine + 1, 3].Value = engSheet.Cells[r.Row, eng_dscr_col];
                                blSheet.Cells[BLastLine + 1, 4].Value = BlStatus[1];
                                if (modifier != 'S')
                                {
                                    blSheet.Cells[BLastLine + 1, 5].Value = eng_children_set;
                                }
                                BLastLine += 1;
                            }
                            prev_search_value = search_value;
                            break;
                        }
                        else
                        {
                            break;
                        }

                    default:
                        {
                            // Detect description changes 
                            if ((DetectDescriptionChange(r.Row, result_row, eng_col: eng_dscr_col, prov_col: prov_dscr_col)) &&
                                (r.Value != blSheet.Cells[BLastLine, 1].Value))
                            {
                                if (Starts(r.Value, "N"))
                                {
                                    Worksheet temp = Global.BLXls.Worksheets["Network"];
                                    temp.Activate();
                                    int TempLine = temp.UsedRange.Rows.Count;
                                    blSheet.Cells[TempLine + 1, 1].Value = r.Value;
                                    blSheet.Cells[TempLine + 1, 2].Value = provSheet.Cells[result_row, prov_dscr_col].Value;
                                    blSheet.Cells[TempLine + 1, 3].Value = engSheet.Cells[eng_search_row, eng_dscr_col];
                                    blSheet.Cells[TempLine + 1, 4].Value = BlStatus[2];
                                }
                                else
                                {
                                    blSheet.Activate();
                                    blSheet.Cells[BLastLine + 1, 1].Value = r.Value;
                                    blSheet.Cells[BLastLine + 1, 2].Value = provSheet.Cells[result_row, prov_dscr_col].Value;
                                    blSheet.Cells[BLastLine + 1, 3].Value = engSheet.Cells[eng_search_row, eng_dscr_col];
                                    blSheet.Cells[BLastLine + 1, 4].Value = BlStatus[2];
                                    if (modifier != 'S')
                                    {
                                        blSheet.Cells[BLastLine + 1, 5].Value = eng_children_set;
                                        blSheet.Cells[BLastLine + 1, 7].Value = prov_children_set;
                                    }
                                }
                                BLastLine += 1;
                            }
                           
                            // Detect identificaiton data hierarchical changes
                            if (!EqualHierarchy(prov_children, eng_children))
                            {
                                blSheet.Activate();
                                blSheet.Cells[BLastLine + 1, 1].Value = r.Value;
                                blSheet.Cells[BLastLine + 1, 2].Value = provSheet.Cells[result_row, prov_dscr_col].Value;
                                blSheet.Cells[BLastLine + 1, 3].Value = engSheet.Cells[eng_search_row, eng_dscr_col];
                                blSheet.Cells[BLastLine + 1, 4].Value = BlStatus[3];
                                if (modifier != 'S')
                                {
                                    blSheet.Cells[BLastLine + 1, 5].Value = eng_children_set;
                                    blSheet.Cells[BLastLine + 1, 6].Value = "Unmatched";
                                    blSheet.Cells[BLastLine + 1, 7].Value = prov_children_set;
                                }
                                BLastLine += 1;
                            }

                            prev_search_value = search_value;
                            break;
                        }
                }
            }
        }


        internal static void XlsCompare(string criterion)
        {
            Range engRngCol;
            switch (criterion)
            {
                case "Network":

                    break;

                case "Zone":
                    engSheet = Global.Engwrkbk.Worksheets["Network"];
                    provSheet = Global.Provwrkbk.Worksheets["Virtual Item Groups"];
                    blSheet = Global.BLXls.Worksheets[criterion];
                    BLastLine = blSheet.UsedRange.Rows.Count;
                    engRngCol = engSheet.UsedRange.Columns["A:A"];

                    ExcuteChanges(BLastLine, engRngCol, prov_search_col: "B:B", eng_dscr_col: 2, 
                        prov_dscr_col: 5, eng_children_col: 3, eng_search_col: "A:A", 
                        prov_children_col: 5, modifier: 'Z');

                    break;

                case "Cluster":
                    engSheet = Global.Engwrkbk.Worksheets["Network"];
                    provSheet = Global.Provwrkbk.Worksheets["Virtual Item Members"];
                    blSheet = Global.BLXls.Worksheets[criterion];
                    BLastLine = blSheet.UsedRange.Rows.Count;
                    engRngCol = engSheet.UsedRange.Columns["C:C"];

                    ExcuteChanges(BLastLine, engRngCol, prov_search_col: "E:E", eng_dscr_col: 4,
                        prov_dscr_col: 6, eng_children_col: 5, eng_search_col: "C:C",
                        prov_children_col: 5, modifier: 'C');

                    break;

                case "Segment":
                    engSheet = Global.Engwrkbk.Worksheets["Network"];
                    provSheet = Global.Provwrkbk.Worksheets["Virtual Item Members"];
                    blSheet = Global.BLXls.Worksheets[criterion];
                    BLastLine = blSheet.UsedRange.Rows.Count;
                    engRngCol = engSheet.UsedRange.Columns["E:E"];

                    ExcuteChanges(BLastLine, engRngCol, prov_search_col: "E:E", eng_dscr_col: 6,
                        prov_dscr_col: 6, eng_children_col: 5, eng_search_col: "E:E",
                        prov_children_col: 5, modifier: 'S');

                    break;

                case "ID1":
                    engSheet = Global.Engwrkbk.Worksheets["Engineering breakdown"];
                    provSheet = Global.Provwrkbk.Worksheets["Virtual Item Groups"];
                    blSheet = Global.BLXls.Worksheets[criterion];
                    BLastLine = blSheet.UsedRange.Rows.Count;
                    engRngCol = engSheet.UsedRange.Columns["B:B"];

                    ExcuteChanges(BLastLine, engRngCol, prov_search_col: "B:B", eng_dscr_col: 1,
                        prov_dscr_col: 5, eng_children_col: 4, eng_search_col: "B:B",
                        prov_children_col: 5, modifier: 'S');
                    break;

                case "ID2":
                    Console.WriteLine("Case 2");
                    break;

                case "ID3":
                    Console.WriteLine("Case 2");
                    break;

                case "ID4":
                    Console.WriteLine("Case 2");
                    break;

                case "ID5":
                    Console.WriteLine("Case 2");
                    break;

                case "ID6":
                    Console.WriteLine("Case 2");
                    break;

                default:
                    Console.WriteLine("Default case");
                    break;
            }
        }


        internal static void XlsSaveClose(Application newXls, params Excel.Workbook[] wrkBooks)
        {
            foreach (Excel.Workbook wrkBook in wrkBooks)
            {
                // Save WorkBook and close 
                wrkBook.Save();
                wrkBook.Close();
            }

            // Quit Excel Application 
            engXLS.Quit();
            newXls.Quit();

            //xlWorkBook.Close(false);

            //releaseObject(xlWorkSheet);

            //releaseObject(xlWorkBook);

            //releaseObject(xlApp);
        }
    }
}
