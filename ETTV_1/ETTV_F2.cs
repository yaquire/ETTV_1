using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.ApplicationServices;
using System.Collections;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Aspose.Cells.Charts;
using System.Data.OleDb;
using Spire.Xls;
using OfficeOpenXml; // Add this using directive for EPPlus
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace ETTV_1
{
    public partial class ETTV_F2 : System.Windows.Forms.Form
    {
        public UIApplication uiapp2;
        public UIDocument uidoc2;
        public Autodesk.Revit.ApplicationServices.Application app;
        public Document doc2;
        public string test;              

        private int spCount;
        private int num1;
        private int num2;
        private int num22;
        private int num4;
        private int num5;
        private double num3;
        private double TotalArea;
        private double CalValue;
        private double TotalCalValue;
        private double CalValue1;
        private double TotalCalValue1;
        private int num10;
        private int num20;
        private int num111;
        private double TotalWallArea;
        private double TotalWndwArea;
        private double N_Wall_HG;
        private double S_Wall_HG;
        private double E_Wall_HG;
        private double W_Wall_HG;
        private double NE_Wall_HG;
        private double NW_Wall_HG;
        private double SE_Wall_HG;
        private double SW_Wall_HG;

        private List<string> Lst00 = new List<string>();
        private List<string> Lst33 = new List<string>();
        private List<string> Lst44 = new List<string>();

        private List<string> WndwTypeLst1_N = new List<string>();
        private List<string> WndwTypeLst_N = new List<string>();
        private List<string> WndwAreaLst_N = new List<string>();
        private List<double> WndwArea_DouLst_N = new List<double>();
        private List<string> WndwUValueLst_N = new List<string>();
        private List<string> WndwSC1ValueLst_N = new List<string>();
        
        private List<string> WallTypeLst1_N = new List<string>();
        private List<string> WallTypeLst_N = new List<string>();
        private List<string> WallAreaLst_N = new List<string>();
        private List<double> WallArea_DouLst_N = new List<double>();
        private List<string> WallUValueLst_N = new List<string>();



        private List<string> WndwTypeLst1_S = new List<string>();
        private List<string> WndwTypeLst_S = new List<string>();
        private List<string> WndwAreaLst_S = new List<string>();
        private List<double> WndwArea_DouLst_S = new List<double>();
        private List<string> WndwUValueLst_S = new List<string>();
        private List<string> WndwSC1ValueLst_S = new List<string>();

        private List<string> WallTypeLst1_S = new List<string>();
        private List<string> WallTypeLst_S = new List<string>();
        private List<string> WallAreaLst_S = new List<string>();
        private List<double> WallArea_DouLst_S = new List<double>();
        private List<string> WallUValueLst_S = new List<string>();



        private List<string> WndwTypeLst1_E = new List<string>();
        private List<string> WndwTypeLst_E = new List<string>();
        private List<string> WndwAreaLst_E = new List<string>();
        private List<double> WndwArea_DouLst_E = new List<double>();
        private List<string> WndwUValueLst_E = new List<string>();
        private List<string> WndwSC1ValueLst_E = new List<string>();

        private List<string> WallTypeLst1_E = new List<string>();
        private List<string> WallTypeLst_E = new List<string>();
        private List<string> WallAreaLst_E = new List<string>();
        private List<double> WallArea_DouLst_E = new List<double>();
        private List<string> WallUValueLst_E = new List<string>();



        private List<string> WndwTypeLst1_W = new List<string>();
        private List<string> WndwTypeLst_W = new List<string>();
        private List<string> WndwAreaLst_W = new List<string>();
        private List<double> WndwArea_DouLst_W = new List<double>();
        private List<string> WndwUValueLst_W = new List<string>();
        private List<string> WndwSC1ValueLst_W = new List<string>();

        private List<string> WallTypeLst1_W = new List<string>();
        private List<string> WallTypeLst_W = new List<string>();
        private List<string> WallAreaLst_W = new List<string>();
        private List<double> WallArea_DouLst_W = new List<double>();
        private List<string> WallUValueLst_W = new List<string>();



        private List<string> WndwTypeLst1_NE = new List<string>();
        private List<string> WndwTypeLst_NE = new List<string>();
        private List<string> WndwAreaLst_NE = new List<string>();
        private List<double> WndwArea_DouLst_NE = new List<double>();
        private List<string> WndwUValueLst_NE = new List<string>();
        private List<string> WndwSC1ValueLst_NE = new List<string>();

        private List<string> WallTypeLst1_NE = new List<string>();
        private List<string> WallTypeLst_NE = new List<string>();
        private List<string> WallAreaLst_NE = new List<string>();
        private List<double> WallArea_DouLst_NE = new List<double>();
        private List<string> WallUValueLst_NE = new List<string>();



        private List<string> WndwTypeLst1_NW = new List<string>();
        private List<string> WndwTypeLst_NW = new List<string>();
        private List<string> WndwAreaLst_NW = new List<string>();
        private List<double> WndwArea_DouLst_NW = new List<double>();
        private List<string> WndwUValueLst_NW = new List<string>();
        private List<string> WndwSC1ValueLst_NW = new List<string>();

        private List<string> WallTypeLst1_NW = new List<string>();
        private List<string> WallTypeLst_NW = new List<string>();
        private List<string> WallAreaLst_NW = new List<string>();
        private List<double> WallArea_DouLst_NW = new List<double>();
        private List<string> WallUValueLst_NW = new List<string>();



        private List<string> WndwTypeLst1_SE = new List<string>();
        private List<string> WndwTypeLst_SE = new List<string>();
        private List<string> WndwAreaLst_SE = new List<string>();
        private List<double> WndwArea_DouLst_SE = new List<double>();
        private List<string> WndwUValueLst_SE = new List<string>();
        private List<string> WndwSC1ValueLst_SE = new List<string>();

        private List<string> WallTypeLst1_SE = new List<string>();
        private List<string> WallTypeLst_SE = new List<string>();
        private List<string> WallAreaLst_SE = new List<string>();
        private List<double> WallArea_DouLst_SE = new List<double>();
        private List<string> WallUValueLst_SE = new List<string>();



        private List<string> WndwTypeLst1_SW = new List<string>();
        private List<string> WndwTypeLst_SW = new List<string>();
        private List<string> WndwAreaLst_SW = new List<string>();
        private List<double> WndwArea_DouLst_SW = new List<double>();
        private List<string> WndwUValueLst_SW = new List<string>();
        private List<string> WndwSC1ValueLst_SW = new List<string>();

        private List<string> WallTypeLst1_SW = new List<string>();
        private List<string> WallTypeLst_SW = new List<string>();
        private List<string> WallAreaLst_SW = new List<string>();
        private List<double> WallArea_DouLst_SW = new List<double>();
        private List<string> WallUValueLst_SW = new List<string>();


        /////////////////        
        private List<string> WndwTypeLst0 = new List<string>();
        private List<string> WndwTypeLst1 = new List<string>();
        private List<string> WndwTypeLst = new List<string>();        
        private List<string> WndwUValueLst = new List<string>();
        private List<string> WndwSC1ValueLst = new List<string>();
        private List<string> WndwUValueLst0 = new List<string>();
        private List<string> WndwSC1ValueLst0 = new List<string>();
        private List<string> WndwTypeLst_FC = new List<string>();
        private List<string> WndwSC1ValueLst_2 = new List<string>();
        ///////////////
        /////////////////        
        private List<string> WallTypeLst0 = new List<string>();
        private List<string> WallTypeLst1 = new List<string>();
        private List<string> WallTypeLst = new List<string>();
        private List<string> WallUValueLst = new List<string>();        
        private List<string> WallUValueLst0 = new List<string>();
        ///////////////

        private List<ElementId> WallTypeId_F2 = new List<ElementId>();
        private List<ElementType> WallTypeLst_F2 = new List<ElementType>();
        private List<Wall> WallLst_F2 = new List<Wall>();

        private List<Element> WndwLst_F2 = new List<Element>();
        private List<ElementType> WndwTypeLst_F2 = new List<ElementType>();
        private List<ElementId> WndwTypeId_F2 = new List<ElementId>();

        const double _inchToMm = 25.4;
        const double _footToMm = 12 * _inchToMm;

        private double SC2_NS;
        private double SC2_EW;
        private double SC2_NENW;
        private double SC2_SESW;

        private double SC2_VP_NS;
        private double SC2_VP_EW;
        private double SC2_VP_NENW;
        private double SC2_VP_SESW;

        private double SC2_EC_NS_b1;
        private double SC2_EC_NS_b2;
        private double SC2_EC_NS;
        private double SC2_EC_EW_b1;
        private double SC2_EC_EW_b2;
        private double SC2_EC_EW;
        private double SC2_EC_NENW_b1;
        private double SC2_EC_NENW_b2;
        private double SC2_EC_NENW;
        private double SC2_EC_SESW_b1;
        private double SC2_EC_SESW_b2;
        private double SC2_EC_SESW;

        private double Total_HG;
        private double Total_Area;
        private Excel.Workbook workbook = null;
        private Excel.Application excelApp = null;

        private double EC_R1;
        private double EC_R2;
        private double EC_b1;
        private double EC_b2;

        //private List<double> Lst_SC2_HP_NS = new List<double>();

        //private ViewSchedule shdl;
        //private IList<SchedulableField> shdlableFields;
        //private List<ScheduleFieldId> shdlFieldIds = new List<ScheduleFieldId>();

        #region Progress

        
        #endregion


        public ETTV_F2(ExternalCommandData commandData)
        {
            InitializeComponent();
            
            uiapp2 = commandData.Application;
            uidoc2 = uiapp2.ActiveUIDocument;
            app = uiapp2.Application;
            doc2 = uidoc2.Document;

            //WallTypeLst_F2 = ettv1.MyList;


            // Select Space Based on Ventilation==> AC 
            FilteredElementCollector spCollector = new FilteredElementCollector(doc2);
            ElementCategoryFilter spFilter = new ElementCategoryFilter(BuiltInCategory.OST_Rooms);
            IList<Element> spList = spCollector.WherePasses(spFilter).WhereElementIsNotElementType().ToElements();
            spCount = spList.Count;

            IList<Element> ACSPList = spList.Where(SP => SP.LookupParameter("AC_Space")
                        .AsInteger() == 1).Cast<Element>().ToList();


            /////////////////////Window_North\\\\\\\\\\\\\\\\\\\\\\\\\   

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_N = new List<string>();
            WndwTypeLst_N = new List<string>();
            WndwAreaLst_N = new List<string>();
            WndwArea_DouLst_N = new List<double>();
            WndwUValueLst_N = new List<string>();
            WndwSC1ValueLst_N = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_N");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_N");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_N");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_N");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_N.Add(st);
                        WndwTypeLst_N.Add(Lst1[num1]);                        
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_N = WndwTypeLst1_N.Distinct().ToList();
            WndwTypeLst_N = WndwTypeLst_N.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_N)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_N.Add(Lst33[num2]);
                        WndwSC1ValueLst_N.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }
                                  

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_N");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_N");
                string input2 = P2.AsString();
               

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                

                num2 = 0;
                foreach (string st in WndwTypeLst1_N)
                {
                    WndwArea_DouLst_N.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_N)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_N[num2] = WndwArea_DouLst_N[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_N) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_N.Add(Convert.ToString(db));
                TotalArea = TotalArea + db; 
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_N) //write into DataGridView
            {
                dataGridView2.Rows.Add(st,WndwTypeLst_N[num1],WndwAreaLst_N[num1], WndwUValueLst_N[num1], WndwSC1ValueLst_N[num1]);
                num1 = num1 + 1;
            }
            dataGridView2.Rows.Add("", "Subtotal", TotalArea);
            TotalWndwArea = TotalArea; //Total Window Area

            foreach (Room rm in ACSPList)
            {
                dataGridView1.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_N");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_N");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_N");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_N");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_N");
                string input4 = P4.AsString();


                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGridView1.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }



            }//For Detail View

            ////////// For Window Types Page //////////////////////
            foreach (string st in WndwTypeLst1_N)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_N)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_N)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_N)
            {
                WndwSC1ValueLst0.Add(st);
            }
            ////////// For Window Types Page //////////////////////

            /////////////////////Window_North\\\\\\\\\\\\\\\\\\\\\\\\\         

            /////////////////////Wall_North\\\\\\\\\\\\\\\\\\\\\\\\\     
            Lst00 = new List<string>();
            Lst33 = new List<string>();            
            WallTypeLst1_N = new List<string>();
            WallTypeLst_N = new List<string>();
            WallAreaLst_N = new List<string>();
            WallArea_DouLst_N = new List<double>();
            WallUValueLst_N = new List<string>();
            
            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_N");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_N");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_N");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_N.Add(st);
                        WallTypeLst_N.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_N = WallTypeLst1_N.Distinct().ToList();
            WallTypeLst_N = WallTypeLst_N.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_N)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_N.Add(Lst33[num2]);                        
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_N");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_N");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_N)
                {
                   WallArea_DouLst_N.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_N)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_N[num2] = WallArea_DouLst_N[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_N) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_N.Add(Convert.ToString(db));
                TotalArea = TotalArea + db; 
            }

            num1 = 0;
            CalValue=0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_N) //write into DataGridView
            {
                CalValue = 12* Convert.ToDouble(WallAreaLst_N[num1]) * Convert.ToDouble(WallUValueLst_N[num1]);
                dataGrid_Wall_N_S.Rows.Add(st, WallTypeLst_N[num1], WallAreaLst_N[num1], WallUValueLst_N[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_N_S.Rows.Add("", "Subtotal", TotalArea,"", (TotalCalValue.ToString("0.00")));
            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_N.Text = (TotalWndwArea + TotalWallArea).ToString();
            N_Wall_HG = TotalCalValue;

            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_N_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_N");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_N");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_N");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_N");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                CalValue = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_N_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }



            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_N)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_N)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_N)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_North\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_South\\\\\\\\\\\\\\\\\\\\\\\\\ 

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_S = new List<string>();
            WndwTypeLst_S = new List<string>();
            WndwAreaLst_S = new List<string>();
            WndwArea_DouLst_S = new List<double>();
            WndwUValueLst_S = new List<string>();
            WndwSC1ValueLst_S = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_S");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_S");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_S");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_S");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_S.Add(st);
                        WndwTypeLst_S.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_S = WndwTypeLst1_S.Distinct().ToList();
            WndwTypeLst_S = WndwTypeLst_S.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_S)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_S.Add(Lst33[num2]);
                        WndwSC1ValueLst_S.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }


            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_S");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_S");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_S)
                {
                    WndwArea_DouLst_S.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_S)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_S[num2] = WndwArea_DouLst_S[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_S) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_S.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_S) //write into DataGridView
            {
                dataGridView3.Rows.Add(st, WndwTypeLst_S[num1], WndwAreaLst_S[num1], WndwUValueLst_S[num1], WndwSC1ValueLst_S[num1]);
                num1 = num1 + 1;
            }
            dataGridView3.Rows.Add("", "Subtotal", TotalArea);
            TotalWndwArea = TotalArea; //Total Wndw Area

            foreach (Room rm in ACSPList)
            {
                dataGridView4.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_S");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_S");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_S");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_S");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_S");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGridView4.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }



            }//For Detail View

            foreach (string st in WndwTypeLst1_S)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_S)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_S)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_S)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_South\\\\\\\\\\\\\\\\\\\\\\\\\ 

            /////////////////////Wall_South\\\\\\\\\\\\\\\\\\\\\\\\\   

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_S = new List<string>();
            WallTypeLst_S = new List<string>();
            WallAreaLst_S = new List<string>();
            WallArea_DouLst_S = new List<double>();
            WallUValueLst_S = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_S");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_S");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_S");
                string input3 = P3.AsString();                      

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_S.Add(st);
                        WallTypeLst_S.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_S = WallTypeLst1_S.Distinct().ToList();
            WallTypeLst_S = WallTypeLst_S.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_S)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_S.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }


            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_S");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_S");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_S)
                {
                    WallArea_DouLst_S.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_S)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_S[num2] = WallArea_DouLst_S[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_S) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_S.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_S) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_S[num1]) * Convert.ToDouble(WallUValueLst_S[num1]);
                dataGrid_Wall_S_S.Rows.Add(st, WallTypeLst_S[num1], WallAreaLst_S[num1], WallUValueLst_S[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_S_S.Rows.Add("", "Subtotal", TotalArea,"", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_S.Text = (TotalWndwArea + TotalWallArea).ToString();
            S_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_S_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_S");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_S");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_S");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_S");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_S_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_S)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_S)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_S)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_South\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_East\\\\\\\\\\\\\\\\\\\\\\\\\ 
            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_E = new List<string>();
            WndwTypeLst_E = new List<string>();
            WndwAreaLst_E = new List<string>();
            WndwArea_DouLst_E = new List<double>();
            WndwUValueLst_E = new List<string>();
            WndwSC1ValueLst_E = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_E");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_E");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_E");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_E");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_E.Add(st);
                        WndwTypeLst_E.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_E = WndwTypeLst1_E.Distinct().ToList();
            WndwTypeLst_E = WndwTypeLst_E.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_E)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_E.Add(Lst33[num2]);
                        WndwSC1ValueLst_E.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_E");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_E");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_E)
                {
                    WndwArea_DouLst_E.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_E)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_E[num2] = WndwArea_DouLst_E[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_E) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_E.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_E) //write into DataGridView
            {
                dataGrid_Wndw_E_S.Rows.Add(st, WndwTypeLst_E[num1], WndwAreaLst_E[num1], WndwUValueLst_E[num1], WndwSC1ValueLst_E[num1]);
                num1 = num1 + 1;
            }
            dataGrid_Wndw_E_S.Rows.Add("", "Subtotal", TotalArea);

            TotalWndwArea = TotalArea; //Total Window Area

            foreach (Room rm in ACSPList) //For Detail View
            {
                dataGrid_Wndw_E_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_E");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_E");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_E");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_E");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_E");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGrid_Wndw_E_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }


            }//For Detail View

            foreach (string st in WndwTypeLst1_E)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_E)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_E)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_E)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_East\\\\\\\\\\\\\\\\\\\\\\\\\ 

            /////////////////////Wall_East\\\\\\\\\\\\\\\\\\\\\\\\\   

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_E = new List<string>();
            WallTypeLst_E = new List<string>();
            WallAreaLst_E = new List<string>();
            WallArea_DouLst_E = new List<double>();
            WallUValueLst_E = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_E");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_E");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_E");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_E.Add(st);
                        WallTypeLst_E.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_E = WallTypeLst1_E.Distinct().ToList();
            WallTypeLst_E = WallTypeLst_E.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_E)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_E.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_E");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_E");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_E)
                {
                    WallArea_DouLst_E.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_E)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_E[num2] = WallArea_DouLst_E[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_E) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_E.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_E) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_E[num1]) * Convert.ToDouble(WallUValueLst_E[num1]);
                dataGrid_Wall_E_S.Rows.Add(st, WallTypeLst_E[num1], WallAreaLst_E[num1], WallUValueLst_E[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_E_S.Rows.Add("", "Subtotal", TotalArea, "", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_E.Text = (TotalWndwArea + TotalWallArea).ToString();
            E_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_E_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_E");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_E");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_E");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_E");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_E_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_E)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_E)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_E)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_East\\\\\\\\\\\\\\\\\\\\\\\\\   

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_West\\\\\\\\\\\\\\\\\\\\\\\\\   
            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_W = new List<string>();
            WndwTypeLst_W = new List<string>();
            WndwAreaLst_W = new List<string>();
            WndwArea_DouLst_W = new List<double>();
            WndwUValueLst_W = new List<string>();
            WndwSC1ValueLst_W = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_W");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_W");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_W");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_W");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_W.Add(st);
                        WndwTypeLst_W.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_W = WndwTypeLst1_W.Distinct().ToList();
            WndwTypeLst_W = WndwTypeLst_W.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_W)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_W.Add(Lst33[num2]);
                        WndwSC1ValueLst_W.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_W");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_W");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_W)
                {
                    WndwArea_DouLst_W.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_W)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_W[num2] = WndwArea_DouLst_W[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_W) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_W.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_W) //write into DataGridView
            {
                dataGrid_Wndw_W_S.Rows.Add(st, WndwTypeLst_W[num1], WndwAreaLst_W[num1], WndwUValueLst_W[num1], WndwSC1ValueLst_W[num1]);
                num1 = num1 + 1;
            }
            dataGrid_Wndw_W_S.Rows.Add("", "Subtotal", TotalArea);

            TotalWndwArea = TotalArea; //Total Window Area

            foreach (Room rm in ACSPList) //For Detail View
            {
                dataGrid_Wndw_W_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_W");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_W");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_W");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_W");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_W");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGrid_Wndw_W_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            foreach (string st in WndwTypeLst1_W)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_W)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_W)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_W)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_West\\\\\\\\\\\\\\\\\\\\\\\\\   

            /////////////////////Wall_West\\\\\\\\\\\\\\\\\\\\\\\\\

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_W = new List<string>();
            WallTypeLst_W = new List<string>();
            WallAreaLst_W = new List<string>();
            WallArea_DouLst_W = new List<double>();
            WallUValueLst_W = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_W");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_W");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_W");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_W.Add(st);
                        WallTypeLst_W.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_W = WallTypeLst1_W.Distinct().ToList();
            WallTypeLst_W = WallTypeLst_W.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_W)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_W.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_W");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_W");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_W)
                {
                    WallArea_DouLst_W.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_W)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_W[num2] = WallArea_DouLst_W[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_W) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_W.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_W) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_W[num1]) * Convert.ToDouble(WallUValueLst_W[num1]);
                dataGrid_Wall_W_S.Rows.Add(st, WallTypeLst_W[num1], WallAreaLst_W[num1], WallUValueLst_W[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_W_S.Rows.Add("", "Subtotal", TotalArea, "", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_W.Text = (TotalWndwArea + TotalWallArea).ToString();
            W_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_W_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_W");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_W");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_W");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_W");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_W_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_W)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_W)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_W)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_West\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_North East\\\\\\\\\\\\\\\\\\\\\\\\\ 
            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_NE = new List<string>();
            WndwTypeLst_NE = new List<string>();
            WndwAreaLst_NE = new List<string>();
            WndwArea_DouLst_NE = new List<double>();
            WndwUValueLst_NE = new List<string>();
            WndwSC1ValueLst_NE = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_NE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_NE");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_NE");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_NE");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_NE.Add(st);
                        WndwTypeLst_NE.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_NE = WndwTypeLst1_NE.Distinct().ToList();
            WndwTypeLst_NE = WndwTypeLst_NE.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_NE)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_NE.Add(Lst33[num2]);
                        WndwSC1ValueLst_NE.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_NE");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_NE");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_NE)
                {
                    WndwArea_DouLst_NE.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_NE)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_NE[num2] = WndwArea_DouLst_NE[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_NE) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_NE.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_NE) //write into DataGridView
            {
                dataGrid_Wndw_NE_S.Rows.Add(st, WndwTypeLst_NE[num1], WndwAreaLst_NE[num1], WndwUValueLst_NE[num1], WndwSC1ValueLst_NE[num1]);
                num1 = num1 + 1;
            }
            dataGrid_Wndw_NE_S.Rows.Add("", "Subtotal", TotalArea);

            TotalWndwArea = TotalArea; //Total Window Area

            foreach (Room rm in ACSPList) //For Detail View
            {
                dataGrid_Wndw_NE_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_NE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_NE");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_NE");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_NE");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_NE");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGrid_Wndw_NE_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            foreach (string st in WndwTypeLst1_NE)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_NE)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_NE)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_NE)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_North East\\\\\\\\\\\\\\\\\\\\\\\\\ 

            /////////////////////Wall_North East\\\\\\\\\\\\\\\\\\\\\\\\\

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_NE = new List<string>();
            WallTypeLst_NE = new List<string>();
            WallAreaLst_NE = new List<string>();
            WallArea_DouLst_NE = new List<double>();
            WallUValueLst_NE = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_NE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_NE");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_NE");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_NE.Add(st);
                        WallTypeLst_NE.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_NE = WallTypeLst1_NE.Distinct().ToList();
            WallTypeLst_NE = WallTypeLst_NE.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_NE)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_NE.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_NE");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_NE");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_NE)
                {
                    WallArea_DouLst_NE.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_NE)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_NE[num2] = WallArea_DouLst_NE[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_NE) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_NE.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_NE) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_NE[num1]) * Convert.ToDouble(WallUValueLst_NE[num1]);
                dataGrid_Wall_NE_S.Rows.Add(st, WallTypeLst_NE[num1], WallAreaLst_NE[num1], WallUValueLst_NE[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_NE_S.Rows.Add("", "Subtotal", TotalArea, "", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_NE.Text = (TotalWndwArea + TotalWallArea).ToString();
            NE_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_NE_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_NE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_NE");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_NE");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_NE");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_NE_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_NE)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_NE)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_NE)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////


            /////////////////////Wall_North East\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_North West\\\\\\\\\\\\\\\\\\\\\\\\\ 
            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_NW = new List<string>();
            WndwTypeLst_NW = new List<string>();
            WndwAreaLst_NW = new List<string>();
            WndwArea_DouLst_NW = new List<double>();
            WndwUValueLst_NW = new List<string>();
            WndwSC1ValueLst_NW = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_NW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_NW");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_NW");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_NW");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_NW.Add(st);
                        WndwTypeLst_NW.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_NW = WndwTypeLst1_NW.Distinct().ToList();
            WndwTypeLst_NW = WndwTypeLst_NW.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_NW)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_NW.Add(Lst33[num2]);
                        WndwSC1ValueLst_NW.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_NW");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_NW");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_NW)
                {
                    WndwArea_DouLst_NW.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_NW)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_NW[num2] = WndwArea_DouLst_NW[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_NW) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_NW.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_NW) //write into DataGridView
            {
                dataGrid_Wndw_NW_S.Rows.Add(st, WndwTypeLst_NW[num1], WndwAreaLst_NW[num1], WndwUValueLst_NW[num1], WndwSC1ValueLst_NW[num1]);
                num1 = num1 + 1;
            }
            dataGrid_Wndw_NW_S.Rows.Add("", "Subtotal", TotalArea);

            TotalWndwArea = TotalArea; //Total Window Area

            foreach (Room rm in ACSPList) //For Detail View
            {
                dataGrid_Wndw_NW_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_NW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_NW");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_NW");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_NW");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_NW");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGrid_Wndw_NW_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            foreach (string st in WndwTypeLst1_NW)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_NW)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_NW)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_NW)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_North West\\\\\\\\\\\\\\\\\\\\\\\\\ 

            /////////////////////Wall_North West\\\\\\\\\\\\\\\\\\\\\\\\\

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_NW = new List<string>();
            WallTypeLst_NW = new List<string>();
            WallAreaLst_NW = new List<string>();
            WallArea_DouLst_NW = new List<double>();
            WallUValueLst_NW = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_NW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_NW");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_NW");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_NW.Add(st);
                        WallTypeLst_NW.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_NW = WallTypeLst1_NW.Distinct().ToList();
            WallTypeLst_NW = WallTypeLst_NW.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_NW)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_NW.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_NW");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_NW");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_NW)
                {
                    WallArea_DouLst_NW.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_NW)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_NW[num2] = WallArea_DouLst_NW[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_NW) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_NW.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_NW) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_NW[num1]) * Convert.ToDouble(WallUValueLst_NW[num1]);
                dataGrid_Wall_NW_S.Rows.Add(st, WallTypeLst_NW[num1], WallAreaLst_NW[num1], WallUValueLst_NW[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_NW_S.Rows.Add("", "Subtotal", TotalArea, "", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_NW.Text = (TotalWndwArea + TotalWallArea).ToString();
            NW_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_NW_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_NW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_NW");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_NW");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_NW");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_NW_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_NW)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_NW)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_NW)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_North West\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_South East\\\\\\\\\\\\\\\\\\\\\\\\\ 
            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_SE = new List<string>();
            WndwTypeLst_SE = new List<string>();
            WndwAreaLst_SE = new List<string>();
            WndwArea_DouLst_SE = new List<double>();
            WndwUValueLst_SE = new List<string>();
            WndwSC1ValueLst_SE = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_SE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_SE");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_SE");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_SE");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_SE.Add(st);
                        WndwTypeLst_SE.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_SE = WndwTypeLst1_SE.Distinct().ToList();
            WndwTypeLst_SE = WndwTypeLst_SE.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_SE)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_SE.Add(Lst33[num2]);
                        WndwSC1ValueLst_SE.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_SE");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_SE");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_SE)
                {
                    WndwArea_DouLst_SE.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_SE)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_SE[num2] = WndwArea_DouLst_SE[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_SE) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_SE.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_SE) //write into DataGridView
            {
                dataGrid_Wndw_SE_S.Rows.Add(st, WndwTypeLst_SE[num1], WndwAreaLst_SE[num1], WndwUValueLst_SE[num1], WndwSC1ValueLst_SE[num1]);
                num1 = num1 + 1;
            }
            dataGrid_Wndw_SE_S.Rows.Add("", "Subtotal", TotalArea);

            TotalWndwArea = TotalArea; //Total Window Area

            foreach (Room rm in ACSPList) //For Detail View
            {
                dataGrid_Wndw_SE_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_SE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_SE");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_SE");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_SE");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_SE");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGrid_Wndw_SE_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            foreach (string st in WndwTypeLst1_SE)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_SE)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_SE)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_SE)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_South East\\\\\\\\\\\\\\\\\\\\\\\\\ 

            /////////////////////Wall_South East\\\\\\\\\\\\\\\\\\\\\\\\\

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_SE = new List<string>();
            WallTypeLst_SE = new List<string>();
            WallAreaLst_SE = new List<string>();
            WallArea_DouLst_SE = new List<double>();
            WallUValueLst_SE = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_SE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_SE");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_SE");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_SE.Add(st);
                        WallTypeLst_SE.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_SE = WallTypeLst1_SE.Distinct().ToList();
            WallTypeLst_SE = WallTypeLst_SE.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_SE)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_SE.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_SE");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_SE");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_SE)
                {
                    WallArea_DouLst_SE.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_SE)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_SE[num2] = WallArea_DouLst_SE[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_SE) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_SE.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_SE) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_SE[num1]) * Convert.ToDouble(WallUValueLst_SE[num1]);
                dataGrid_Wall_SE_S.Rows.Add(st, WallTypeLst_SE[num1], WallAreaLst_SE[num1], WallUValueLst_SE[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_SE_S.Rows.Add("", "Subtotal", TotalArea, "", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_SE.Text = (TotalWndwArea + TotalWallArea).ToString();
            SE_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_SE_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_SE");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_SE");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_SE");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_SE");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_SE_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_SE)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_SE)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_SE)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_South East\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            /////////////////////Window_South West\\\\\\\\\\\\\\\\\\\\\\\\\ 
            Lst00 = new List<string>();
            Lst33 = new List<string>();
            Lst44 = new List<string>();
            WndwTypeLst1_SW = new List<string>();
            WndwTypeLst_SW = new List<string>();
            WndwAreaLst_SW = new List<string>();
            WndwArea_DouLst_SW = new List<double>();
            WndwUValueLst_SW = new List<string>();
            WndwSC1ValueLst_SW = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_SW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_SW");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_SW");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_SW");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);
                        Lst44.Add(Lst4[num1]);

                        WndwTypeLst1_SW.Add(st);
                        WndwTypeLst_SW.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WndwTypeLst1_SW = WndwTypeLst1_SW.Distinct().ToList();
            WndwTypeLst_SW = WndwTypeLst_SW.Distinct().ToList();

            num1 = 0;
            foreach (string st in WndwTypeLst1_SW)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WndwUValueLst_SW.Add(Lst33[num2]);
                        WndwSC1ValueLst_SW.Add(Lst44[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WndwTypeLst1_SW");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_SW");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WndwTypeLst1_SW)
                {
                    WndwArea_DouLst_SW.Add(0);
                }

                num2 = 0;
                foreach (string WT in WndwTypeLst1_SW)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WndwArea_DouLst_SW[num2] = WndwArea_DouLst_SW[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WndwArea_DouLst_SW) // to convert Double List to String List & add all areas together
            {
                WndwAreaLst_SW.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1_SW) //write into DataGridView
            {
                dataGrid_Wndw_SW_S.Rows.Add(st, WndwTypeLst_SW[num1], WndwAreaLst_SW[num1], WndwUValueLst_SW[num1], WndwSC1ValueLst_SW[num1]);
                num1 = num1 + 1;
            }
            dataGrid_Wndw_SW_S.Rows.Add("", "Subtotal", TotalArea);

            TotalWndwArea = TotalArea; //Total Window Area            

            foreach (Room rm in ACSPList) //For Detail View
            {
                dataGrid_Wndw_SW_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WndwTypeLst1_SW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WndwTypeLst_SW");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WndwAreaLst_SW");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WndwUValueLst_SW");
                string input3 = P3.AsString();
                Parameter P4 = rm.LookupParameter("WndwSC1ValueLst_SW");
                string input4 = P4.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst4 = new List<string>(
                           input4.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        dataGrid_Wndw_SW_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], Lst4[num1]);
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            foreach (string st in WndwTypeLst1_SW)
            {
                WndwTypeLst0.Add(st);
                WndwTypeLst1.Add(st);
            }

            foreach (string st in WndwTypeLst_SW)
            {
                WndwTypeLst.Add(st);
            }

            foreach (string st in WndwUValueLst_SW)
            {
                WndwUValueLst0.Add(st);
            }

            foreach (string st in WndwSC1ValueLst_SW)
            {
                WndwSC1ValueLst0.Add(st);
            }

            /////////////////////Window_South West\\\\\\\\\\\\\\\\\\\\\\\\\ 

            /////////////////////Wall_South West\\\\\\\\\\\\\\\\\\\\\\\\\

            Lst00 = new List<string>();
            Lst33 = new List<string>();
            WallTypeLst1_SW = new List<string>();
            WallTypeLst_SW = new List<string>();
            WallAreaLst_SW = new List<string>();
            WallArea_DouLst_SW = new List<double>();
            WallUValueLst_SW = new List<string>();

            foreach (Room rm in ACSPList)//to Create List for Column 1 & 2
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_SW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_SW");
                string input1 = P1.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_SW");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                num1 = 0;
                foreach (string st in Lst0)
                {
                    if (st != "")
                    {
                        Lst00.Add(st);
                        Lst33.Add(Lst3[num1]);

                        WallTypeLst1_SW.Add(st);
                        WallTypeLst_SW.Add(Lst1[num1]);
                    }
                    num1 = num1 + 1;
                }

            }

            WallTypeLst1_SW = WallTypeLst1_SW.Distinct().ToList();
            WallTypeLst_SW = WallTypeLst_SW.Distinct().ToList();

            num1 = 0;
            foreach (string st in WallTypeLst1_SW)
            {
                num2 = 0;
                foreach (string st1 in Lst00)
                {
                    if (st == st1)
                    {
                        WallUValueLst_SW.Add(Lst33[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            foreach (Room rm in ACSPList) //to Create List for Areas
            {
                Parameter P0 = rm.LookupParameter("WallTypeLst1_SW");
                string input0 = P0.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_SW");
                string input2 = P2.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));


                num2 = 0;
                foreach (string st in WallTypeLst1_SW)
                {
                    WallArea_DouLst_SW.Add(0);
                }

                num2 = 0;
                foreach (string WT in WallTypeLst1_SW)
                {
                    num1 = 0;
                    foreach (string st in Lst0)
                    {
                        if (st == WT)
                        {

                            WallArea_DouLst_SW[num2] = WallArea_DouLst_SW[num2] + (Convert.ToDouble(Lst2[num1]));

                        }
                        num1 = num1 + 1;
                    }
                    num2 = num2 + 1;
                }

            }

            TotalArea = 0;
            foreach (double db in WallArea_DouLst_SW) // to convert Double List to String List & add all areas together
            {
                WallAreaLst_SW.Add(Convert.ToString(db));
                TotalArea = TotalArea + db;
            }

            num1 = 0;
            CalValue = 0;
            TotalCalValue = 0;
            foreach (string st in WallTypeLst1_SW) //write into DataGridView
            {
                CalValue = 12 * Convert.ToDouble(WallAreaLst_SW[num1]) * Convert.ToDouble(WallUValueLst_SW[num1]);
                dataGrid_Wall_SW_S.Rows.Add(st, WallTypeLst_SW[num1], WallAreaLst_SW[num1], WallUValueLst_SW[num1], (CalValue.ToString("0.00")));
                num1 = num1 + 1;
                TotalCalValue = TotalCalValue + CalValue;
            }
            dataGrid_Wall_SW_S.Rows.Add("", "Subtotal", TotalArea, "", (TotalCalValue.ToString("0.00")));

            TotalWallArea = TotalArea; //Total Wall Area
            Lb_Area_SW.Text = (TotalWndwArea + TotalWallArea).ToString();
            SW_Wall_HG = TotalCalValue;
            //For Detail View
            foreach (Room rm in ACSPList)
            {
                dataGrid_Wall_SW_D.Rows.Add(rm.Name, "", "", "", "");

                Parameter P0 = rm.LookupParameter("WallTypeLst1_SW");
                string input0 = P0.AsString();
                Parameter P1 = rm.LookupParameter("WallTypeLst_SW");
                string input1 = P1.AsString();
                Parameter P2 = rm.LookupParameter("WallAreaLst_SW");
                string input2 = P2.AsString();
                Parameter P3 = rm.LookupParameter("WallUValueLst_SW");
                string input3 = P3.AsString();

                List<string> Lst0 = new List<string>(
                           input0.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst1 = new List<string>(
                           input1.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst2 = new List<string>(
                           input2.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                List<string> Lst3 = new List<string>(
                           input3.Split(new string[] { "\r\n", "\n" },
                           StringSplitOptions.None));

                CalValue = 0;
                num1 = 0;
                foreach (string st in Lst1)
                {
                    if (st != "")
                    {
                        CalValue = 12 * Convert.ToDouble(Lst2[num1]) * Convert.ToDouble(Lst3[num1]);
                        dataGrid_Wall_SW_D.Rows.Add(Lst0[num1], st, Lst2[num1], Lst3[num1], (CalValue.ToString("0.00")));
                    }
                    num1 = num1 + 1;
                }

            }//For Detail View

            ////////// For Wall Types Page //////////////////////
            foreach (string st in WallTypeLst1_SW)
            {
                WallTypeLst0.Add(st);
                WallTypeLst1.Add(st);
            }

            foreach (string st in WallTypeLst_SW)
            {
                WallTypeLst.Add(st);
            }

            foreach (string st in WallUValueLst_SW)
            {
                WallUValueLst0.Add(st);
            }
            ////////// For Wall Types Page //////////////////////

            /////////////////////Wall_South West\\\\\\\\\\\\\\\\\\\\\\\\\

            ////////////////////////////////////////////////////////////////////////////////////////////////////////

            ////////////////////// For Window and Wall Types Summary Page ////////////////////////////////

            ////////////////// For Window Types
            WndwTypeLst1 = WndwTypeLst1.Distinct().ToList();
            WndwTypeLst = WndwTypeLst.Distinct().ToList();

            ///Get window elements list (1 window per type)
            foreach (string st in WndwTypeLst)
            {
                num1 = 0;
                FilteredElementCollector wndw_collector = new FilteredElementCollector(doc2);
                ElementCategoryFilter wndw_filter = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
                IList<Element> wndws = wndw_collector.WherePasses(wndw_filter).WhereElementIsNotElementType().Cast<Element>().Where(w => (((doc2.GetElement(w.GetTypeId()) as ElementType).FamilyName) + " " + (w.Name)) == st).ToList();

                WndwLst_F2.Add(wndws[num1]);
            }

            foreach (Element w in WndwLst_F2)
            {
                WndwTypeId_F2.Add(w.GetTypeId());
            }

            foreach (ElementId WnTId in WndwTypeId_F2)
            {
                WndwTypeLst_F2.Add(doc2.GetElement(WnTId) as ElementType);
            }


            num1 = 0;
            foreach (string st in WndwTypeLst1)
            {
                num2 = 0;
                foreach (string st1 in WndwTypeLst0)
                {
                    if (st == st1)
                    {
                        WndwUValueLst.Add(WndwUValueLst0[num2]);
                        WndwSC1ValueLst.Add(WndwSC1ValueLst0[num2]);
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            num1 = 0;
            foreach (string st in WndwTypeLst1)
            {
                Parameter U1 = WndwTypeLst_F2[num1].LookupParameter("SC_Angle");
                double WndwAgl = U1.AsDouble();
                Parameter U2 = WndwTypeLst_F2[num1].LookupParameter("Height");
                double WndwH = U2.AsDouble();
                Parameter U3 = WndwTypeLst_F2[num1].LookupParameter("Width");
                double WndwW = U3.AsDouble();
                Parameter U4 = WndwTypeLst_F2[num1].LookupParameter("Horizontal Projection");
                int HP = U4.AsInteger();
                Parameter U5 = WndwTypeLst_F2[num1].LookupParameter("Vertical Projection");
                int VP = U5.AsInteger();
                Parameter U6 = WndwTypeLst_F2[num1].LookupParameter("Egg Crate Window");
                int ECW = U6.AsInteger();
                Parameter U7 = WndwTypeLst_F2[num1].LookupParameter("SC_P");
                double WndwP = U7.AsDouble();

                if (HP==1)
                {
                    dataGrid_Wndw_Types.Rows.Add(st, (WndwTypeLst[num1]), (WndwUValueLst[num1]), (WndwSC1ValueLst[num1]), ("Horizontal Projection"), (WndwAgl * _footToMm), ((WndwP * _footToMm)/1000), ((WndwH * _footToMm) / 1000), ((WndwW * _footToMm) / 1000));
                    num1 = num1 + 1;
                }

                else if (VP==1)
                {
                    dataGrid_Wndw_Types.Rows.Add(st, (WndwTypeLst[num1]), (WndwUValueLst[num1]), (WndwSC1ValueLst[num1]), ("Vertical Projection"), (WndwAgl * _footToMm), ((WndwP * _footToMm)/1000), ((WndwH * _footToMm) / 1000), ((WndwW * _footToMm) / 1000));
                    num1 = num1 + 1;
                }

                else if (ECW==1)
                {
                    dataGrid_Wndw_Types.Rows.Add(st, (WndwTypeLst[num1]), (WndwUValueLst[num1]), (WndwSC1ValueLst[num1]), ("Egg Crate Window"), (WndwAgl * _footToMm), ((WndwP * _footToMm)/1000), ((WndwH * _footToMm) / 1000), ((WndwW * _footToMm) / 1000));
                    num1 = num1 + 1;
                }

                else
                {
                    dataGrid_Wndw_Types.Rows.Add(st, (WndwTypeLst[num1]), (WndwUValueLst[num1]), (WndwSC1ValueLst[num1]), ("None"), (WndwAgl * _footToMm), (0), ((WndwH * _footToMm) / 1000), ((WndwW * _footToMm) / 1000));
                    num1 = num1 + 1;
                }


               
            }
            dataGrid_Wndw_Types.Sort(dataGrid_Wndw_Types.Columns["TypeName"], ListSortDirection.Ascending);

            
            ////////////////// For Window Types

            ////////////////// For Wall Types

            WallTypeLst1 = WallTypeLst1.Distinct().ToList();
            WallTypeLst = WallTypeLst.Distinct().ToList();

            ///Get wall elements list (1 wall per walltype)
            foreach (string st in WallTypeLst)
            {
                num1 = 0;
                FilteredElementCollector wall_collector = new FilteredElementCollector(doc2);
                ElementCategoryFilter wall_filter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
                IList<Wall> walls = wall_collector.WherePasses(wall_filter).WhereElementIsNotElementType().Cast<Wall>().Where(w => ((w.WallType.FamilyName) + " " + (w.WallType.Name)) == st).ToList();

                WallLst_F2.Add(walls[num1]);
            }

            foreach (Wall w in WallLst_F2)
            {
                WallTypeId_F2.Add(w.GetTypeId());
            }

            foreach (ElementId WTId in WallTypeId_F2)
            {
                WallTypeLst_F2.Add(doc2.GetElement(WTId) as ElementType);
            }

            num1 = 0;
            foreach (string st in WallTypeLst1)
            {
                num2 = 0;
                foreach (string st1 in WallTypeLst0)
                {
                    if (st == st1)
                    {
                        WallUValueLst.Add(WallUValueLst0[num2]);                        
                        break;
                    }
                    num2 = num2 + 1;
                }
                num1 = num1 + 1;
            }

            num1 = 0;
            foreach (string st in WallTypeLst1)
            {
                Parameter U = WallTypeLst_F2[num1].LookupParameter("Width");
                double WallThk = U.AsDouble();

                dataGrid_Wall_Types.Rows.Add(st, (WallTypeLst[num1]), (WallUValueLst[num1]),(WallThk * _footToMm));
                num1 = num1 + 1;
            }
            dataGrid_Wall_Types.Sort(dataGrid_Wall_Types.Columns["WallTypeName"], ListSortDirection.Ascending);

            ////////////////// For Wall Types        


            ///////////////////// For Wall Layer Processing///////////////////////////                     
            num1 = 0;
            foreach (Wall w in WallLst_F2)
            {
                
                CompoundStructure cs = w.WallType.GetCompoundStructure();
                int layerIndex = cs.GetFirstCoreLayerIndex();
                IList<CompoundStructureLayer> cslayers = cs.GetLayers();
                //dataGridView5.Rows.Add("W" + num1, " ", " ", " ", " ", " ");
                dataGridView5.Rows.Add(((WallTypeLst1[num1]))," ", " ", " "," "," ");
                num2 = 0;
                num3 = 0;
                foreach (CompoundStructureLayer csl in cslayers)
                {
                    Material M = doc2.GetElement(cs.GetMaterialId(num2)) as Material;

                    ElementId ThId = M.ThermalAssetId;
                                        
                    double KValue = UnitUtils.ConvertFromInternalUnits((doc2.GetElement(ThId).LookupParameter("Thermal Conductivity")).AsDouble(), UnitTypeId.WattsPerMeterKelvin);
                    double DensityValue = UnitUtils.ConvertFromInternalUnits((doc2.GetElement(ThId).LookupParameter("Density")).AsDouble(), UnitTypeId.KilogramsPerCubicMeter);
                    double RValue = ((cs.GetLayerWidth(num2) * _footToMm)/1000) / KValue;
                    //Parameter U = doc2.GetElement(ThId).LookupParameter("Thermal Conductivity");
                    //double UValue = U.AsDouble();
                    num3 = num3 + RValue;
                    dataGridView5.Rows.Add(cs.GetLayerFunction(num2),M.Name, (cs.GetLayerWidth(num2) * _footToMm),KValue,DensityValue,RValue.ToString("0.0000"));
                    num2++;
                }
                dataGridView5.Rows.Add(" ", " ", " ", "Total R Value:", " ", num3.ToString("0.0000"));
                dataGridView5.Rows.Add(" ", " ", " ", "U Value:", " ", (1/num3).ToString("0.0000"));
                dataGridView5.Rows.Add(" ", " ", " ", " "," "," ");
                num1++;
            }
            ///////////////////// For Wall Layer Processing///////////////////////////
            

            ////////////////////////Transfer SC2 Values to DataGridView///////////////////////////////
            ////////////////////////For Interpolations/////////////////////////////////
            
            ///////////////////For Horizontal Projection///////////////////////////////
            Workbook workbook = new Workbook();
            string excelpath = BuildingCoder.Util.GetFilePath("Effective_SC.xlsx");
            workbook.LoadFromFile(excelpath);
            //workbook.LoadFromFile("D:\\WIP\\01_API\\ETTV\\Effective_SC.xlsx");

            //////////////////North-South///////////////////////////////////////
            Worksheet worksheet = workbook.Worksheets[0];                       
            DataTable dt = worksheet.ExportDataTable();            
            dataGrid_SC2_NS.DataSource = dt;
            //////////////////East-West///////////////////////////////////////            
            Worksheet worksheet1 = workbook.Worksheets[1];
            DataTable dt1 = worksheet1.ExportDataTable();
            dataGrid_SC2_EW.DataSource = dt1;
            //////////////////NorthEast-NorthWest///////////////////////////////////////
            Worksheet worksheet2 = workbook.Worksheets[2];
            DataTable dt2 = worksheet2.ExportDataTable();
            dataGrid_SC2_NENW.DataSource = dt2;
            //////////////////NorthEast-NorthWest///////////////////////////////////////
            Worksheet worksheet3 = workbook.Worksheets[3];
            DataTable dt3 = worksheet3.ExportDataTable();
            dataGrid_SC2_SESW.DataSource = dt3;

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            ///////////////////For Vertical Projection///////////////////////////////
            Workbook workbook_VP = new Workbook();
            string excelpath_vp = BuildingCoder.Util.GetFilePath("Effective_SC_VP.xlsx");
            workbook_VP.LoadFromFile(excelpath_vp);
            //workbook_VP.LoadFromFile("D:\\WIP\\01_API\\ETTV\\Effective_SC_VP.xlsx");

            //////////////////North-South///////////////////////////////////////
            Worksheet worksheet_VP = workbook_VP.Worksheets[0];
            DataTable dt_VP = worksheet_VP.ExportDataTable();
            dataGrid_SC2_VP_NS.DataSource = dt_VP;
            //////////////////East-West///////////////////////////////////////            
            Worksheet worksheet1_VP = workbook_VP.Worksheets[1];
            DataTable dt1_VP = worksheet1_VP.ExportDataTable();
            dataGrid_SC2_VP_EW.DataSource = dt1_VP;
            //////////////////NorthEast-NorthWest///////////////////////////////////////
            Worksheet worksheet2_VP = workbook_VP.Worksheets[2];
            DataTable dt2_VP = worksheet2_VP.ExportDataTable();
            dataGrid_SC2_VP_NENW.DataSource = dt2_VP;
            //////////////////NorthEast-NorthWest///////////////////////////////////////
            Worksheet worksheet3_VP = workbook_VP.Worksheets[3];
            DataTable dt3_VP = worksheet3_VP.ExportDataTable();
            dataGrid_SC2_VP_SESW.DataSource = dt3_VP;

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ///////////////////For Egg Crate///////////////////////////////
            Workbook workbook_EC_NS = new Workbook();
            string excelpath_EC_NS = BuildingCoder.Util.GetFilePath("Effective_SC_EC_NS.xlsx");
            workbook_EC_NS.LoadFromFile(excelpath_EC_NS);
            //workbook_EC_NS.LoadFromFile("D:\\WIP\\01_API\\ETTV\\Effective_SC_EC_NS.xlsx");

            Workbook workbook_EC_EW = new Workbook();
            string excelpath_EC_EW = BuildingCoder.Util.GetFilePath("Effective_SC_EC_EW.xlsx");
            workbook_EC_EW.LoadFromFile(excelpath_EC_EW);
            //workbook_EC_EW.LoadFromFile("D:\\WIP\\01_API\\ETTV\\Effective_SC_EC_EW.xlsx");

            Workbook workbook_EC_NENW = new Workbook();
            string excelpath_EC_NENW = BuildingCoder.Util.GetFilePath("Effective_SC_EC_NE_NW.xlsx");
            workbook_EC_NENW.LoadFromFile(excelpath_EC_NENW);
            //workbook_EC_NENW.LoadFromFile("D:\\WIP\\01_API\\ETTV\\Effective_SC_EC_NE_NW.xlsx");

            Workbook workbook_EC_SESW = new Workbook();
            string excelpath_EC_SESW = BuildingCoder.Util.GetFilePath("Effective_SC_EC_SE_SW.xlsx");
            workbook_EC_SESW.LoadFromFile(excelpath_EC_SESW);
            //workbook_EC_SESW.LoadFromFile("D:\\WIP\\01_API\\ETTV\\Effective_SC_EC_SE_SW.xlsx");


            //////////////////North-South///////////////////////////////////////
            ///R1(0.2)///
            Worksheet worksheet_EC_NS_0 = workbook_EC_NS.Worksheets[0];
            DataTable dt_EC_NS_0 = worksheet_EC_NS_0.ExportDataTable();
            dataGrid_SC2_EC_NS_0.DataSource = dt_EC_NS_0;
            ///R1(0.4)///
            Worksheet worksheet_EC_NS_1 = workbook_EC_NS.Worksheets[1];
            DataTable dt_EC_NS_1 = worksheet_EC_NS_1.ExportDataTable();
            dataGrid_SC2_EC_NS_1.DataSource = dt_EC_NS_1;
            ///R1(0.6)///
            Worksheet worksheet_EC_NS_2 = workbook_EC_NS.Worksheets[2];
            DataTable dt_EC_NS_2 = worksheet_EC_NS_2.ExportDataTable();
            dataGrid_SC2_EC_NS_2.DataSource = dt_EC_NS_2;
            ///R1(0.8)///
            Worksheet worksheet_EC_NS_3 = workbook_EC_NS.Worksheets[3];
            DataTable dt_EC_NS_3 = worksheet_EC_NS_3.ExportDataTable();
            dataGrid_SC2_EC_NS_3.DataSource = dt_EC_NS_3;
            ///R1(1.0)///
            Worksheet worksheet_EC_NS_4 = workbook_EC_NS.Worksheets[4];
            DataTable dt_EC_NS_4 = worksheet_EC_NS_4.ExportDataTable();
            dataGrid_SC2_EC_NS_4.DataSource = dt_EC_NS_4;
            ///R1(1.2)///
            Worksheet worksheet_EC_NS_5 = workbook_EC_NS.Worksheets[5];
            DataTable dt_EC_NS_5 = worksheet_EC_NS_5.ExportDataTable();
            dataGrid_SC2_EC_NS_5.DataSource = dt_EC_NS_5;
            ///R1(1.4)///
            Worksheet worksheet_EC_NS_6 = workbook_EC_NS.Worksheets[6];
            DataTable dt_EC_NS_6 = worksheet_EC_NS_6.ExportDataTable();
            dataGrid_SC2_EC_NS_6.DataSource = dt_EC_NS_6;
            ///R1(1.6)///
            Worksheet worksheet_EC_NS_7 = workbook_EC_NS.Worksheets[7];
            DataTable dt_EC_NS_7 = worksheet_EC_NS_7.ExportDataTable();
            dataGrid_SC2_EC_NS_7.DataSource = dt_EC_NS_7;
            ///R1(1.8)///
            Worksheet worksheet_EC_NS_8 = workbook_EC_NS.Worksheets[8];
            DataTable dt_EC_NS_8 = worksheet_EC_NS_8.ExportDataTable();
            dataGrid_SC2_EC_NS_8.DataSource = dt_EC_NS_8;

            //////////////////East-West///////////////////////////////////////
            ///R1(0.2)///
            Worksheet worksheet_EC_EW_0 = workbook_EC_EW.Worksheets[0];
            DataTable dt_EC_EW_0 = worksheet_EC_EW_0.ExportDataTable();
            dataGrid_SC2_EC_EW_0.DataSource = dt_EC_EW_0;
            ///R1(0.4)///
            Worksheet worksheet_EC_EW_1 = workbook_EC_EW.Worksheets[1];
            DataTable dt_EC_EW_1 = worksheet_EC_EW_1.ExportDataTable();
            dataGrid_SC2_EC_EW_1.DataSource = dt_EC_EW_1;
            ///R1(0.6)///
            Worksheet worksheet_EC_EW_2 = workbook_EC_EW.Worksheets[2];
            DataTable dt_EC_EW_2 = worksheet_EC_EW_2.ExportDataTable();
            dataGrid_SC2_EC_EW_2.DataSource = dt_EC_EW_2;
            ///R1(0.8)///
            Worksheet worksheet_EC_EW_3 = workbook_EC_EW.Worksheets[3];
            DataTable dt_EC_EW_3 = worksheet_EC_EW_3.ExportDataTable();
            dataGrid_SC2_EC_EW_3.DataSource = dt_EC_EW_3;
            ///R1(1.0)///
            Worksheet worksheet_EC_EW_4 = workbook_EC_EW.Worksheets[4];
            DataTable dt_EC_EW_4 = worksheet_EC_EW_4.ExportDataTable();
            dataGrid_SC2_EC_EW_4.DataSource = dt_EC_EW_4;
            ///R1(1.2)///
            Worksheet worksheet_EC_EW_5 = workbook_EC_EW.Worksheets[5];
            DataTable dt_EC_EW_5 = worksheet_EC_EW_5.ExportDataTable();
            dataGrid_SC2_EC_EW_5.DataSource = dt_EC_EW_5;
            ///R1(1.4)///
            Worksheet worksheet_EC_EW_6 = workbook_EC_EW.Worksheets[6];
            DataTable dt_EC_EW_6 = worksheet_EC_EW_6.ExportDataTable();
            dataGrid_SC2_EC_EW_6.DataSource = dt_EC_EW_6;
            ///R1(1.6)///
            Worksheet worksheet_EC_EW_7 = workbook_EC_EW.Worksheets[7];
            DataTable dt_EC_EW_7 = worksheet_EC_EW_7.ExportDataTable();
            dataGrid_SC2_EC_EW_7.DataSource = dt_EC_EW_7;
            ///R1(1.8)///
            Worksheet worksheet_EC_EW_8 = workbook_EC_EW.Worksheets[8];
            DataTable dt_EC_EW_8 = worksheet_EC_EW_8.ExportDataTable();
            dataGrid_SC2_EC_EW_8.DataSource = dt_EC_EW_8;

            //////////////////NorthEast-NorthWest///////////////////////////////////////
            ///R1(0.2)///
            Worksheet worksheet_EC_NENW_0 = workbook_EC_NENW.Worksheets[0];
            DataTable dt_EC_NENW_0 = worksheet_EC_NENW_0.ExportDataTable();
            dataGrid_SC2_EC_NENW_0.DataSource = dt_EC_NENW_0;
            ///R1(0.4)///
            Worksheet worksheet_EC_NENW_1 = workbook_EC_NENW.Worksheets[1];
            DataTable dt_EC_NENW_1 = worksheet_EC_NENW_1.ExportDataTable();
            dataGrid_SC2_EC_NENW_1.DataSource = dt_EC_NENW_1;
            ///R1(0.6)///
            Worksheet worksheet_EC_NENW_2 = workbook_EC_NENW.Worksheets[2];
            DataTable dt_EC_NENW_2 = worksheet_EC_NENW_2.ExportDataTable();
            dataGrid_SC2_EC_NENW_2.DataSource = dt_EC_NENW_2;
            ///R1(0.8)///
            Worksheet worksheet_EC_NENW_3 = workbook_EC_NENW.Worksheets[3];
            DataTable dt_EC_NENW_3 = worksheet_EC_NENW_3.ExportDataTable();
            dataGrid_SC2_EC_NENW_3.DataSource = dt_EC_NENW_3;
            ///R1(1.0)///
            Worksheet worksheet_EC_NENW_4 = workbook_EC_NENW.Worksheets[4];
            DataTable dt_EC_NENW_4 = worksheet_EC_NENW_4.ExportDataTable();
            dataGrid_SC2_EC_NENW_4.DataSource = dt_EC_NENW_4;
            ///R1(1.2)///
            Worksheet worksheet_EC_NENW_5 = workbook_EC_NENW.Worksheets[5];
            DataTable dt_EC_NENW_5 = worksheet_EC_NENW_5.ExportDataTable();
            dataGrid_SC2_EC_NENW_5.DataSource = dt_EC_NENW_5;
            ///R1(1.4)///
            Worksheet worksheet_EC_NENW_6 = workbook_EC_NENW.Worksheets[6];
            DataTable dt_EC_NENW_6 = worksheet_EC_NENW_6.ExportDataTable();
            dataGrid_SC2_EC_NENW_6.DataSource = dt_EC_NENW_6;
            ///R1(1.6)///
            Worksheet worksheet_EC_NENW_7 = workbook_EC_NENW.Worksheets[7];
            DataTable dt_EC_NENW_7 = worksheet_EC_NENW_7.ExportDataTable();
            dataGrid_SC2_EC_NENW_7.DataSource = dt_EC_NENW_7;
            ///R1(1.8)///
            Worksheet worksheet_EC_NENW_8 = workbook_EC_NENW.Worksheets[8];
            DataTable dt_EC_NENW_8 = worksheet_EC_NENW_8.ExportDataTable();
            dataGrid_SC2_EC_NENW_8.DataSource = dt_EC_NENW_8;

            //////////////////NorthEast-NorthWest///////////////////////////////////////
            ///R1(0.2)///
            Worksheet worksheet_EC_SESW_0 = workbook_EC_SESW.Worksheets[0];
            DataTable dt_EC_SESW_0 = worksheet_EC_SESW_0.ExportDataTable();
            dataGrid_SC2_EC_SESW_0.DataSource = dt_EC_SESW_0;
            ///R1(0.4)///
            Worksheet worksheet_EC_SESW_1 = workbook_EC_SESW.Worksheets[1];
            DataTable dt_EC_SESW_1 = worksheet_EC_SESW_1.ExportDataTable();
            dataGrid_SC2_EC_SESW_1.DataSource = dt_EC_SESW_1;
            ///R1(0.6)///
            Worksheet worksheet_EC_SESW_2 = workbook_EC_SESW.Worksheets[2];
            DataTable dt_EC_SESW_2 = worksheet_EC_SESW_2.ExportDataTable();
            dataGrid_SC2_EC_SESW_2.DataSource = dt_EC_SESW_2;
            ///R1(0.8)///
            Worksheet worksheet_EC_SESW_3 = workbook_EC_SESW.Worksheets[3];
            DataTable dt_EC_SESW_3 = worksheet_EC_SESW_3.ExportDataTable();
            dataGrid_SC2_EC_SESW_3.DataSource = dt_EC_SESW_3;
            ///R1(1.0)///
            Worksheet worksheet_EC_SESW_4 = workbook_EC_SESW.Worksheets[4];
            DataTable dt_EC_SESW_4 = worksheet_EC_SESW_4.ExportDataTable();
            dataGrid_SC2_EC_SESW_4.DataSource = dt_EC_SESW_4;
            ///R1(1.2)///
            Worksheet worksheet_EC_SESW_5 = workbook_EC_SESW.Worksheets[5];
            DataTable dt_EC_SESW_5 = worksheet_EC_SESW_5.ExportDataTable();
            dataGrid_SC2_EC_SESW_5.DataSource = dt_EC_SESW_5;
            ///R1(1.4)///
            Worksheet worksheet_EC_SESW_6 = workbook_EC_SESW.Worksheets[6];
            DataTable dt_EC_SESW_6 = worksheet_EC_SESW_6.ExportDataTable();
            dataGrid_SC2_EC_SESW_6.DataSource = dt_EC_SESW_6;
            ///R1(1.6)///
            Worksheet worksheet_EC_SESW_7 = workbook_EC_SESW.Worksheets[7];
            DataTable dt_EC_SESW_7 = worksheet_EC_SESW_7.ExportDataTable();
            dataGrid_SC2_EC_SESW_7.DataSource = dt_EC_SESW_7;
            ///R1(1.8)///
            Worksheet worksheet_EC_SESW_8 = workbook_EC_SESW.Worksheets[8];
            DataTable dt_EC_SESW_8 = worksheet_EC_SESW_8.ExportDataTable();
            dataGrid_SC2_EC_SESW_8.DataSource = dt_EC_SESW_8;

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            ////////////////////// For Window SC Pages (Dynamically Create Forms Controls) /////////////////////////////
            num22 = Main_TabPage.TabCount;
            WndwTypeLst1.Sort();
            foreach (string st in WndwTypeLst1)
            {
                Main_TabPage.TabPages.Add("AP5_"+st);
            }

            num1 = 0;
            num5 = 0;
            /////////////////// Dynamically Create Forms Controls for all Fenestration Pages
            foreach (TabPage tab in Main_TabPage.TabPages)
            {                
                if (num5>(num22-1))
                {
                    Label resultLabel = new Label();
                    resultLabel.Location = new System.Drawing.Point(520, 5);
                    resultLabel.Width = 288;
                    resultLabel.Height = 36;
                    resultLabel.Font = new Font("Microsoft Sans Serif", 18, FontStyle.Bold | FontStyle.Underline);
                    resultLabel.Text = "SC CALCULATION";

                    Label resultLabel1 = new Label();
                    resultLabel1.Location = new System.Drawing.Point(530, 45);
                    resultLabel1.Width = 210;
                    resultLabel1.Height = 29;
                    resultLabel1.Font = new Font("Microsoft Sans Serif", 14, FontStyle.Bold);
                    resultLabel1.Text = "FENESTRATION: ";

                    Label resultLabel4 = new Label();
                    resultLabel4.Location = new System.Drawing.Point(750, 45);
                    resultLabel4.Width = 100;
                    resultLabel4.Height = 29;
                    resultLabel4.Font = new Font("Microsoft Sans Serif", 14, FontStyle.Bold);
                    resultLabel4.Text = WndwTypeLst1[num1];

                    //////////////////////////////////////////////////////////////////////////////////////////////////

                    Label resultLabel2 = new Label();
                    resultLabel2.Location = new System.Drawing.Point(10, 100);
                    resultLabel2.Width = 50;
                    resultLabel2.Height = 20;
                    resultLabel2.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    resultLabel2.Text = "SC1:";

                    System.Windows.Forms.TextBox TB1 = new System.Windows.Forms.TextBox();
                    TB1.Location = new System.Drawing.Point(122, 100);                                        
                    TB1.Text = dataGrid_Wndw_Types.Rows[num1].Cells[3].Value.ToString(); 

                    //////////////////////////////////////////////////////////////////////////////////////////////////
                    Label resultLabel3 = new Label();
                    resultLabel3.Location = new System.Drawing.Point(10, 150);
                    resultLabel3.Width = 81;
                    resultLabel3.Height = 20;
                    resultLabel3.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    resultLabel3.Text = "U Value:";

                    System.Windows.Forms.TextBox TB2 = new System.Windows.Forms.TextBox();
                    TB2.Location = new System.Drawing.Point(122, 150);
                    TB2.Text = dataGrid_Wndw_Types.Rows[num1].Cells[2].Value.ToString();

                    Label LB1 = new Label();
                    LB1.Location = new System.Drawing.Point(227, 150);
                    LB1.Width = 81;
                    LB1.Height = 20;
                    LB1.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB1.Text = "W/m2K";
                    //////////////////////////////////////////////////////////////////////////////////////////////////
                    PictureBox PB1 = new PictureBox();   
                    

                    //////////////////////////////////////////////////////////////////////////////////////////////////
                    Label LB2 = new Label();
                    LB2.Location = new System.Drawing.Point(10, 200);
                    LB2.Width = 81;
                    LB2.Height = 20;
                    LB2.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB2.Text = "Shades:";

                    System.Windows.Forms.ComboBox CB1 = new System.Windows.Forms.ComboBox();
                    CB1.Location = new System.Drawing.Point(122, 197);
                    CB1.Width = 200;
                    CB1.Height = 20;
                    CB1.Items.Add("None");
                    CB1.Items.Add("Horizontal Projection");
                    CB1.Items.Add("Vertical Projection");
                    CB1.Items.Add("Egg Crate Window");
                                                        


                    Label LB3 = new Label();
                    LB3.Location = new System.Drawing.Point(7, 555);
                    LB3.Width = 110;
                    LB3.Height = 20;
                    LB3.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB3.Text = "North-South";

                    Label LB3a = new Label();
                    LB3a.Location = new System.Drawing.Point(256, 555);
                    LB3a.Width = 50;
                    LB3a.Height = 20;
                    LB3a.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB3a.Text = "SC2:";

                    Label LB3b = new Label();
                    LB3b.Location = new System.Drawing.Point(475, 555);
                    LB3b.Width = 150;
                    LB3b.Height = 20;
                    LB3b.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB3b.Text = "SC = SC1 x SC2";

                    System.Windows.Forms.TextBox TB_SC2_NS = new System.Windows.Forms.TextBox();
                    TB_SC2_NS.Location = new System.Drawing.Point(313, 555);
                    System.Windows.Forms.TextBox TB_SC_NS = new System.Windows.Forms.TextBox();
                    TB_SC_NS.Location = new System.Drawing.Point(630, 555);


                    Label LB4 = new Label();
                    LB4.Location = new System.Drawing.Point(7, 605);
                    LB4.Width = 110;
                    LB4.Height = 20;
                    LB4.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB4.Text = "East-West";

                    Label LB4a = new Label();
                    LB4a.Location = new System.Drawing.Point(256, 605);
                    LB4a.Width = 50;
                    LB4a.Height = 20;
                    LB4a.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB4a.Text = "SC2:";

                    Label LB4b = new Label();
                    LB4b.Location = new System.Drawing.Point(475, 605);
                    LB4b.Width = 150;
                    LB4b.Height = 20;
                    LB4b.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB4b.Text = "SC = SC1 x SC2";

                    System.Windows.Forms.TextBox TB_SC2_EW = new System.Windows.Forms.TextBox();
                    TB_SC2_EW.Location = new System.Drawing.Point(313, 605);
                    System.Windows.Forms.TextBox TB_SC_EW = new System.Windows.Forms.TextBox();
                    TB_SC_EW.Location = new System.Drawing.Point(630, 605);


                    Label LB5 = new Label();
                    LB5.Location = new System.Drawing.Point(7, 655);
                    LB5.Width = 190;
                    LB5.Height = 20;
                    LB5.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB5.Text = "NorthEast-NorthWest";

                    Label LB5a = new Label();
                    LB5a.Location = new System.Drawing.Point(256, 655);
                    LB5a.Width = 50;
                    LB5a.Height = 20;
                    LB5a.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB5a.Text = "SC2:";

                    Label LB5b = new Label();
                    LB5b.Location = new System.Drawing.Point(475, 655);
                    LB5b.Width = 150;
                    LB5b.Height = 20;
                    LB5b.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB5b.Text = "SC = SC1 x SC2";

                    System.Windows.Forms.TextBox TB_SC2_NENW = new System.Windows.Forms.TextBox();
                    TB_SC2_NENW.Location = new System.Drawing.Point(313, 655);
                    System.Windows.Forms.TextBox TB_SC_NENW = new System.Windows.Forms.TextBox();
                    TB_SC_NENW.Location = new System.Drawing.Point(630, 655);


                    Label LB6 = new Label();
                    LB6.Location = new System.Drawing.Point(7, 705);
                    LB6.Width = 190;
                    LB6.Height = 20;
                    LB6.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB6.Text = "SouthEast-SouthWest";

                    Label LB6a = new Label();
                    LB6a.Location = new System.Drawing.Point(256, 705);
                    LB6a.Width = 50;
                    LB6a.Height = 20;
                    LB6a.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB6a.Text = "SC2:";

                    Label LB6b = new Label();
                    LB6b.Location = new System.Drawing.Point(475, 705);
                    LB6b.Width = 150;
                    LB6b.Height = 20;
                    LB6b.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB6b.Text = "SC = SC1 x SC2";

                    System.Windows.Forms.TextBox TB_SC2_SESW = new System.Windows.Forms.TextBox();
                    TB_SC2_SESW.Location = new System.Drawing.Point(313, 705);
                    System.Windows.Forms.TextBox TB_SC_SESW = new System.Windows.Forms.TextBox();
                    TB_SC_SESW.Location = new System.Drawing.Point(630, 705);

                    ////////////////////////////////////////////////////////////////////////
                    Label LB7 = new Label();
                    LB7.Location = new System.Drawing.Point(1000, 100);
                    LB7.Width = 105;
                    LB7.Height = 20;
                    LB7.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB7.Text = "Angle (deg):";
                    System.Windows.Forms.TextBox TB7 = new System.Windows.Forms.TextBox();
                    TB7.Location = new System.Drawing.Point(1130, 100);
                    TB7.Text = dataGrid_Wndw_Types.Rows[num1].Cells[5].Value.ToString();


                    Label LB8 = new Label();
                    LB8.Location = new System.Drawing.Point(1000, 150);
                    LB8.Width = 105;
                    LB8.Height = 20;
                    LB8.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB8.Text = "P (m):";
                    System.Windows.Forms.TextBox TB8 = new System.Windows.Forms.TextBox();
                    TB8.Location = new System.Drawing.Point(1130, 150);
                    TB8.Text = dataGrid_Wndw_Types.Rows[num1].Cells[6].Value.ToString();

                    Label LB9 = new Label();
                    LB9.Location = new System.Drawing.Point(1000, 200);
                    LB9.Width = 105;
                    LB9.Height = 20;
                    LB9.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB9.Text = "H (m):";
                    System.Windows.Forms.TextBox TB9 = new System.Windows.Forms.TextBox();
                    TB9.Location = new System.Drawing.Point(1130, 200);
                    TB9.Text = dataGrid_Wndw_Types.Rows[num1].Cells[7].Value.ToString();

                    Label LB10 = new Label();
                    LB10.Location = new System.Drawing.Point(1000, 250);
                    LB10.Width = 105;
                    LB10.Height = 20;
                    LB10.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB10.Text = "W (m):";
                    System.Windows.Forms.TextBox TB10 = new System.Windows.Forms.TextBox();
                    TB10.Location = new System.Drawing.Point(1130, 250);
                    TB10.Text = dataGrid_Wndw_Types.Rows[num1].Cells[8].Value.ToString();

                    Label LB11 = new Label();
                    LB11.Location = new System.Drawing.Point(1000, 300);
                    LB11.Width = 105;
                    LB11.Height = 20;
                    LB11.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB11.Text = "R1 = P/H:";
                    System.Windows.Forms.TextBox TB11 = new System.Windows.Forms.TextBox();
                    TB11.Location = new System.Drawing.Point(1130, 300);
                    

                    Label LB12 = new Label();
                    LB12.Location = new System.Drawing.Point(1000, 350);
                    LB12.Width = 105;
                    LB12.Height = 20;
                    LB12.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    LB12.Text = "R2 = P/W:";
                    System.Windows.Forms.TextBox TB12 = new System.Windows.Forms.TextBox();
                    TB12.Location = new System.Drawing.Point(1130, 350);

                    ///////////////////////////////////////////////////////////////////////
                    Button btn = new Button();
                    btn.Location = new System.Drawing.Point(1050, 450);
                    btn.Text = "Update";
                    btn.BackColor= System.Drawing.Color.BlueViolet;
                    btn.Width = 160;
                    btn.Height = 50;
                    ///////////////////////////////////////////////////////////////////////

                    Main_TabPage.TabPages[num5].Controls.Add(resultLabel);
                    Main_TabPage.TabPages[num5].Controls.Add(resultLabel1);
                    Main_TabPage.TabPages[num5].Controls.Add(resultLabel2);
                    Main_TabPage.TabPages[num5].Controls.Add(resultLabel3);
                    Main_TabPage.TabPages[num5].Controls.Add(resultLabel4);
                    Main_TabPage.TabPages[num5].Controls.Add(TB1);
                    Main_TabPage.TabPages[num5].Controls.Add(TB2);
                    Main_TabPage.TabPages[num5].Controls.Add(LB1);
                    Main_TabPage.TabPages[num5].Controls.Add(LB2);
                    Main_TabPage.TabPages[num5].Controls.Add(CB1);
                    Main_TabPage.TabPages[num5].Controls.Add(PB1);
                    Main_TabPage.TabPages[num5].Controls.Add(LB3);
                    Main_TabPage.TabPages[num5].Controls.Add(LB4);
                    Main_TabPage.TabPages[num5].Controls.Add(LB5);
                    Main_TabPage.TabPages[num5].Controls.Add(LB6);
                    Main_TabPage.TabPages[num5].Controls.Add(LB3a);
                    Main_TabPage.TabPages[num5].Controls.Add(LB4a);
                    Main_TabPage.TabPages[num5].Controls.Add(LB5a);
                    Main_TabPage.TabPages[num5].Controls.Add(LB6a);
                    Main_TabPage.TabPages[num5].Controls.Add(LB3b);
                    Main_TabPage.TabPages[num5].Controls.Add(LB4b);
                    Main_TabPage.TabPages[num5].Controls.Add(LB5b);
                    Main_TabPage.TabPages[num5].Controls.Add(LB6b);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC2_NS);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC_NS);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC2_EW);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC_EW);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC2_NENW);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC_NENW);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC2_SESW);
                    Main_TabPage.TabPages[num5].Controls.Add(TB_SC_SESW);
                    Main_TabPage.TabPages[num5].Controls.Add(LB7);
                    Main_TabPage.TabPages[num5].Controls.Add(TB7);
                    Main_TabPage.TabPages[num5].Controls.Add(LB8);
                    Main_TabPage.TabPages[num5].Controls.Add(TB8);
                    Main_TabPage.TabPages[num5].Controls.Add(LB9);
                    Main_TabPage.TabPages[num5].Controls.Add(TB9);
                    Main_TabPage.TabPages[num5].Controls.Add(LB10);
                    Main_TabPage.TabPages[num5].Controls.Add(TB10);
                    Main_TabPage.TabPages[num5].Controls.Add(LB11);
                    Main_TabPage.TabPages[num5].Controls.Add(TB11);
                    Main_TabPage.TabPages[num5].Controls.Add(LB12);
                    Main_TabPage.TabPages[num5].Controls.Add(TB12);
                    Main_TabPage.TabPages[num5].Controls.Add(btn);

                    ////////////////////////////////////////////////////////////////////////
                    btn.Click += new EventHandler(btnClick);

                    //CB1.SelectedIndexChanged += new System.EventHandler(cboSourceType_SelectedIndexChanged);

                    ////////////////////////////////////////////////////////////////////////
                    ////////////Dynamic Events for Shades//////////////////////////
                    
                    if (dataGrid_Wndw_Types.Rows[num1].Cells[4].Value.ToString() == "Horizontal Projection")
                    {
                        CB1.SelectedIndex = 1;
                        PB1.Location = new System.Drawing.Point(500, 100);
                        PB1.Width = 315;
                        PB1.Height = 435;
                        PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Hor_Pro_Wndw.png"); //"D:\WIP\01_API\ETTV\Hor_Pro_Wndw.png";
                        PB1.SizeMode = PictureBoxSizeMode.Zoom;

                        double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)),4);
                        TB11.Text = R1.ToString();

                        LB12.Enabled = false;
                        TB12.Enabled = false;

                        //////SC2 Calculations
                        //////NS
                        ///
                        double Agl = double.Parse(TB7.Text);
                        SC2_NS = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NS);
                        if (SC2_NS == 0)
                        {
                            SC2_NS = 1;
                        }
                        TB_SC2_NS.Text = SC2_NS.ToString("0.0000");
                        TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");

                        //Lst_SC2_HP_NS.Add(SC2_NS);


                        //////EW
                        SC2_EW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_EW);
                        if (SC2_EW == 0)
                        {
                            SC2_EW = 1;
                        }
                        TB_SC2_EW.Text = SC2_EW.ToString("0.0000");
                        TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");
                        //////NENW
                        SC2_NENW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NENW);
                        if (SC2_NENW == 0)
                        {
                            SC2_NENW = 1;
                        }
                        TB_SC2_NENW.Text = SC2_NENW.ToString("0.0000");
                        TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");
                        //////SESW
                        SC2_SESW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_SESW);
                        if (SC2_SESW == 0)
                        {
                            SC2_SESW = 1;
                        }
                        TB_SC2_SESW.Text = SC2_SESW.ToString("0.0000");
                        TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");


                        ///////// For North Facade Table
                        ///summary table
                        num10 = (dataGridView2.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView2.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView2.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView2.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView2.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView2.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView2.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView2.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGridView1.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView1.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView1.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView1.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView1.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView1.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South Facade Table
                        ///summary table
                        num10 = (dataGridView3.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView3.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGridView3.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView3.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView3.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView3.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView3.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView3.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGridView4.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView4.Rows[num20].Cells[0].Value).ToString())
                            {
                                
                                dataGridView4.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView4.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView4.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView4.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num20].Cells[0].Value).ToString())
                            {                               
                                dataGrid_Wndw_E_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_E_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num20].Cells[0].Value).ToString())
                            {                               
                                dataGrid_Wndw_W_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num20].Cells[0].Value).ToString())
                            {                               
                                dataGrid_Wndw_W_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_W_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;                           
                        }                        
                        ///detail table
                        num10 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                        num20 = 0;                       
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                        num20 = 0;                        
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }                        
                        ///detail table
                        num10 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num20].Cells[0].Value).ToString())
                            {                                
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                    }
                    else if (dataGrid_Wndw_Types.Rows[num1].Cells[4].Value.ToString() == "Vertical Projection")
                    {
                        CB1.SelectedIndex = 2;
                        PB1.Location = new System.Drawing.Point(365, 100);
                        PB1.Width = 575;
                        PB1.Height = 280;
                        PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Ver_Pro_Wndw.png"); //"D:\WIP\01_API\ETTV\Ver_Pro_Wndw.png";
                        PB1.SizeMode = PictureBoxSizeMode.Zoom;

                        double R2 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB10.Text)), 4);
                        TB12.Text = R2.ToString();

                        LB11.Enabled = false;
                        TB11.Enabled = false;

                        //////SC2 Calculations
                        //////NS
                        double Agl_VP = double.Parse(TB7.Text);
                        SC2_VP_NS = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_NS);
                        if (SC2_VP_NS == 0)
                        {
                            SC2_VP_NS = 1;
                        }
                        TB_SC2_NS.Text = SC2_VP_NS.ToString("0.0000");
                        TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");
                        //////EW
                        SC2_VP_EW = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_EW);
                        if (SC2_VP_EW == 0)
                        {
                            SC2_VP_EW = 1;
                        }
                        TB_SC2_EW.Text = SC2_VP_EW.ToString("0.0000");
                        TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");
                        //////NENW
                        SC2_VP_NENW = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_NENW);
                        if (SC2_VP_NENW == 0)
                        {
                            SC2_VP_NENW = 1;
                        }
                        TB_SC2_NENW.Text = SC2_VP_NENW.ToString("0.0000");
                        TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");
                        //////SESW
                        SC2_VP_SESW = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_SESW);
                        if (SC2_VP_SESW == 0)
                        {
                            SC2_VP_SESW = 1;
                        }
                        TB_SC2_SESW.Text = SC2_VP_SESW.ToString("0.0000");
                        TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");

                        ///////// For North Facade Table
                        ///summary table
                        num10 = (dataGridView2.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView2.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView2.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView2.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView2.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView2.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView2.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView2.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGridView1.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView1.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView1.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView1.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView1.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView1.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South Facade Table
                        ///summary table
                        num10 = (dataGridView3.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView3.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView3.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView3.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView3.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView3.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView3.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView3.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGridView4.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView4.Rows[num20].Cells[0].Value).ToString())
                            {

                                dataGridView4.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView4.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView4.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView4.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_E_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_W_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_W_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_W_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }


                    }
                    else if (dataGrid_Wndw_Types.Rows[num1].Cells[4].Value.ToString() == "Egg Crate Window")
                    {
                        CB1.SelectedIndex = 3;
                        PB1.Location = new System.Drawing.Point(400, 100);
                        PB1.Width = 480;
                        PB1.Height = 410;
                        PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Egg_Crt_Wndw.png");  //D:\WIP\01_API\ETTV\Egg_Crt_Wndw.png";
                        PB1.SizeMode = PictureBoxSizeMode.Zoom;

                        double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)), 4);
                        TB11.Text = R1.ToString();
                        double R2 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB10.Text)), 4);
                        TB12.Text = R2.ToString();

                        double Agl_EC = double.Parse(TB7.Text);
                        double b1 = 0;
                        double b2 = 0;
                        SC2_EC_NS_b1 = new double();
                        SC2_EC_NS_b2 = new double();
                        SC2_EC_EW_b1 = new double();
                        SC2_EC_EW_b2 = new double();
                        SC2_EC_NENW_b1 = new double();
                        SC2_EC_NENW_b2 = new double();
                        SC2_EC_SESW_b1 = new double();
                        SC2_EC_SESW_b2 = new double();

                        List<double> R1_EC_Lst = new List<double>
                        {
                            0.2,0.4,0.6,0.8,1.0,1.2,1.4,1.6,1.8
                        };

                        foreach (double b in R1_EC_Lst)
                        {
                            if (R1 < 0.2)
                            {
                                b1 = 0.2;
                                b2 = 0.2;
                                break;
                            }
                            if (R1 >= 1.8)
                            {
                                b1 = 1.8;
                                b2 = 1.8;
                                break;
                            }
                            else if (R1 == b)
                            {
                                b1 = b;
                                b2 = b;
                                break;
                            }
                            else if (R1 > 0.2 && b > R1)
                            {
                                b2 = b;
                                b1 = Math.Round((b - 0.2), 1);
                                break;
                            }

                        }

                        ////////////////North-South SC2_EC  
                        ////for b1
                        if (b1 == 0.2)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_0);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.4)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_1);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.6)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_2);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.8)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_3);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.0)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_4);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.2)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_5);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.4)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_6);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.6)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_7);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.8)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_8);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }                            
                        }
                        ////for b2
                        if (b2 == 0.2)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_0);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.4)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_1);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.6)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_2);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.8)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_3);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.0)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_4);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.2)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_5);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.4)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_6);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.6)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_7);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.8)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_8);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }                            
                        }

                        if (SC2_EC_NS_b1 == SC2_EC_NS_b2)
                        {
                            SC2_EC_NS = SC2_EC_NS_b1;
                        }
                        else
                        {
                            SC2_EC_NS = SC2_EC_NS_b1 + ((SC2_EC_NS_b2 - SC2_EC_NS_b1) * ((R1 - b1) / (b2 - b1)));
                        }                       

                        TB_SC2_NS.Text = SC2_EC_NS.ToString("0.0000");
                        TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");

                        ////////////////East-West SC2_EC  
                        ////for b1
                        if (b1 == 0.2)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_0);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                           
                        }
                        else if (b1 == 0.4)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_1);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.6)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_2);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.8)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_3);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.0)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_4);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.2)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_5);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.4)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_6);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.6)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_7);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.8)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_8);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }                            
                        }
                        ////for b2
                        if (b2 == 0.2)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_0);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.4)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_1);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.6)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_2);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.8)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_3);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.0)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_4);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.2)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_5);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.4)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_6);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.6)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_7);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.8)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_8);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }                            
                        }

                        if (SC2_EC_EW_b1 == SC2_EC_EW_b2)
                        {
                            SC2_EC_EW = SC2_EC_EW_b1;
                        }
                        else
                        {
                            SC2_EC_EW = SC2_EC_EW_b1 + ((SC2_EC_EW_b2 - SC2_EC_EW_b1) * ((R1 - b1) / (b2 - b1)));
                        }                        

                        TB_SC2_EW.Text = SC2_EC_EW.ToString("0.0000");
                        TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");

                        ////////////////NorthEast-NorthWest SC2_EC  
                        ////for b1
                        if (b1 == 0.2)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_0);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.4)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_1);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.6)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_2);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                           
                        }
                        else if (b1 == 0.8)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_3);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.0)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_4);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.2)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_5);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.4)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_6);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.6)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_7);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.8)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_8);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }                            
                        }
                        ////for b2
                        if (b2 == 0.2)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_0);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.4)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_1);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.6)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_2);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.8)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_3);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.0)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_4);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.2)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_5);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.4)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_6);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.6)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_7);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.8)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_8);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }                            
                        }

                        if (SC2_EC_NENW_b1 == SC2_EC_NENW_b2)
                        {
                            SC2_EC_NENW = SC2_EC_NENW_b1;
                        }
                        else
                        {
                            SC2_EC_NENW = SC2_EC_NENW_b1 + ((SC2_EC_NENW_b2 - SC2_EC_NENW_b1) * ((R1 - b1) / (b2 - b1)));
                        }                        
                        TB_SC2_NENW.Text = SC2_EC_NENW.ToString("0.0000");
                        TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");

                        ////////////////SouthEast-SouthWest SC2_EC  
                        ////for b1
                        if (b1 == 0.2)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_0);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.4)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_1);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.6)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_2);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 0.8)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_3);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.0)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_4);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.2)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_5);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.4)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_6);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        else if (b1 == 1.6)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_7);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                           
                        }
                        else if (b1 == 1.8)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_8);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }                            
                        }
                        ////for b2
                        if (b2 == 0.2)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_0);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                           
                        }
                        else if (b2 == 0.4)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_1);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.6)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_2);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 0.8)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_3);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.0)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_4);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.2)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_5);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.4)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_6);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.6)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_7);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }
                        else if (b2 == 1.8)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_8);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }                            
                        }

                        if (SC2_EC_SESW_b1 == SC2_EC_SESW_b2)
                        {
                            SC2_EC_SESW = SC2_EC_SESW_b1;
                        }
                        else
                        {
                            SC2_EC_SESW = SC2_EC_SESW_b1 + ((SC2_EC_SESW_b2 - SC2_EC_SESW_b1) * ((R1 - b1) / (b2 - b1)));
                        }
                        
                        TB_SC2_SESW.Text = SC2_EC_SESW.ToString("0.0000");
                        TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");

                        ///////// For North Facade Table
                        ///summary table
                        num10 = (dataGridView2.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView2.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView2.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView2.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView2.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView2.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView2.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView2.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGridView1.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView1.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView1.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView1.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView1.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView1.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South Facade Table
                        ///summary table
                        num10 = (dataGridView3.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView3.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView3.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView3.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView3.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView3.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView3.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView3.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGridView4.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView4.Rows[num20].Cells[0].Value).ToString())
                            {

                                dataGridView4.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView4.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView4.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView4.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_E_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_W_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_W_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_W_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                    }
                    else
                    {
                        CB1.SelectedIndex = 0;
                        PB1.Location = new System.Drawing.Point(500, 100);
                        PB1.Width = 315;
                        PB1.Height = 435;
                        PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Hor_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Hor_Pro_Wndw.png";
                        PB1.SizeMode = PictureBoxSizeMode.Zoom;

                        TB8.Text = "0";
                        double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)), 4);
                        TB11.Text = R1.ToString();
                        double R2 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB10.Text)), 4);
                        TB12.Text = R2.ToString();

                        //////SC2 Calculations
                        //////NS
                        TB_SC2_NS.Text = "1";
                        TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");
                        //////EW
                        TB_SC2_EW.Text = "1";
                        TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");
                        //////NENW
                        TB_SC2_NENW.Text = "1";
                        TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");
                        //////SESW
                        TB_SC2_SESW.Text = "1";
                        TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");

                        ///////// For North Facade Table
                        ///summary table
                        num10 = (dataGridView2.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView2.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView2.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView2.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView2.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView2.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView2.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView2.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGridView1.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView1.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView1.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView1.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView1.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView1.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num20].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South Facade Table
                        ///summary table
                        num10 = (dataGridView3.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView3.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGridView3.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView3.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView3.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGridView3.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGridView3.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGridView3.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGridView4.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGridView4.Rows[num20].Cells[0].Value).ToString())
                            {

                                dataGridView4.Rows[num20].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                dataGridView4.Rows[num20].Cells[6].Value = TB_SC_NS.Text; //SC
                                dataGridView4.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGridView4.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num20].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_E_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_E_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_E_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num20].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_W_S.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_W_D.Rows[num20].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value = TB_SC_EW.Text; //SC
                                dataGrid_Wndw_W_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_W_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num20].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num20].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For North West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value = TB_SC_NENW.Text; //SC
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_NW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num20].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South East Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SE_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num20].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }

                        ///////// For South West Facade Table
                        ///summary table
                        num10 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                        num20 = 0;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value);
                                TotalCalValue = TotalCalValue + CalValue;
                                dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value);
                                TotalCalValue1 = TotalCalValue1 + CalValue1;
                            }
                            num20 = num20 + 1;
                        }
                        ///detail table
                        num10 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                        num20 = 0;
                        while (num20 < num10)
                        {
                            if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num20].Cells[0].Value).ToString())
                            {
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value = TB_SC_SESW.Text; //SC
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                dataGrid_Wndw_SW_D.Rows[num20].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num20].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                            }
                            num20 = num20 + 1;
                        }


                    }

                   
                    ///////////////////////////////////////////////////////////////////////
                    ////////////////Button Click Events for Shades/////////////////

                    void btnClick(object sender, EventArgs e)
                    {
                        //Lst_SC2_HP_NS = new List<double>();

                        if (CB1.SelectedIndex==0) 
                        {
                            PB1.Location = new System.Drawing.Point(500, 100);
                            PB1.Width = 315;
                            PB1.Height = 435;
                            PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Hor_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Hor_Pro_Wndw.png";
                            PB1.SizeMode = PictureBoxSizeMode.Zoom;

                            LB12.Enabled = true;
                            TB12.Enabled = true;
                            LB11.Enabled = true;
                            TB11.Enabled = true;

                            TB8.Text = "0";
                            TB7.Text = "0";
                            double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)), 4);
                            TB11.Text = R1.ToString();
                            double R2 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB10.Text)), 4);
                            TB12.Text = R2.ToString();

                            //////SC2 Calculations
                            //////NS
                            TB_SC2_NS.Text = "1";
                            TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");
                            //////EW
                            TB_SC2_EW.Text = "1";
                            TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");
                            //////NENW
                            TB_SC2_NENW.Text = "1";
                            TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");
                            //////SESW
                            TB_SC2_SESW.Text = "1";
                            TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");

                            ////////// to update back the DataGrid Tables    
                            
                            ///////// For Window Types Table                      
                            foreach (DataGridViewRow row in dataGrid_Wndw_Types.Rows)
                            {
                                if (resultLabel4.Text == (row.Cells[0].Value).ToString())
                                {
                                    row.Cells[2].Value = TB2.Text; //UValue
                                    row.Cells[3].Value = TB1.Text; //SC1
                                    row.Cells[4].Value = CB1.Text; //Shades
                                    row.Cells[5].Value = "0"; //Angle
                                    row.Cells[6].Value = "0"; //P
                                    row.Cells[7].Value = TB9.Text; //H
                                    row.Cells[8].Value = TB10.Text; //W
                                    break;
                                }
                            }

                            ///////// For North Facade Table
                            ///summary table
                            num1 = (dataGridView2.Rows.Count)-1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView2.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView2.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView2.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView2.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView2.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView2.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView2.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView2.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView2.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView1.Rows.Count)-1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView1.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView1.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView1.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView1.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView1.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView1.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView1.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South Facade Table
                            ///summary table
                            num1 = (dataGridView3.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView3.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView3.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView3.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView3.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView3.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView3.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView3.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView3.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView3.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView4.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView4.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView4.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView4.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView4.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView4.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView4.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView4.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }                                                       

                            ///////// For West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                        } ////Shades==> None
                        else if (CB1.SelectedIndex == 1)
                        {
                            PB1.Location = new System.Drawing.Point(500, 100);
                            PB1.Width = 315;
                            PB1.Height = 435;
                            PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Hor_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Hor_Pro_Wndw.png";
                            PB1.SizeMode = PictureBoxSizeMode.Zoom;

                            LB12.Enabled = true;
                            TB12.Enabled = true;
                            LB11.Enabled = true;
                            TB11.Enabled = true;

                            double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)), 4);
                            TB11.Text = R1.ToString();

                            LB12.Enabled = false;
                            TB12.Enabled = false;

                            //////SC2 Calculations
                            //////NS
                            double Agl = double.Parse(TB7.Text);
                            SC2_NS = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NS);
                            if (SC2_NS == 0)
                            {
                                SC2_NS = 1;
                            }
                            TB_SC2_NS.Text = SC2_NS.ToString("0.0000");
                            TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");

                            //Lst_SC2_HP_NS.Add(SC2_NS);


                            //////EW
                            SC2_EW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_EW);
                            if (SC2_EW == 0)
                            {
                                SC2_EW = 1;
                            }
                            TB_SC2_EW.Text = SC2_EW.ToString("0.0000");
                            TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");
                            //////NENW
                            SC2_NENW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NENW);
                            if (SC2_NENW == 0)
                            {
                                SC2_NENW = 1;
                            }
                            TB_SC2_NENW.Text = SC2_NENW.ToString("0.0000");
                            TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");
                            //////SESW
                            SC2_SESW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_SESW);
                            if (SC2_SESW == 0)
                            {
                                SC2_SESW = 1;
                            }
                            TB_SC2_SESW.Text = SC2_SESW.ToString("0.0000");
                            TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");


                            ////////// to update back the Window Types Table                                                    
                            foreach (DataGridViewRow row in dataGrid_Wndw_Types.Rows)
                            {
                                if (resultLabel4.Text == (row.Cells[0].Value).ToString())
                                {
                                    row.Cells[2].Value = TB2.Text; //UValue
                                    row.Cells[3].Value = TB1.Text; //SC1
                                    row.Cells[4].Value = CB1.Text; //Shades
                                    row.Cells[5].Value = TB7.Text; //Angle
                                    row.Cells[6].Value = TB8.Text; //P
                                    row.Cells[7].Value = TB9.Text; //H
                                    row.Cells[8].Value = TB10.Text; //W
                                    break;
                                }
                            }

                            ///////// For North Facade Table
                            ///summary table
                            num1 = (dataGridView2.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView2.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView2.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView2.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView2.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView2.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView2.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView2.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView2.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView2.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView1.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView1.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView1.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView1.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView1.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView1.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView1.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView1.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South Facade Table
                            ///summary table
                            num1 = (dataGridView3.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView3.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView3.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView3.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView3.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView3.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView3.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView3.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView3.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView3.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView4.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView4.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView4.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView4.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView4.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView4.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView4.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView4.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                        } ////Shades==> Horizontal
                        else if (CB1.SelectedIndex == 2)
                        {
                            PB1.Location = new System.Drawing.Point(365, 100);
                            PB1.Width = 575;
                            PB1.Height = 280;
                            PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Ver_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Ver_Pro_Wndw.png";
                            PB1.SizeMode = PictureBoxSizeMode.Zoom;

                            LB12.Enabled = true;
                            TB12.Enabled = true;
                            LB11.Enabled = true;
                            TB11.Enabled = true;

                            double R2 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB10.Text)), 4);
                            TB12.Text = R2.ToString();

                            LB11.Enabled = false;
                            TB11.Enabled = false;

                            //////SC2 Calculations
                            //////NS
                            double Agl_VP = double.Parse(TB7.Text);
                            SC2_VP_NS = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_NS);
                            if (SC2_VP_NS == 0)
                            {
                                SC2_VP_NS = 1;
                            }
                            TB_SC2_NS.Text = SC2_VP_NS.ToString("0.0000");
                            TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");
                            //////EW
                            SC2_VP_EW = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_EW);
                            if (SC2_VP_EW == 0)
                            {
                                SC2_VP_EW = 1;
                            }
                            TB_SC2_EW.Text = SC2_VP_EW.ToString("0.0000");
                            TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");
                            //////NENW
                            SC2_VP_NENW = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_NENW);
                            if (SC2_VP_NENW == 0)
                            {
                                SC2_VP_NENW = 1;
                            }
                            TB_SC2_NENW.Text = SC2_VP_NENW.ToString("0.0000");
                            TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");
                            //////SESW
                            SC2_VP_SESW = Interpolation_for_SC2(Agl_VP, R2, dataGrid_SC2_VP_SESW);
                            if (SC2_VP_SESW == 0)
                            {
                                SC2_VP_SESW = 1;
                            }
                            TB_SC2_SESW.Text = SC2_VP_SESW.ToString("0.0000");
                            TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");

                            ////////// to update back the Window Types Table                                                    
                            foreach (DataGridViewRow row in dataGrid_Wndw_Types.Rows)
                            {
                                if (resultLabel4.Text == (row.Cells[0].Value).ToString())
                                {
                                    row.Cells[2].Value = TB2.Text; //UValue
                                    row.Cells[3].Value = TB1.Text; //SC1
                                    row.Cells[4].Value = CB1.Text; //Shades
                                    row.Cells[5].Value = TB7.Text; //Angle
                                    row.Cells[6].Value = TB8.Text; //P
                                    row.Cells[7].Value = TB9.Text; //H
                                    row.Cells[8].Value = TB10.Text; //W
                                    break;
                                }
                            }

                            ///////// For North Facade Table
                            ///summary table
                            num1 = (dataGridView2.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView2.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView2.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView2.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView2.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView2.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView2.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView2.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView2.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView2.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView1.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView1.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView1.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView1.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView1.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView1.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView1.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView1.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South Facade Table
                            ///summary table
                            num1 = (dataGridView3.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView3.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView3.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView3.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView3.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView3.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView3.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView3.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView3.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView3.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView4.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView4.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView4.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView4.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView4.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView4.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView4.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView4.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }



                        } ////Shades==> Vertical
                        else if (CB1.SelectedIndex == 3)
                        {
                            PB1.Location = new System.Drawing.Point(400, 100);
                            PB1.Width = 480;
                            PB1.Height = 410;
                            PB1.ImageLocation = BuildingCoder.Util.GetFilePath("Egg_Crt_Wndw.png");  //"D:\WIP\01_API\ETTV\Egg_Crt_Wndw.png";
                            PB1.SizeMode = PictureBoxSizeMode.Zoom;

                            LB12.Enabled = true;
                            TB12.Enabled = true;
                            LB11.Enabled = true;
                            TB11.Enabled = true;

                            double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)), 4);
                            TB11.Text = R1.ToString();
                            double R2 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB10.Text)), 4);
                            TB12.Text = R2.ToString();

                            double Agl_EC = double.Parse(TB7.Text);
                            double b1 = 0;
                            double b2 = 0;
                            SC2_EC_NS_b1 = new double();
                            SC2_EC_NS_b2 = new double();
                            SC2_EC_EW_b1 = new double();
                            SC2_EC_EW_b2 = new double();
                            SC2_EC_NENW_b1 = new double();
                            SC2_EC_NENW_b2 = new double();
                            SC2_EC_SESW_b1 = new double();
                            SC2_EC_SESW_b2 = new double();

                            List<double> R1_EC_Lst = new List<double>
                        {
                            0.2,0.4,0.6,0.8,1.0,1.2,1.4,1.6,1.8
                        };

                            foreach (double b in R1_EC_Lst)
                            {
                                if (R1 < 0.2)
                                {
                                    b1 = 0.2;
                                    b2 = 0.2;
                                    break;
                                }
                                if (R1 >= 1.8)
                                {
                                    b1 = 1.8;
                                    b2 = 1.8;
                                    break;
                                }
                                else if (R1 == b)
                                {
                                    b1 = b;
                                    b2 = b;
                                    break;
                                }
                                else if (R1 > 0.2 && b > R1)
                                {
                                    b2 = b;
                                    b1 = Math.Round((b - 0.2), 1);
                                    break;
                                }

                            }

                            ////////////////North-South SC2_EC  
                            ////for b1
                            if (b1 == 0.2)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_0);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 0.4)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_1);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 0.6)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_2);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 0.8)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_3);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 1.0)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_4);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 1.2)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_5);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 1.4)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_6);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 1.6)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_7);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            else if (b1 == 1.8)
                            {
                                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_8);
                                if (SC2_EC_NS_b1 == 0)
                                {
                                    SC2_EC_NS_b1 = 1;
                                }
                            }
                            ////for b2
                            if (b2 == 0.2)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_0);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 0.4)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_1);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 0.6)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_2);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 0.8)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_3);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 1.0)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_4);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 1.2)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_5);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 1.4)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_6);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 1.6)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_7);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }
                            else if (b2 == 1.8)
                            {
                                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NS_8);
                                if (SC2_EC_NS_b2 == 0)
                                {
                                    SC2_EC_NS_b2 = 1;
                                }
                            }

                            if (SC2_EC_NS_b1 == SC2_EC_NS_b2)
                            {
                                SC2_EC_NS = SC2_EC_NS_b1;
                            }
                            else
                            {
                                SC2_EC_NS = SC2_EC_NS_b1 + ((SC2_EC_NS_b2 - SC2_EC_NS_b1) * ((R1 - b1) / (b2 - b1)));
                            }

                            TB_SC2_NS.Text = SC2_EC_NS.ToString("0.0000");
                            TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");

                            ////////////////East-West SC2_EC  
                            ////for b1
                            if (b1 == 0.2)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_0);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.4)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_1);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.6)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_2);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.8)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_3);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.0)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_4);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.2)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_5);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.4)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_6);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.6)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_7);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.8)
                            {
                                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_8);
                                if (SC2_EC_EW_b1 == 0)
                                {
                                    SC2_EC_EW_b1 = 1;
                                }
                            }
                            ////for b2
                            if (b2 == 0.2)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_0);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.4)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_1);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.6)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_2);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.8)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_3);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.0)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_4);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.2)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_5);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.4)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_6);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.6)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_7);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.8)
                            {
                                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_EW_8);
                                if (SC2_EC_EW_b2 == 0)
                                {
                                    SC2_EC_EW_b2 = 1;
                                }
                            }

                            if (SC2_EC_EW_b1 == SC2_EC_EW_b2)
                            {
                                SC2_EC_EW = SC2_EC_EW_b1;
                            }
                            else
                            {
                                SC2_EC_EW = SC2_EC_EW_b1 + ((SC2_EC_EW_b2 - SC2_EC_EW_b1) * ((R1 - b1) / (b2 - b1)));
                            }

                            TB_SC2_EW.Text = SC2_EC_EW.ToString("0.0000");
                            TB_SC_EW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_EW.Text)).ToString("0.0000");

                            ////////////////NorthEast-NorthWest SC2_EC  
                            ////for b1
                            if (b1 == 0.2)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_0);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.4)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_1);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.6)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_2);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.8)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_3);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.0)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_4);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.2)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_5);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.4)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_6);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.6)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_7);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.8)
                            {
                                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_8);
                                if (SC2_EC_NENW_b1 == 0)
                                {
                                    SC2_EC_NENW_b1 = 1;
                                }
                            }
                            ////for b2
                            if (b2 == 0.2)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_0);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.4)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_1);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.6)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_2);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.8)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_3);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.0)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_4);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.2)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_5);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.4)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_6);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.6)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_7);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.8)
                            {
                                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_NENW_8);
                                if (SC2_EC_NENW_b2 == 0)
                                {
                                    SC2_EC_NENW_b2 = 1;
                                }
                            }

                            if (SC2_EC_NENW_b1 == SC2_EC_NENW_b2)
                            {
                                SC2_EC_NENW = SC2_EC_NENW_b1;
                            }
                            else
                            {
                                SC2_EC_NENW = SC2_EC_NENW_b1 + ((SC2_EC_NENW_b2 - SC2_EC_NENW_b1) * ((R1 - b1) / (b2 - b1)));
                            }
                            TB_SC2_NENW.Text = SC2_EC_NENW.ToString("0.0000");
                            TB_SC_NENW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString("0.0000");

                            ////////////////SouthEast-SouthWest SC2_EC  
                            ////for b1
                            if (b1 == 0.2)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_0);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.4)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_1);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.6)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_2);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 0.8)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_3);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.0)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_4);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.2)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_5);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.4)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_6);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.6)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_7);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            else if (b1 == 1.8)
                            {
                                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_8);
                                if (SC2_EC_SESW_b1 == 0)
                                {
                                    SC2_EC_SESW_b1 = 1;
                                }
                            }
                            ////for b2
                            if (b2 == 0.2)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_0);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.4)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_1);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.6)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_2);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 0.8)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_3);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.0)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_4);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.2)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_5);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.4)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_6);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.6)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_7);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }
                            else if (b2 == 1.8)
                            {
                                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2, dataGrid_SC2_EC_SESW_8);
                                if (SC2_EC_SESW_b2 == 0)
                                {
                                    SC2_EC_SESW_b2 = 1;
                                }
                            }

                            if (SC2_EC_SESW_b1 == SC2_EC_SESW_b2)
                            {
                                SC2_EC_SESW = SC2_EC_SESW_b1;
                            }
                            else
                            {
                                SC2_EC_SESW = SC2_EC_SESW_b1 + ((SC2_EC_SESW_b2 - SC2_EC_SESW_b1) * ((R1 - b1) / (b2 - b1)));
                            }

                            TB_SC2_SESW.Text = SC2_EC_SESW.ToString("0.0000");
                            TB_SC_SESW.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString("0.0000");


                            ////////// to update back the Window Types Table                                                    
                            foreach (DataGridViewRow row in dataGrid_Wndw_Types.Rows)
                            {
                                if (resultLabel4.Text == (row.Cells[0].Value).ToString())
                                {
                                    row.Cells[2].Value = TB2.Text; //UValue
                                    row.Cells[3].Value = TB1.Text; //SC1
                                    row.Cells[4].Value = CB1.Text; //Shades
                                    row.Cells[5].Value = TB7.Text; //Angle
                                    row.Cells[6].Value = TB8.Text; //P
                                    row.Cells[7].Value = TB9.Text; //H
                                    row.Cells[8].Value = TB10.Text; //W
                                    break;
                                }
                            }

                            ///////// For North Facade Table
                            ///summary table
                            num1 = (dataGridView2.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView2.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView2.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView2.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView2.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView2.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView2.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView2.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView2.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView2.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView2.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView2.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView1.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView1.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView1.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView1.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView1.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView1.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView1.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView1.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView1.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView1.Rows[num2].Cells[6].Value)) * 0.8).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South Facade Table
                            ///summary table
                            num1 = (dataGridView3.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView3.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView3.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView3.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView3.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView3.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView3.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGridView3.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGridView3.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView3.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView3.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGridView3.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGridView4.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGridView4.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGridView4.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGridView4.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGridView4.Rows[num2].Cells[5].Value = TB_SC2_NS.Text; //SC2
                                    dataGridView4.Rows[num2].Cells[6].Value = TB_SC_NS.Text; //SC
                                    dataGridView4.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGridView4.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGridView4.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGridView4.Rows[num2].Cells[6].Value)) * 0.83).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_E_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_E_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_E_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_E_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_E_D.Rows[num2].Cells[6].Value)) * 1.13).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_W_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_W_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_W_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[5].Value = TB_SC2_EW.Text; //SC2
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value = TB_SC_EW.Text; //SC
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_W_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_W_D.Rows[num2].Cells[6].Value)) * 1.23).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NE_D.Rows[num2].Cells[6].Value)) * 0.97).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For North West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_NW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_NW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_NW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[5].Value = TB_SC2_NENW.Text; //SC2
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value = TB_SC_NENW.Text; //SC
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_NW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_NW_D.Rows[num2].Cells[6].Value)) * 1.03).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South East Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SE_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SE_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SE_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SE_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SE_D.Rows[num2].Cells[6].Value)) * 0.98).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }

                            ///////// For South West Facade Table
                            ///summary table
                            num1 = (dataGrid_Wndw_SW_S.Rows.Count) - 1;
                            num2 = 0;
                            CalValue = 0;
                            TotalCalValue = 0;
                            CalValue1 = 0;
                            TotalCalValue1 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_S.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[7].Value);
                                    TotalCalValue = TotalCalValue + CalValue;
                                    dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                    CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num2].Cells[8].Value);
                                    TotalCalValue1 = TotalCalValue1 + CalValue1;
                                }
                                num2 = num2 + 1;
                            }
                            ///detail table
                            num1 = (dataGrid_Wndw_SW_D.Rows.Count) - 1;
                            num2 = 0;
                            while (num2 < num1)
                            {
                                if (resultLabel4.Text == (dataGrid_Wndw_SW_D.Rows[num2].Cells[0].Value).ToString())
                                {
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value = TB2.Text; //UValue
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[4].Value = TB1.Text; //SC1
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[5].Value = TB_SC2_SESW.Text; //SC2
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value = TB_SC_SESW.Text; //SC
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[7].Value = (3.4 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[3].Value))).ToString("0.00");//Wndw Conduction HG
                                    dataGrid_Wndw_SW_D.Rows[num2].Cells[8].Value = (211 * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[2].Value)) * (double.Parse((string)dataGrid_Wndw_SW_D.Rows[num2].Cells[6].Value)) * 1.06).ToString("0.00");//Wndw Radiation HG
                                }
                                num2 = num2 + 1;
                            }


                        } ////Shades==> Egg Crate

                        /////////////////////////////////////Summation of Heat Gains////////////////////////////////////////////////         
                        ///////// For North Facade Summary Table                    
                        num10 = (dataGridView2.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGridView2.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGridView2.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGridView2.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGridView2.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGridView2.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_N.Text = (N_Wall_HG + (double.Parse((string)dataGridView2.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGridView2.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_N.Text = (double.Parse(Lb_HG_N.Text) / double.Parse(Lb_Area_N.Text)).ToString("0.00");
                        Lb_ETTVRes_N.Text = Lb_ETTV_N.Text;

                        ///////// For South Facade Summary Table                    
                        num10 = (dataGridView3.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGridView3.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGridView3.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGridView3.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGridView3.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGridView3.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_S.Text = (S_Wall_HG + (double.Parse((string)dataGridView3.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGridView3.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_S.Text = (double.Parse(Lb_HG_S.Text) / double.Parse(Lb_Area_S.Text)).ToString("0.00");
                        Lb_ETTVRes_S.Text = Lb_ETTV_S.Text;

                        ///////// For East Facade Summary Table                    
                        num10 = (dataGrid_Wndw_E_S.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGrid_Wndw_E_S.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGrid_Wndw_E_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGrid_Wndw_E_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_E.Text = (E_Wall_HG + (double.Parse((string)dataGrid_Wndw_E_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_E_S.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_E.Text = (double.Parse(Lb_HG_E.Text) / double.Parse(Lb_Area_E.Text)).ToString("0.00");
                        Lb_ETTVRes_E.Text = Lb_ETTV_E.Text;

                        ///////// For West Facade Summary Table                    
                        num10 = (dataGrid_Wndw_W_S.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGrid_Wndw_W_S.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGrid_Wndw_W_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGrid_Wndw_W_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_W.Text = (W_Wall_HG + (double.Parse((string)dataGrid_Wndw_W_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_W_S.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_W.Text = (double.Parse(Lb_HG_W.Text) / double.Parse(Lb_Area_W.Text)).ToString("0.00");
                        Lb_ETTVRes_W.Text = Lb_ETTV_W.Text;

                        ///////// For NorthEast Facade Summary Table                    
                        num10 = (dataGrid_Wndw_NE_S.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGrid_Wndw_NE_S.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGrid_Wndw_NE_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGrid_Wndw_NE_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_NE.Text = (NE_Wall_HG + (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_NE.Text = (double.Parse(Lb_HG_NE.Text) / double.Parse(Lb_Area_NE.Text)).ToString("0.00");
                        Lb_ETTVRes_NE.Text = Lb_ETTV_NE.Text;

                        ///////// For NorthWest Facade Summary Table                    
                        num10 = (dataGrid_Wndw_NW_S.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGrid_Wndw_NW_S.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGrid_Wndw_NW_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGrid_Wndw_NW_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_NW.Text = (NW_Wall_HG + (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_NW.Text = (double.Parse(Lb_HG_NW.Text) / double.Parse(Lb_Area_NW.Text)).ToString("0.00");
                        Lb_ETTVRes_NW.Text = Lb_ETTV_NW.Text;

                        ///////// For SouthEast Facade Summary Table                    
                        num10 = (dataGrid_Wndw_SE_S.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGrid_Wndw_SE_S.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGrid_Wndw_SE_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGrid_Wndw_SE_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_SE.Text = (SE_Wall_HG + (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_SE.Text = (double.Parse(Lb_HG_SE.Text) / double.Parse(Lb_Area_SE.Text)).ToString("0.00");
                        Lb_ETTVRes_SE.Text = Lb_ETTV_SE.Text;

                        ///////// For SouthWest Facade Summary Table                    
                        num10 = (dataGrid_Wndw_SW_S.Rows.Count) - 2;
                        num20 = 0;
                        num111 = (dataGrid_Wndw_SW_S.Rows.Count) - 2;
                        CalValue = 0;
                        TotalCalValue = 0;
                        CalValue1 = 0;
                        TotalCalValue1 = 0;
                        while (num20 < num10)
                        {
                            CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value);
                            TotalCalValue = TotalCalValue + CalValue;
                            CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value);
                            TotalCalValue1 = TotalCalValue1 + CalValue1;
                            num20 = num20 + 1;
                        }
                        dataGrid_Wndw_SW_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
                        dataGrid_Wndw_SW_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
                        Lb_HG_SW.Text = (SW_Wall_HG + (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num111].Cells[8].Value))).ToString("0.00");
                        Lb_ETTV_SW.Text = (double.Parse(Lb_HG_SW.Text) / double.Parse(Lb_Area_SW.Text)).ToString("0.00");
                        Lb_ETTVRes_SW.Text = Lb_ETTV_SW.Text;

                        ///////// For Average ETTV //////////////
                        Total_HG = double.Parse(Lb_HG_N.Text) + double.Parse(Lb_HG_S.Text) + double.Parse(Lb_HG_E.Text) + double.Parse(Lb_HG_W.Text)
                            + double.Parse(Lb_HG_NE.Text) + double.Parse(Lb_HG_NW.Text) + double.Parse(Lb_HG_SE.Text) + double.Parse(Lb_HG_SW.Text);
                        Total_Area = double.Parse(Lb_Area_N.Text) + double.Parse(Lb_Area_S.Text) + double.Parse(Lb_Area_E.Text) + double.Parse(Lb_Area_W.Text)
                            + double.Parse(Lb_Area_NE.Text) + double.Parse(Lb_Area_NW.Text) + double.Parse(Lb_Area_SE.Text) + double.Parse(Lb_Area_SW.Text);
                        Lb_ETTV_Avg.Text = (Total_HG / Total_Area).ToString("0.00");

                    }                                     
                   
                    //////////////////////////////////////////////////////////////////////////////////////////////////                    

                    num1 = num1 + 1;
                }
                num5 = num5 + 1;
                
            }

            /////////////////////////////////////Summation of Heat Gains////////////////////////////////////////////////         
            ///////// For North Facade Summary Table                    
            num10 = (dataGridView2.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGridView2.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGridView2.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGridView2.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGridView2.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGridView2.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_N.Text = (N_Wall_HG + (double.Parse((string)dataGridView2.Rows[num111].Cells[7].Value))+ (double.Parse((string)dataGridView2.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_N.Text = (double.Parse(Lb_HG_N.Text) / double.Parse(Lb_Area_N.Text)).ToString("0.00");
            Lb_ETTVRes_N.Text = Lb_ETTV_N.Text;

            ///////// For South Facade Summary Table                    
            num10 = (dataGridView3.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGridView3.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGridView3.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGridView3.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGridView3.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGridView3.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_S.Text = (S_Wall_HG + (double.Parse((string)dataGridView3.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGridView3.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_S.Text = (double.Parse(Lb_HG_S.Text) / double.Parse(Lb_Area_S.Text)).ToString("0.00");
            Lb_ETTVRes_S.Text = Lb_ETTV_S.Text;

            ///////// For East Facade Summary Table                    
            num10 = (dataGrid_Wndw_E_S.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGrid_Wndw_E_S.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGrid_Wndw_E_S.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGrid_Wndw_E_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGrid_Wndw_E_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_E.Text = (E_Wall_HG + (double.Parse((string)dataGrid_Wndw_E_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_E_S.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_E.Text = (double.Parse(Lb_HG_E.Text) / double.Parse(Lb_Area_E.Text)).ToString("0.00");
            Lb_ETTVRes_E.Text = Lb_ETTV_E.Text;

            ///////// For West Facade Summary Table                    
            num10 = (dataGrid_Wndw_W_S.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGrid_Wndw_W_S.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGrid_Wndw_W_S.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGrid_Wndw_W_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGrid_Wndw_W_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_W.Text = (W_Wall_HG + (double.Parse((string)dataGrid_Wndw_W_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_W_S.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_W.Text = (double.Parse(Lb_HG_W.Text) / double.Parse(Lb_Area_W.Text)).ToString("0.00");
            Lb_ETTVRes_W.Text = Lb_ETTV_W.Text;

            ///////// For NorthEast Facade Summary Table                    
            num10 = (dataGrid_Wndw_NE_S.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGrid_Wndw_NE_S.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGrid_Wndw_NE_S.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGrid_Wndw_NE_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGrid_Wndw_NE_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_NE.Text = (NE_Wall_HG + (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_NE_S.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_NE.Text = (double.Parse(Lb_HG_NE.Text) / double.Parse(Lb_Area_NE.Text)).ToString("0.00");
            Lb_ETTVRes_NE.Text = Lb_ETTV_NE.Text;

            ///////// For NorthWest Facade Summary Table                    
            num10 = (dataGrid_Wndw_NW_S.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGrid_Wndw_NW_S.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGrid_Wndw_NW_S.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGrid_Wndw_NW_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGrid_Wndw_NW_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_NW.Text = (NW_Wall_HG + (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_NW_S.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_NW.Text = (double.Parse(Lb_HG_NW.Text) / double.Parse(Lb_Area_NW.Text)).ToString("0.00");
            Lb_ETTVRes_NW.Text = Lb_ETTV_NW.Text;

            ///////// For SouthEast Facade Summary Table                    
            num10 = (dataGrid_Wndw_SE_S.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGrid_Wndw_SE_S.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGrid_Wndw_SE_S.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGrid_Wndw_SE_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGrid_Wndw_SE_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_SE.Text = (SE_Wall_HG + (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_SE_S.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_SE.Text = (double.Parse(Lb_HG_SE.Text) / double.Parse(Lb_Area_SE.Text)).ToString("0.00");
            Lb_ETTVRes_SE.Text = Lb_ETTV_SE.Text;

            ///////// For SouthWest Facade Summary Table                    
            num10 = (dataGrid_Wndw_SW_S.Rows.Count) - 2;
            num20 = 0;
            num111 = (dataGrid_Wndw_SW_S.Rows.Count) - 2;
            CalValue = 0;
            TotalCalValue = 0;
            CalValue1 = 0;
            TotalCalValue1 = 0;
            while (num20 < num10)
            {
                CalValue = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[7].Value);
                TotalCalValue = TotalCalValue + CalValue;
                CalValue1 = double.Parse((string)dataGrid_Wndw_SW_S.Rows[num20].Cells[8].Value);
                TotalCalValue1 = TotalCalValue1 + CalValue1;
                num20 = num20 + 1;
            }
            dataGrid_Wndw_SW_S.Rows[num111].Cells[7].Value = (TotalCalValue).ToString();
            dataGrid_Wndw_SW_S.Rows[num111].Cells[8].Value = (TotalCalValue1).ToString();
            Lb_HG_SW.Text = (SW_Wall_HG + (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num111].Cells[7].Value)) + (double.Parse((string)dataGrid_Wndw_SW_S.Rows[num111].Cells[8].Value))).ToString("0.00");
            Lb_ETTV_SW.Text = (double.Parse(Lb_HG_SW.Text) / double.Parse(Lb_Area_SW.Text)).ToString("0.00");
            Lb_ETTVRes_SW.Text = Lb_ETTV_SW.Text;

            ///////// For Average ETTV //////////////
            Total_HG = double.Parse(Lb_HG_N.Text) + double.Parse(Lb_HG_S.Text) + double.Parse(Lb_HG_E.Text) + double.Parse(Lb_HG_W.Text) 
                + double.Parse(Lb_HG_NE.Text) + double.Parse(Lb_HG_NW.Text) + double.Parse(Lb_HG_SE.Text) + double.Parse(Lb_HG_SW.Text);
            Total_Area = double.Parse(Lb_Area_N.Text) + double.Parse(Lb_Area_S.Text) + double.Parse(Lb_Area_E.Text) + double.Parse(Lb_Area_W.Text)
                + double.Parse(Lb_Area_NE.Text) + double.Parse(Lb_Area_NW.Text) + double.Parse(Lb_Area_SE.Text) + double.Parse(Lb_Area_SW.Text);
            Lb_ETTV_Avg.Text = (Total_HG / Total_Area).ToString("0.00");

        }

        private void ETTV_F2_Load(object sender, EventArgs e)
        {

            //tabNorth.BackColor = System.Drawing.Color.Gray;
            //tabSouth.BackColor = System.Drawing.Color.AliceBlue;
            //tabEast.BackColor = System.Drawing.Color.AliceBlue;
            //tabWest.BackColor = System.Drawing.Color.AliceBlue;
            //tabNorthEast.BackColor = System.Drawing.Color.AliceBlue;
            //tabNorthWest.BackColor = System.Drawing.Color.AliceBlue;           


            //For North Window
            button1.Enabled = false;
            button2.Enabled = true;
            dataGridView1.Visible = false;
            //For North Wall
            btn_Wall_N_S.Enabled = false;
            btn_Wall_N_D.Enabled = true;
            dataGrid_Wall_N_D.Visible = false;

            //For South Window
            button3.Enabled = false;
            button4.Enabled = true;
            dataGridView4.Visible = false;
            //For South Wall
            btn_Wall_S_S.Enabled = false;
            btn_Wall_S_D.Enabled = true;
            dataGrid_Wall_S_D.Visible = false;

            //For East Window
            button8.Enabled = false;
            button7.Enabled = true;
            dataGrid_Wndw_E_D.Visible = false;
            //For East Wall
            button6.Enabled = false;
            button5.Enabled = true;
            dataGrid_Wall_E_D.Visible = false;

            //For West Window
            button12.Enabled = false;
            button11.Enabled = true;
            dataGrid_Wndw_W_D.Visible = false;
            //For West Wall
            button10.Enabled = false;
            button9.Enabled = true;
            dataGrid_Wall_W_D.Visible = false;

            //For North East Window
            button16.Enabled = false;
            button15.Enabled = true;
            dataGrid_Wndw_NE_D.Visible = false;
            //For North East Wall
            button14.Enabled = false;
            button13.Enabled = true;
            dataGrid_Wall_NE_D.Visible = false;

            //For North West Window
            button20.Enabled = false;
            button19.Enabled = true;
            dataGrid_Wndw_NW_D.Visible = false;
            //For North West Wall
            button18.Enabled = false;
            button17.Enabled = true;
            dataGrid_Wall_NW_D.Visible = false;

            //For South East Window
            button24.Enabled = false;
            button23.Enabled = true;
            dataGrid_Wndw_SE_D.Visible = false;
            //For South East Wall
            button22.Enabled = false;
            button21.Enabled = true;
            dataGrid_Wall_SE_D.Visible = false;

            //For South West Window
            button28.Enabled = false;
            button27.Enabled = true;
            dataGrid_Wndw_SW_D.Visible = false;
            //For South West Wall
            button26.Enabled = false;
            button25.Enabled = true;
            dataGrid_Wall_SW_D.Visible = false;



            //////////////////////////// Window SC Calculations_Horizontal /////////////////////////////////////

            double R1 = double.Parse(TB_P.Text) / double.Parse(TB_H.Text);
            TB_R1.Text = R1.ToString("0.0000");
            double Agl = double.Parse(TB_Angle.Text);

            ////////////////////////// For North-South SC2 /////////////////////////////////
            SC2_NS = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NS);
            if (SC2_NS == 0)
            {
                SC2_NS = 1;
            }
            TB_SC2_NS.Text = SC2_NS.ToString("0.0000");
            TB_SC_NS.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_NS.Text)).ToString();

            ////////////////////////// For East-West SC2 /////////////////////////////////
            SC2_EW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_EW);
            if (SC2_EW == 0)
            {
                SC2_EW = 1;
            }
            TB_SC2_EW.Text = SC2_EW.ToString("0.0000");
            TB_SC_EW.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_EW.Text)).ToString();

            ////////////////////////// For NorthEast-NorthWest SC2 /////////////////////////////////
            SC2_NENW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NENW);
            if (SC2_NENW == 0)
            {
                SC2_NENW = 1;
            }
            TB_SC2_NENW.Text = SC2_NENW.ToString("0.0000");
            TB_SC_NENW.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString();

            ////////////////////////// For SouthEast-SouthWest SC2 /////////////////////////////////
            SC2_SESW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_SESW);
            if (SC2_SESW == 0)
            {
                SC2_SESW = 1;
            }
            TB_SC2_SESW.Text = SC2_SESW.ToString("0.0000");
            TB_SC_SESW.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString();

            //////////////////////////// Window SC Calculations_Horizontal /////////////////////////////////////

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////////////////////////// Window SC Calculations_Vertical /////////////////////////////////////

            double R1_VP = double.Parse(TB_P_VP.Text) / double.Parse(TB_H_VP.Text);
            TB_R1_VP.Text = R1_VP.ToString("0.0000");
            double Agl_VP = double.Parse(TB_Angle_VP.Text);

            ////////////////////////// For North-South SC2 /////////////////////////////////
            SC2_VP_NS = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_NS);
            if (SC2_VP_NS == 0)
            {
                SC2_VP_NS = 1;
            }
            TB_SC2_VP_NS.Text = SC2_VP_NS.ToString("0.0000");
            TB_SC_VP_NS.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_NS.Text)).ToString();

            ////////////////////////// For East-West SC2 /////////////////////////////////
            SC2_VP_EW = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_EW);
            if (SC2_VP_EW == 0)
            {
                SC2_VP_EW = 1;
            }
            TB_SC2_VP_EW.Text = SC2_VP_EW.ToString("0.0000");
            TB_SC_VP_EW.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_EW.Text)).ToString();

            ////////////////////////// For NorthEast-NorthWest SC2 /////////////////////////////////
            SC2_VP_NENW = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_NENW);
            if (SC2_VP_NENW == 0)
            {
                SC2_VP_NENW = 1;
            }
            TB_SC2_VP_NENW.Text = SC2_VP_NENW.ToString("0.0000");
            TB_SC_VP_NENW.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_NENW.Text)).ToString();

            ////////////////////////// For SouthEast-SouthWest SC2 /////////////////////////////////
            SC2_VP_SESW = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_SESW);
            if (SC2_VP_SESW == 0)
            {
                SC2_VP_SESW = 1;
            }
            TB_SC2_VP_SESW.Text = SC2_VP_SESW.ToString("0.0000");
            TB_SC_VP_SESW.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_SESW.Text)).ToString();


            //////////////////////////// Window SC Calculations_Vertical /////////////////////////////////////

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////////////////////////// Window SC Calculations_Egg Crate /////////////////////////////////////
            double Agl_EC = double.Parse(TB_Agl_EC.Text);
            double R1_EC = double.Parse(TB_P_EC.Text) / double.Parse(TB_H_EC.Text);
            TB_R1_EC.Text = R1_EC.ToString("0.0000");
            double R2_EC = double.Parse(TB_P_EC.Text) / double.Parse(TB_W_EC.Text);
            TB_R2_EC.Text = R2_EC.ToString("0.0000");

            double b1 = 0;
            double b2 = 0;
            SC2_EC_NS_b1 = new double();
            SC2_EC_NS_b2 = new double();
            SC2_EC_EW_b1 = new double();
            SC2_EC_EW_b2 = new double();
            SC2_EC_NENW_b1 = new double();
            SC2_EC_NENW_b2 = new double();
            SC2_EC_SESW_b1 = new double();
            SC2_EC_SESW_b2 = new double();

            List<double> R1_EC_Lst = new List<double>
            {
             0.2,0.4,0.6,0.8,1.0,1.2,1.4,1.6,1.8
            };

            foreach (double b in R1_EC_Lst)
            {
                if (R1_EC < 0.2)
                {
                    b1 = 0.2;
                    b2 = 0.2;
                    break;
                }
                if (R1_EC >= 1.8)
                {
                    b1 = 1.8;
                    b2 = 1.8;
                    break;
                }
                else if (R1_EC == b)
                {
                    b1 = b;
                    b2 = b;
                    break;
                }
                else if (R1_EC > 0.2 && b > R1_EC)
                {
                    b2 = b;
                    b1 = Math.Round((b - 0.2), 1);
                    break;
                }

            }


            ////////////////North-South SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_0);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_1);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_2);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_3);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_4);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_5);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_6);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_7);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_8);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_0);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_1);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_2);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_3);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_4);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_5);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_6);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_7);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_8);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }

            if (SC2_EC_NS_b1 == SC2_EC_NS_b2)
            {
                SC2_EC_NS = SC2_EC_NS_b1;
            }
            else
            {
                SC2_EC_NS = SC2_EC_NS_b1 + ((SC2_EC_NS_b2 - SC2_EC_NS_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_NS_EC.Text = SC2_EC_NS.ToString("0.0000");

            TB_SC2_NS_EC.Text = SC2_EC_NS.ToString("0.0000");
            TB_SC_NS_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_NS_EC.Text)).ToString("0.0000");


            ////////////////East-West SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_0);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_1);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_2);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_3);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_4);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_5);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_6);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_7);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_8);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_0);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_1);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_2);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_3);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_4);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_5);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_6);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_7);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_8);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }

            if (SC2_EC_EW_b1 == SC2_EC_EW_b2)
            {
                SC2_EC_EW = SC2_EC_EW_b1;
            }
            else
            {
                SC2_EC_EW = SC2_EC_EW_b1 + ((SC2_EC_EW_b2 - SC2_EC_EW_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_EW_EC.Text = SC2_EC_EW.ToString("0.0000");

            TB_SC2_EW_EC.Text = SC2_EC_EW.ToString("0.0000");
            TB_SC_EW_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_EW_EC.Text)).ToString("0.0000");

            ////////////////NorthEast-NorthWest SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_0);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_1);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_2);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_3);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_4);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_5);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_6);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_7);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_8);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_0);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_1);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_2);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_3);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_4);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_5);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_6);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_7);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_8);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }

            if (SC2_EC_NENW_b1 == SC2_EC_NENW_b2)
            {
                SC2_EC_NENW = SC2_EC_NENW_b1;
            }
            else
            {
                SC2_EC_NENW = SC2_EC_NENW_b1 + ((SC2_EC_NENW_b2 - SC2_EC_NENW_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_NENW_EC.Text = SC2_EC_NENW.ToString("0.0000");

            TB_SC2_NENW_EC.Text = SC2_EC_NENW.ToString("0.0000");
            TB_SC_NENW_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_NENW_EC.Text)).ToString("0.0000");


            ////////////////SouthEast-SouthWest SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_0);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_1);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_2);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_3);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_4);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_5);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_6);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_7);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_8);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_0);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_1);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_2);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_3);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_4);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_5);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_6);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_7);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_8);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }

            if (SC2_EC_SESW_b1 == SC2_EC_SESW_b2)
            {
                SC2_EC_SESW = SC2_EC_SESW_b1;
            }
            else
            {
                SC2_EC_SESW = SC2_EC_SESW_b1 + ((SC2_EC_SESW_b2 - SC2_EC_SESW_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_SESW_EC.Text = SC2_EC_SESW.ToString("0.0000");

            TB_SC2_SESW_EC.Text = SC2_EC_SESW.ToString("0.0000");
            TB_SC_SESW_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_SESW_EC.Text)).ToString("0.0000");

            //////////////////////////// Window SC Calculations_Egg Crate /////////////////////////////////////

            //////////////////////////// Creating KeySchedule for Window Shades //////////////////////////////////
            //using (Transaction trans = new Transaction(doc2, "ETTV_F2"))
            //{
                //trans.Start();

                // Create KeySchedule
               // shdl = ViewSchedule.CreateKeySchedule(doc2, new ElementId(BuiltInCategory.OST_Windows));
                               
               // trans.Commit();
            //}

            //shdlableFields = shdl.Definition.GetSchedulableFields();

            //using (Transaction trans = new Transaction(doc2, "ETTV_F2"))
            //{
               // trans.Start();

               // foreach (SchedulableField sf in shdlableFields)
                //{
                    //string SF_Srting = sf.GetName(doc2);
                    //if (SF_Srting == "Shades")
                    //{
                        //shdl.Definition.AddField(sf);                        
                   // }                  

               // }
              //  trans.Commit();
            //}

            //using (Transaction trans = new Transaction(doc2, "ETTV_F2"))
            //{
            //trans.Start();
            //shdlFieldIds.Add(shdl.Definition.GetFieldId(0));
            //shdlFieldIds.Add(shdl.Definition.GetFieldId(1));

            //shdl.Definition.RemoveField(shdlFieldIds[0]);
            //trans.Commit();
            //}

            //uidoc2.ActiveView = shdl;

            //using (Transaction trans = new Transaction(doc2, "ETTV_F2"))
            //{
                //trans.Start();

                //TableData tableData = shdl.GetTableData();
                //TableSectionData tsd = tableData.GetSectionData(SectionType.Body);
                //tsd.InsertRow(tsd.FirstRowNumber);
                //tsd.SetCellText(0, 0, "test");

                //tsd.InsertRow(tsd.FirstRowNumber + 1);


                //tsd.InsertRow(tsd.FirstRowNumber + 2);


                //tsd.InsertRow(tsd.FirstRowNumber + 3);
                // tsd.SetCellText(tsd.FirstRowNumber + 1, tsd.FirstColumnNumber, "Schedule of column top and base levels with offsets");
                //tsd.SetCellText(0, 0, "Text");
                //tsd.InsertRow(1);

                //tsd.InsertRow(2);

                //tsd.InsertRow(3);

               // trans.Commit();
            //}
           
            ////FilteredElementCollector wndw_collector = new FilteredElementCollector(doc2);
            //ElementCategoryFilter wndw_filter = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            //IList<Element> wndws = wndw_collector.WherePasses(wndw_filter).WhereElementIsNotElementType().ToList();

            //List<Element> elmts = (List<Element>)new FilteredElementCollector(doc2, shdl.Id).ToElements();
           // List<Parameter> paras = new List<Parameter>();

            //foreach (Element el in elmts)
            //{
                //ElementType WTYN = doc2.GetElement(el.GetTypeId()) as ElementType;

               // paras.Add(WTYN.LookupParameter("Shades"));

            //}
                       


            //foreach (Element el in elmts)
            //{
            // paras.Add(el.LookupParameter("Shades"));
            //}

            //using (Transaction trans = new Transaction(doc2, "ETTV_F2"))
            //{

            //trans.Start();

            //foreach (Element el in wndws)
            //{
            // el.LookupParameter("Shades").Set("Test");
            //}                 

            //trans.RollBack();
            //}

            //////////////////////////// Creating KeySchedule for Window Shades //////////////////////////////////

        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button2.Enabled = true;
            button1.Enabled = false;
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button4.Enabled = true;
            button3.Enabled = false;
            dataGridView3.Visible = true;
            dataGridView4.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button3.Enabled = true;
            button4.Enabled = false;
            dataGridView4.Visible = true;
            dataGridView3.Visible = false;
        }

        private void tabNorth_Click(object sender, EventArgs e)
        {

        }

        private void btn_Wall_N_S_Click(object sender, EventArgs e)
        {
            btn_Wall_N_D.Enabled = true;
            btn_Wall_N_S.Enabled = false;
            dataGrid_Wall_N_S.Visible = true;
            dataGrid_Wall_N_D.Visible = false;
        }

        private void btn_Wall_N_D_Click(object sender, EventArgs e)
        {
            btn_Wall_N_S.Enabled = true;
            btn_Wall_N_D.Enabled = false;
            dataGrid_Wall_N_D.Visible = true;
            dataGrid_Wall_N_S.Visible = false;
        }

        private void btn_Wall_S_S_Click(object sender, EventArgs e)
        {
            btn_Wall_S_D.Enabled = true;
            btn_Wall_S_S.Enabled = false;
            dataGrid_Wall_S_S.Visible = true;
            dataGrid_Wall_S_D.Visible = false;
        }

        private void btn_Wall_S_D_Click(object sender, EventArgs e)
        {
            btn_Wall_S_S.Enabled = true;
            btn_Wall_S_D.Enabled = false;
            dataGrid_Wall_S_D.Visible = true;
            dataGrid_Wall_S_S.Visible = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button7.Enabled = true;
            button8.Enabled = false;
            dataGrid_Wndw_E_S.Visible = true;
            dataGrid_Wndw_E_D.Visible = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button8.Enabled = true;
            button7.Enabled = false;
            dataGrid_Wndw_E_D.Visible = true;
            dataGrid_Wndw_E_S.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;
            button6.Enabled = false;
            dataGrid_Wall_E_S.Visible = true;
            dataGrid_Wall_E_D.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button6.Enabled = true;
            button5.Enabled = false;
            dataGrid_Wall_E_D.Visible = true;
            dataGrid_Wall_E_S.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            button11.Enabled = true;
            button12.Enabled = false;
            dataGrid_Wndw_W_S.Visible = true;
            dataGrid_Wndw_W_D.Visible = false;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            button12.Enabled = true;
            button11.Enabled = false;
            dataGrid_Wndw_W_D.Visible = true;
            dataGrid_Wndw_W_S.Visible = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            button9.Enabled = true;
            button10.Enabled = false;
            dataGrid_Wall_W_S.Visible = true;
            dataGrid_Wall_W_D.Visible = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            button10.Enabled = true;
            button9.Enabled = false;
            dataGrid_Wall_W_D.Visible = true;
            dataGrid_Wall_W_S.Visible = false;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            button15.Enabled = true;
            button16.Enabled = false;
            dataGrid_Wndw_NE_S.Visible = true;
            dataGrid_Wndw_NE_D.Visible = false;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            button16.Enabled = true;
            button15.Enabled = false;
            dataGrid_Wndw_NE_D.Visible = true;
            dataGrid_Wndw_NE_S.Visible = false;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            button13.Enabled = true;
            button14.Enabled = false;
            dataGrid_Wall_NE_S.Visible = true;
            dataGrid_Wall_NE_D.Visible = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            button14.Enabled = true;
            button13.Enabled = false;
            dataGrid_Wall_NE_D.Visible = true;
            dataGrid_Wall_NE_S.Visible = false;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            button19.Enabled = true;
            button20.Enabled = false;
            dataGrid_Wndw_NW_S.Visible = true;
            dataGrid_Wndw_NW_D.Visible = false;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            button20.Enabled = true;
            button19.Enabled = false;
            dataGrid_Wndw_NW_D.Visible = true;
            dataGrid_Wndw_NW_S.Visible = false;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            button17.Enabled = true;
            button18.Enabled = false;
            dataGrid_Wall_NW_S.Visible = true;
            dataGrid_Wall_NW_D.Visible = false;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            button18.Enabled = true;
            button17.Enabled = false;
            dataGrid_Wall_NW_D.Visible = true;
            dataGrid_Wall_NW_S.Visible = false;
        }

        private void button24_Click(object sender, EventArgs e)
        {
            button23.Enabled = true;
            button24.Enabled = false;
            dataGrid_Wndw_SE_S.Visible = true;
            dataGrid_Wndw_SE_D.Visible = false;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            button24.Enabled = true;
            button23.Enabled = false;
            dataGrid_Wndw_SE_D.Visible = true;
            dataGrid_Wndw_SE_S.Visible = false;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            button21.Enabled = true;
            button22.Enabled = false;
            dataGrid_Wall_SE_S.Visible = true;
            dataGrid_Wall_SE_D.Visible = false;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            button22.Enabled = true;
            button21.Enabled = false;
            dataGrid_Wall_SE_D.Visible = true;
            dataGrid_Wall_SE_S.Visible = false;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            button27.Enabled = true;
            button28.Enabled = false;
            dataGrid_Wndw_SW_S.Visible = true;
            dataGrid_Wndw_SW_D.Visible = false;
        }

        private void button27_Click(object sender, EventArgs e)
        {
            button28.Enabled = true;
            button27.Enabled = false;
            dataGrid_Wndw_SW_D.Visible = true;
            dataGrid_Wndw_SW_S.Visible = false;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            button25.Enabled = true;
            button26.Enabled = false;
            dataGrid_Wall_SW_S.Visible = true;
            dataGrid_Wall_SW_D.Visible = false;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            button26.Enabled = true;
            button25.Enabled = false;
            dataGrid_Wall_SW_D.Visible = true;
            dataGrid_Wall_SW_S.Visible = false;
        }

        private void button29_Click(object sender, EventArgs e)
        {      

        }

        private void label42_Click(object sender, EventArgs e)
        {

        }      


        private void Calculate_Click(object sender, EventArgs e)
        {
            double R1 = double.Parse(TB_P.Text) / double.Parse(TB_H.Text);
            TB_R1.Text = R1.ToString("0.0000");
            double Agl = double.Parse(TB_Angle.Text); 

            ////////////////North-South SC2            
            SC2_NS= Interpolation_for_SC2(Agl,R1,dataGrid_SC2_NS);
            if (SC2_NS == 0)
            {
                SC2_NS = 1;
            }
            TB_SC2_NS.Text = SC2_NS.ToString("0.0000");
            TB_SC_NS.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_NS.Text)).ToString();

            ////////////////East-West SC2            
            SC2_EW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_EW);
            if (SC2_EW == 0)
            {
                SC2_EW = 1;
            }
            TB_SC2_EW.Text = SC2_EW.ToString("0.0000");
            TB_SC_EW.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_EW.Text)).ToString();

            ////////////////NorthEast-NorthWest SC2            
            SC2_NENW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_NENW);
            if (SC2_NENW == 0)
            {
                SC2_NENW = 1;
            }
            TB_SC2_NENW.Text = SC2_NENW.ToString("0.0000");
            TB_SC_NENW.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_NENW.Text)).ToString();


            ////////////////SouthEast-SouthWest SC2            
            SC2_SESW = Interpolation_for_SC2(Agl, R1, dataGrid_SC2_SESW);
            if (SC2_SESW == 0)
            {
                SC2_SESW = 1;
            }
            TB_SC2_SESW.Text = SC2_SESW.ToString("0.0000");
            TB_SC_SESW.Text = (double.Parse(TB_SC1.Text) * double.Parse(TB_SC2_SESW.Text)).ToString();

        }

        private static double Interpolation_for_SC2(double Agl, double R1, DataGridView Data_Table)
        {
            List<double> BLst = new List<double>
            {
             0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0,
             1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.8,1.9,2.0,
             2.1,2.2,2.3,2.4,2.5,2.6,2.7,2.8,2.9,3.0
            };

            
            int a = 0;
            double y = new double();

            foreach (double b in BLst)
            {
                if (b<3.0)
                {
                    if ((Agl >= 0 && Agl < 10) && (R1 >= b && R1 < (b + 0.1)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[0].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[1].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[0].Value.ToString())))) * ((Agl - 0) / (10 - 0)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[0].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[1].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[0].Value.ToString())))) * ((Agl - 0) / (10 - 0)));
                        y = x1 + ((x2 - x1) * ((R1 - b) / ((b + 0.1) - b)));

                    }

                    else if ((Agl >= 10 && Agl < 20) && (R1 >= b && R1 < (b + 0.1)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[1].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[2].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[1].Value.ToString())))) * ((Agl - 10) / (20 - 10)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[1].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[2].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[1].Value.ToString())))) * ((Agl - 10) / (20 - 10)));
                        y = x1 + ((x2 - x1) * ((R1 - b) / ((b + 0.1) - b)));

                    }

                    else if ((Agl >= 20 && Agl < 30) && (R1 >= b && R1 < (b + 0.1)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[2].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[3].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[2].Value.ToString())))) * ((Agl - 20) / (30 - 20)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[2].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[3].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[2].Value.ToString())))) * ((Agl - 20) / (30 - 20)));
                        y = x1 + ((x2 - x1) * ((R1 - b) / ((b + 0.1) - b)));

                    }

                    else if ((Agl >= 30 && Agl < 40) && (R1 >= b && R1 < (b + 0.1)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[3].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[4].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[3].Value.ToString())))) * ((Agl - 30) / (40 - 30)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[3].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[4].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[3].Value.ToString())))) * ((Agl - 30) / (40 - 30)));
                        y = x1 + ((x2 - x1) * ((R1 - b) / ((b + 0.1) - b)));

                    }

                    else if ((Agl >= 40 && Agl <= 50) && (R1 >= b && R1 < (b + 0.1)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[4].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[5].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[4].Value.ToString())))) * ((Agl - 40) / (50 - 40)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[4].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[5].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[4].Value.ToString())))) * ((Agl - 40) / (50 - 40)));
                        y = x1 + ((x2 - x1) * ((R1 - b) / ((b + 0.1) - b)));

                    }

                    a = a + 1;

                }

                else if (b >= 3.0)
                {
                    if ((Agl >= 0 && Agl < 10) && (R1 >= b))
                    {
                        double x1 = double.Parse((Data_Table.Rows[29].Cells[0].Value.ToString())) + (((double.Parse((Data_Table.Rows[29].Cells[1].Value.ToString()))) - (double.Parse((Data_Table.Rows[29].Cells[0].Value.ToString())))) * ((Agl - 0) / (10 - 0)));                        
                        y = x1 ;
                    }

                    else if ((Agl >= 10 && Agl < 20) && (R1 >= b ))
                    {
                        double x1 = double.Parse((Data_Table.Rows[29].Cells[1].Value.ToString())) + (((double.Parse((Data_Table.Rows[29].Cells[2].Value.ToString()))) - (double.Parse((Data_Table.Rows[29].Cells[1].Value.ToString())))) * ((Agl - 10) / (20 - 10)));                        
                        y = x1;
                    }

                    else if ((Agl >= 20 && Agl < 30) && (R1 >= b ))
                    {
                        double x1 = double.Parse((Data_Table.Rows[29].Cells[2].Value.ToString())) + (((double.Parse((Data_Table.Rows[29].Cells[3].Value.ToString()))) - (double.Parse((Data_Table.Rows[29].Cells[2].Value.ToString())))) * ((Agl - 20) / (30 - 20)));                        
                        y = x1;
                    }

                    else if ((Agl >= 30 && Agl < 40) && (R1 >= b ))
                    {
                        double x1 = double.Parse((Data_Table.Rows[29].Cells[3].Value.ToString())) + (((double.Parse((Data_Table.Rows[29].Cells[4].Value.ToString()))) - (double.Parse((Data_Table.Rows[29].Cells[3].Value.ToString())))) * ((Agl - 30) / (40 - 30)));                        
                        y = x1;
                    }

                    else if ((Agl >= 40 && Agl <= 50) && (R1 >= b && R1 < (b + 0.1)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[29].Cells[4].Value.ToString())) + (((double.Parse((Data_Table.Rows[29].Cells[5].Value.ToString()))) - (double.Parse((Data_Table.Rows[29].Cells[4].Value.ToString())))) * ((Agl - 40) / (50 - 40)));                       
                        y = x1;
                    }

                }

                

            }

            

            return y;

        }

        private static double Interpolation_for_SC2_EC(double Agl, double R2, DataGridView Data_Table)
        {
            List<double> BLst = new List<double>
            {
             0.2,0.4,0.6,0.8,1.0,
             1.2,1.4,1.6,1.8
            };


            int a = 0;
            double y = new double();

            foreach (double b in BLst)
            {
                if (b < 1.8)
                {
                    if ((Agl >= 0 && Agl < 10) && (R2 >= b && R2 < (b + 0.2)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[0].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[1].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[0].Value.ToString())))) * ((Agl - 0) / (10 - 0)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[0].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[1].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[0].Value.ToString())))) * ((Agl - 0) / (10 - 0)));
                        y = x1 + ((x2 - x1) * ((R2 - b) / ((b + 0.2) - b)));

                    }

                    else if ((Agl >= 10 && Agl < 20) && (R2 >= b && R2 < (b + 0.2)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[1].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[2].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[1].Value.ToString())))) * ((Agl - 10) / (20 - 10)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[1].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[2].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[1].Value.ToString())))) * ((Agl - 10) / (20 - 10)));
                        y = x1 + ((x2 - x1) * ((R2 - b) / ((b + 0.2) - b)));

                    }

                    else if ((Agl >= 20 && Agl < 30) && (R2 >= b && R2 < (b + 0.2)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[2].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[3].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[2].Value.ToString())))) * ((Agl - 20) / (30 - 20)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[2].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[3].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[2].Value.ToString())))) * ((Agl - 20) / (30 - 20)));
                        y = x1 + ((x2 - x1) * ((R2 - b) / ((b + 0.2) - b)));

                    }

                    else if ((Agl >= 30 && Agl <= 40) && (R2 >= b && R2 < (b + 0.2)))
                    {
                        double x1 = double.Parse((Data_Table.Rows[a].Cells[3].Value.ToString())) + (((double.Parse((Data_Table.Rows[a].Cells[4].Value.ToString()))) - (double.Parse((Data_Table.Rows[a].Cells[3].Value.ToString())))) * ((Agl - 30) / (40 - 30)));
                        double x2 = double.Parse((Data_Table.Rows[a + 1].Cells[3].Value.ToString())) + (((double.Parse((Data_Table.Rows[a + 1].Cells[4].Value.ToString()))) - (double.Parse((Data_Table.Rows[a + 1].Cells[3].Value.ToString())))) * ((Agl - 30) / (40 - 30)));
                        y = x1 + ((x2 - x1) * ((R2 - b) / ((b + 0.2) - b)));

                    }                    

                    a = a + 1;

                }

                else if (b >= 1.8)
                {
                    if ((Agl >= 0 && Agl < 10) && (R2 >= b))
                    {
                        double x1 = double.Parse((Data_Table.Rows[8].Cells[0].Value.ToString())) + (((double.Parse((Data_Table.Rows[8].Cells[1].Value.ToString()))) - (double.Parse((Data_Table.Rows[8].Cells[0].Value.ToString())))) * ((Agl - 0) / (10 - 0)));
                        y = x1;
                    }

                    else if ((Agl >= 10 && Agl < 20) && (R2 >= b))
                    {
                        double x1 = double.Parse((Data_Table.Rows[8].Cells[1].Value.ToString())) + (((double.Parse((Data_Table.Rows[8].Cells[2].Value.ToString()))) - (double.Parse((Data_Table.Rows[8].Cells[1].Value.ToString())))) * ((Agl - 10) / (20 - 10)));
                        y = x1;
                    }

                    else if ((Agl >= 20 && Agl < 30) && (R2 >= b))
                    {
                        double x1 = double.Parse((Data_Table.Rows[8].Cells[2].Value.ToString())) + (((double.Parse((Data_Table.Rows[8].Cells[3].Value.ToString()))) - (double.Parse((Data_Table.Rows[8].Cells[2].Value.ToString())))) * ((Agl - 20) / (30 - 20)));
                        y = x1;
                    }

                    else if ((Agl >= 30 && Agl <= 40) && (R2 >= b))
                    {
                        double x1 = double.Parse((Data_Table.Rows[8].Cells[3].Value.ToString())) + (((double.Parse((Data_Table.Rows[8].Cells[4].Value.ToString()))) - (double.Parse((Data_Table.Rows[8].Cells[3].Value.ToString())))) * ((Agl - 30) / (40 - 30)));
                        y = x1;
                    }

                }



            }



            return y;

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void label41_Click(object sender, EventArgs e)
        {

        }

        private void button29_Click_1(object sender, EventArgs e)
        {
            double Agl_EC = double.Parse(TB_Agl_EC.Text);
            double R1_EC = double.Parse(TB_P_EC.Text) / double.Parse(TB_H_EC.Text);
            TB_R1_EC.Text = R1_EC.ToString("0.0000");
            double R2_EC = double.Parse(TB_P_EC.Text) / double.Parse(TB_W_EC.Text);
            TB_R2_EC.Text = R2_EC.ToString("0.0000");

            double b1 = 0;
            double b2 = 0;
            SC2_EC_NS_b1 = new double();
            SC2_EC_NS_b2 = new double();
            SC2_EC_EW_b1 = new double();
            SC2_EC_EW_b2 = new double();
            SC2_EC_NENW_b1 = new double();
            SC2_EC_NENW_b2 = new double();
            SC2_EC_SESW_b1 = new double();
            SC2_EC_SESW_b2 = new double();

            List<double> R1_EC_Lst = new List<double>
            {
             0.2,0.4,0.6,0.8,1.0,1.2,1.4,1.6,1.8
            };

            foreach (double b in R1_EC_Lst)
            {
                if (R1_EC<0.2)
                {
                    b1 = 0.2;
                    b2 = 0.2;                    
                    break;
                }
                if (R1_EC >= 1.8)
                {
                    b1 = 1.8;
                    b2 = 1.8;                    
                    break;
                }
                else if (R1_EC == b)
                {
                    b1 = b;
                    b2 = b;                    
                    break;
                }
                else if (R1_EC>0.2 && b > R1_EC)
                {
                    b2 = b;
                    b1 = Math.Round((b - 0.2),1);                    
                    break;
                }

            }


            ////////////////North-South SC2_EC  
            ////for b1
            if (b1==0.2)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_0);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");              
            }
            else if (b1 == 0.4)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_1);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_2);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_3);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_4);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_5);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_6);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_7);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_8);
                if (SC2_EC_NS_b1 == 0)
                {
                    SC2_EC_NS_b1 = 1;
                }
                TB_Result_NS_EC1.Text = SC2_EC_NS_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_0);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_1);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_2);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_3);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_4);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_5);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_6);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_7);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NS_8);
                if (SC2_EC_NS_b2 == 0)
                {
                    SC2_EC_NS_b2 = 1;
                }
                TB_Result_NS_EC2.Text = SC2_EC_NS_b2.ToString("0.0000");
            }

            if (SC2_EC_NS_b1== SC2_EC_NS_b2)
            {
                SC2_EC_NS = SC2_EC_NS_b1;
            }
            else
            {
                SC2_EC_NS = SC2_EC_NS_b1 + ((SC2_EC_NS_b2 - SC2_EC_NS_b1) * ((R1_EC - b1) / (b2 - b1)));                
            }
            TB_Result_NS_EC.Text = SC2_EC_NS.ToString("0.0000");

            TB_SC2_NS_EC.Text = SC2_EC_NS.ToString("0.0000"); 
            TB_SC_NS_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_NS_EC.Text)).ToString("0.0000");

            ////////////////East-West SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_0);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_1);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_2);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_3);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_4);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_5);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_6);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_7);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_8);
                if (SC2_EC_EW_b1 == 0)
                {
                    SC2_EC_EW_b1 = 1;
                }
                TB_Result_EW_EC1.Text = SC2_EC_EW_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_0);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_1);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_2);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_3);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_4);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_5);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_6);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_7);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_EW_8);
                if (SC2_EC_EW_b2 == 0)
                {
                    SC2_EC_EW_b2 = 1;
                }
                TB_Result_EW_EC2.Text = SC2_EC_EW_b2.ToString("0.0000");
            }

            if (SC2_EC_EW_b1 == SC2_EC_EW_b2)
            {
                SC2_EC_EW = SC2_EC_EW_b1;
            }
            else
            {
                SC2_EC_EW = SC2_EC_EW_b1 + ((SC2_EC_EW_b2 - SC2_EC_EW_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_EW_EC.Text = SC2_EC_EW.ToString("0.0000");

            TB_SC2_EW_EC.Text = SC2_EC_EW.ToString("0.0000");
            TB_SC_EW_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_EW_EC.Text)).ToString("0.0000");

            ////////////////NorthEast-NorthWest SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_0);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_1);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_2);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_3);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_4);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_5);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_6);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_7);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_8);
                if (SC2_EC_NENW_b1 == 0)
                {
                    SC2_EC_NENW_b1 = 1;
                }
                TB_Result_NENW_EC1.Text = SC2_EC_NENW_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_0);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_1);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_2);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_3);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_4);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_5);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_6);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_7);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_NENW_8);
                if (SC2_EC_NENW_b2 == 0)
                {
                    SC2_EC_NENW_b2 = 1;
                }
                TB_Result_NENW_EC2.Text = SC2_EC_NENW_b2.ToString("0.0000");
            }

            if (SC2_EC_NENW_b1 == SC2_EC_NENW_b2)
            {
                SC2_EC_NENW = SC2_EC_NENW_b1;
            }
            else
            {
                SC2_EC_NENW = SC2_EC_NENW_b1 + ((SC2_EC_NENW_b2 - SC2_EC_NENW_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_NENW_EC.Text = SC2_EC_NENW.ToString("0.0000");

            TB_SC2_NENW_EC.Text = SC2_EC_NENW.ToString("0.0000");
            TB_SC_NENW_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_NENW_EC.Text)).ToString("0.0000");


            ////////////////SouthEast-SouthWest SC2_EC  
            ////for b1
            if (b1 == 0.2)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_0);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 0.4)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_1);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 0.6)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_2);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 0.8)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_3);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.0)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_4);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.2)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_5);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.4)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_6);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.6)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_7);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            else if (b1 == 1.8)
            {
                SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_8);
                if (SC2_EC_SESW_b1 == 0)
                {
                    SC2_EC_SESW_b1 = 1;
                }
                TB_Result_SESW_EC1.Text = SC2_EC_SESW_b1.ToString("0.0000");
            }
            ////for b2
            if (b2 == 0.2)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_0);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 0.4)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_1);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 0.6)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_2);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 0.8)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_3);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.0)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_4);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.2)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_5);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.4)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_6);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.6)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_7);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }
            else if (b2 == 1.8)
            {
                SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl_EC, R2_EC, dataGrid_SC2_EC_SESW_8);
                if (SC2_EC_SESW_b2 == 0)
                {
                    SC2_EC_SESW_b2 = 1;
                }
                TB_Result_SESW_EC2.Text = SC2_EC_SESW_b2.ToString("0.0000");
            }

            if (SC2_EC_SESW_b1 == SC2_EC_SESW_b2)
            {
                SC2_EC_SESW = SC2_EC_SESW_b1;
            }
            else
            {
                SC2_EC_SESW = SC2_EC_SESW_b1 + ((SC2_EC_SESW_b2 - SC2_EC_SESW_b1) * ((R1_EC - b1) / (b2 - b1)));
            }
            TB_Result_SESW_EC.Text = SC2_EC_SESW.ToString("0.0000");

            TB_SC2_SESW_EC.Text = SC2_EC_SESW.ToString("0.0000");
            TB_SC_SESW_EC.Text = (double.Parse(TB_SC1_EC.Text) * double.Parse(TB_SC2_SESW_EC.Text)).ToString("0.0000");


        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button29_Click_2(object sender, EventArgs e)
        {
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////////////////////////// Window SC Calculations_Vertical /////////////////////////////////////

            double R1_VP = double.Parse(TB_P_VP.Text) / double.Parse(TB_H_VP.Text);
            TB_R1_VP.Text = R1_VP.ToString("0.0000");
            double Agl_VP = double.Parse(TB_Angle_VP.Text);

            ////////////////////////// For North-South SC2 /////////////////////////////////
            SC2_VP_NS = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_NS);
            if (SC2_VP_NS == 0)
            {
                SC2_VP_NS = 1;
            }
            TB_SC2_VP_NS.Text = SC2_VP_NS.ToString("0.0000");
            TB_SC_VP_NS.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_NS.Text)).ToString();

            ////////////////////////// For East-West SC2 /////////////////////////////////
            SC2_VP_EW = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_EW);
            if (SC2_VP_EW == 0)
            {
                SC2_VP_EW = 1;
            }
            TB_SC2_VP_EW.Text = SC2_VP_EW.ToString("0.0000");
            TB_SC_VP_EW.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_EW.Text)).ToString();

            ////////////////////////// For NorthEast-NorthWest SC2 /////////////////////////////////
            SC2_VP_NENW = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_NENW);
            if (SC2_VP_NENW == 0)
            {
                SC2_VP_NENW = 1;
            }
            TB_SC2_VP_NENW.Text = SC2_VP_NENW.ToString("0.0000");
            TB_SC_VP_NENW.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_NENW.Text)).ToString();

            ////////////////////////// For SouthEast-SouthWest SC2 /////////////////////////////////
            SC2_VP_SESW = Interpolation_for_SC2(Agl_VP, R1_VP, dataGrid_SC2_VP_SESW);
            if (SC2_VP_SESW == 0)
            {
                SC2_VP_SESW = 1;
            }
            TB_SC2_VP_SESW.Text = SC2_VP_SESW.ToString("0.0000");
            TB_SC_VP_SESW.Text = (double.Parse(TB_SC1_VP.Text) * double.Parse(TB_SC2_VP_SESW.Text)).ToString();


            //////////////////////////// Window SC Calculations_Vertical /////////////////////////////////////

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        }

        private void dataGrid_SC2_EW_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //// Generate An Excel
       
        private void button30_Click(object sender, EventArgs e)
        {

            //Excel.Application excelApp = null;
            //Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Create a new instance of Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.DisplayAlerts = false;

                // Create a new workbook
                workbook = excelApp.Workbooks.Add();              


                ///////////////////////////////////////////////////////////// for Cover Page worksheet - Starts ////////////////////////////////////////////////////////////////
                // Get the first worksheet in the workbook
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                // Set the name of the worksheet to "Cover Page"
                worksheet.Name = "Cover Page";
                // Set the tab color to light red
                worksheet.Tab.Color = System.Drawing.Color.LightSalmon;

                // Merge cells A1 to J3 and set alignment
                Excel.Range mergeRange0 = worksheet.Range["H1", "J2"];
                mergeRange0.Merge();
                mergeRange0.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                mergeRange0.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                // Merge cells A4 to J4 and set alignment
                Excel.Range mergeRange = worksheet.Range["A4", "J4"];
                mergeRange.Merge();
                mergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Merge cells A5 to J5 and set alignment
                Excel.Range mergeRange1 = worksheet.Range["A5", "J5"];
                mergeRange1.Merge();
                mergeRange1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Insert BCA Logo image from directory
                string imagePath = BuildingCoder.Util.GetFilePath("BCA_Logo.png");  //"D:\WIP\01_API\ETTV\BCA_Logo.png"; // Specify the path to your PNG image
                if (File.Exists(imagePath))
                {
                    // Calculate the position and size of the image to fit within the merge cell
                    float left = (float)mergeRange0.Left;
                    float top = (float)mergeRange0.Top;
                    float width = (float)mergeRange0.Width;
                    float height = (float)mergeRange0.Height;
                    // Add the image to the worksheet
                    worksheet.Shapes.AddPicture(imagePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
                }
                else
                {
                    MessageBox.Show("Image file not found at the specified path.");
                }



                // Write Heading text to the merged cells                
                mergeRange.Value = "ETTV CALCULATION FORMAT IN RESPECT OF";
                mergeRange1.Value = "AN AIRCONDITIONED BUILDING";

                // Apply font stype & bold formatting to the text
                mergeRange.Font.Bold = true;
                mergeRange.Font.Name = "Times New Roman";
                mergeRange1.Font.Bold = true;
                mergeRange1.Font.Name = "Times New Roman";

                // Add outside borders to the merged range
                mergeRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                mergeRange1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange1.Borders.Weight = Excel.XlBorderWeight.xlThin;



                // Merge cells A7 to D7 and set alignment
                Excel.Range mergeRange2 = worksheet.Range["A7", "D7"];
                mergeRange2.Merge();
                mergeRange2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply top & right border to the merged cells
                mergeRange2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange2.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange2.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange2.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells A14 to D14 and set alignment
                Excel.Range mergeRange3 = worksheet.Range["A14", "D14"];
                mergeRange3.Merge();
                mergeRange3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply top & right border to the merged cells
                mergeRange3.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange3.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange3.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;



                // Merge cells and apply borders for ranges A8:D8 to A13:D13
                MergeCellsAndApplyBorders(worksheet.Range["A8", "D8"], "Commissioner of Building Control");
                MergeCellsAndApplyBorders(worksheet.Range["A9", "D9"], "Building & Construction Authority");
                MergeCellsAndApplyBorders(worksheet.Range["A10", "D10"], "52 Jurong Gateway Road");
                MergeCellsAndApplyBorders(worksheet.Range["A11", "D11"], "#11-01");
                MergeCellsAndApplyBorders(worksheet.Range["A12", "D12"], "Singapore 608550");
                MergeCellsAndApplyBorders(worksheet.Range["A13", "D13"], "");
                                              

                // Merge cells E7 to J7 and set alignment
                Excel.Range mergeRange4 = worksheet.Range["E7", "J7"];
                mergeRange4.Merge();
                mergeRange4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply top & right border to the merged cells
                mergeRange4.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange4.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange4.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange4.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells E14 to J14 and set alignment
                Excel.Range mergeRange5 = worksheet.Range["E14", "J14"];
                mergeRange5.Merge();
                mergeRange5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply top & right border to the merged cells
                mergeRange5.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange5.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange5.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange5.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells and apply borders for ranges E8:J8 to E13:J13
                MergeCellsAndApplyBorders(worksheet.Range["E8", "J8"], "INSTRUCTIONS:");
                MergeCellsAndApplyBorders(worksheet.Range["E9", "J9"], "(1) Please refer to the Explanatory Notes attached before completing the form.");
                MergeCellsAndApplyBorders(worksheet.Range["E10", "J10"], "(2) Use a separate set of forms for each building.");
                MergeCellsAndApplyBorders(worksheet.Range["E11", "J11"], "(3) *Delete accordingly.");
                MergeCellsAndApplyBorders(worksheet.Range["E12", "J12"], "(4) One copy of this form together with Appendix 1 to 4");
                MergeCellsAndApplyBorders(worksheet.Range["E13", "J13"], "to be submitted for each façade.");
                ///////------------------------------------------------------------------------------------------//////

                // Merge cells and apply borders for ranges
                MergeCellsAndApplyBorders(worksheet.Range["A15", "J15"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A16", "J16"], "(1) *I/We, the Qualified Person(s) responsible for the preparation of the ETTV calculation and the building");
                MergeCellsAndApplyBorders(worksheet.Range["A17", "J17"], "plans hereby submit, for your consideration, the ETTV calculation and detail plans for:-");
                MergeCellsAndApplyBorders(worksheet.Range["A18", "J18"], "");

                // Merge cells A18 to C18 and set alignment
                Excel.Range mergeRange6 = worksheet.Range["A19", "C19"];
                mergeRange6.Merge();
                mergeRange6.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange6.Value = "Project Ref. No. :";
                mergeRange6.Font.Name = "Times New Roman";

                // Merge cells D19 to J19 and set alignment
                Excel.Range mergeRange7 = worksheet.Range["D19", "J19"];
                mergeRange7.Merge();
                mergeRange7.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange7.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange7.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange7.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange7.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange7.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A20", "J20"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A21", "J21"], "Description of Building/Building Works:");

                // Merge cells A22 to J24 and set alignment
                Excel.Range mergeRange8 = worksheet.Range["A22", "J24"];
                mergeRange8.Merge();
                mergeRange8.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange8.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange8.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange8.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange8.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange8.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells A25 to B25 and set alignment
                Excel.Range mergeRange9 = worksheet.Range["A25", "B25"];
                mergeRange9.Merge();
                mergeRange9.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange9.Value = "*Lot/Plot :";
                mergeRange9.Font.Name = "Times New Roman";
                // Merge cells C25 to E25 and set alignment
                Excel.Range mergeRange10 = worksheet.Range["C25", "E25"];
                mergeRange10.Merge();
                mergeRange10.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange10.Value = "";
                mergeRange10.Font.Name = "Times New Roman";

                // Merge cells F25 to G25 and set alignment
                Excel.Range mergeRange11 = worksheet.Range["F25", "G25"];
                mergeRange11.Merge();
                mergeRange11.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange11.Value = "*TS/MK :";
                mergeRange11.Font.Name = "Times New Roman";
                // Merge cells H25 to J25 and set alignment
                Excel.Range mergeRange12 = worksheet.Range["H25", "J25"];
                mergeRange12.Merge();
                mergeRange12.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange12.Value = "";
                mergeRange12.Font.Name = "Times New Roman";
                mergeRange12.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange12.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells A26 to C26 and set alignment
                Excel.Range mergeRange13 = worksheet.Range["A26", "C26"];
                mergeRange13.Merge();
                mergeRange13.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange13.Value = "Address/Road :";
                mergeRange13.Font.Name = "Times New Roman";

                // Merge cells D26 to J26 and set alignment
                Excel.Range mergeRange14 = worksheet.Range["D26", "J26"];
                mergeRange14.Merge();
                mergeRange14.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange14.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange14.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange14.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange14.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange14.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A27", "J27"], "");

                // Merge cells A28 to J28 and set alignment
                Excel.Range mergeRange15 = worksheet.Range["A28", "J28"];
                mergeRange15.Merge();
                mergeRange15.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange15.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange15.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange15.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange15.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A29", "J29"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A30", "J30"], "(2) The following are attached:-");

                


                // Merge cells A37 to J37 and set alignment
                Excel.Range mergeRange16 = worksheet.Range["A37", "J37"];
                mergeRange16.Merge();
                mergeRange16.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange16.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange16.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange16.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange16.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                ///////------------------------------------------------------------------------------------------//////

                MergeCellsAndApplyBorders(worksheet.Range["A38", "J38"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A39", "J39"], "(3) The ETTVs of the various Façades are as follows:-");
                MergeCellsAndApplyBorders(worksheet.Range["A40", "J40"], "");

                WriteETTVValuesToExcel("(a)", "Façade 1","N", Lb_ETTVRes_N.Text, 41);
                WriteETTVValuesToExcel("(b)", "Façade 2","S", Lb_ETTVRes_S.Text, 42);
                WriteETTVValuesToExcel("(c)", "Façade 3","E", Lb_ETTVRes_E.Text, 43);
                WriteETTVValuesToExcel("(d)", "Façade 4","W", Lb_ETTVRes_W.Text, 44);
                WriteETTVValuesToExcel("(e)", "Façade 5", "NE", Lb_ETTVRes_NE.Text, 45);
                WriteETTVValuesToExcel("(f)", "Façade 6", "NW", Lb_ETTVRes_NW.Text, 46);
                WriteETTVValuesToExcel("(g)", "Façade 7", "SE", Lb_ETTVRes_SE.Text, 47);
                WriteETTVValuesToExcel("(h)", "Façade 8", "SW", Lb_ETTVRes_SW.Text, 48);

                // Merge cells A49 to J49 and set alignment
                Excel.Range mergeRange17 = worksheet.Range["A49", "J49"];
                mergeRange17.Merge();
                mergeRange17.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange17.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange17.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange17.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange17.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A50", "J50"], "");

                Excel.Range mergeRange18 = worksheet.Range["A51", "E51"];
                mergeRange18.Merge();
                mergeRange18.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange18.Value = "The average ETTV of the building envelope is";
                mergeRange18.Font.Name = "Times New Roman";

                worksheet.Cells[51, 6].Font.Name = "Times New Roman";
                worksheet.Cells[51, 6].Value = Lb_ETTV_Avg.Text;
                worksheet.Cells[51, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[51, 6].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                worksheet.Cells[51, 7].Font.Name = "Times New Roman";
                worksheet.Cells[51, 7].Value = "W/m2";
                worksheet.Cells[51, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                MergeCellsAndApplyBorders(worksheet.Range["H51", "J51"], "");

                // Merge cells A52 to J52 and set alignment
                Excel.Range mergeRange19 = worksheet.Range["A52", "J52"];
                mergeRange19.Merge();
                mergeRange19.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange19.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange19.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange19.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange19.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                void WriteETTVValuesToExcel(string label, string facade, string Ort, string ettvValue, int row)
                {
                    worksheet.Cells[row, 1].Font.Name = "Times New Roman";
                    worksheet.Cells[row, 1].Value = label;
                    worksheet.Cells[row, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    worksheet.Cells[row, 2].Font.Name = "Times New Roman";
                    worksheet.Cells[row, 2].Value = facade + " facing:";
                    worksheet.Cells[row, 4].Font.Name = "Times New Roman";
                    worksheet.Cells[row, 4].Value = Ort;
                    worksheet.Cells[row, 4].Font.Bold = true;
                    worksheet.Cells[row, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[row, 5].Font.Name = "Times New Roman";
                    worksheet.Cells[row, 5].Value = "ETTV=";
                    worksheet.Cells[row, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    worksheet.Cells[row, 6].Font.Name = "Times New Roman";
                    worksheet.Cells[row, 6].Value = ettvValue;
                    worksheet.Cells[row, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[row, 6].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    Excel.Range rangeVV = worksheet.Cells[row, 6];
                    rangeVV.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                    rangeVV.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                    worksheet.Cells[row, 7].Font.Name = "Times New Roman";
                    worksheet.Cells[row, 7].Value = "W/m2";
                    worksheet.Cells[row, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    MergeCellsAndApplyBorders(worksheet.Range["H" + row, "J" + row], "");
                }

                ///////------------------------------------------------------------------------------------------//////

                MergeCellsAndApplyBorders(worksheet.Range["A53", "J53"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A54", "J54"], "(4) I hereby certify that the ETTV calculation as detailed in this submission has been done in accordance with");
                MergeCellsAndApplyBorders(worksheet.Range["A55", "J55"], "the Guidelines On Envelope Thermal Transfer Value For Buildings and that, to the best of my knowledge,");
                MergeCellsAndApplyBorders(worksheet.Range["A56", "J56"], "the computed ETTV is correct.");

                // Merge cells A57 to J57 and set alignment
                Excel.Range mergeRange20 = worksheet.Range["A57", "J57"];
                mergeRange20.Merge();
                mergeRange20.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange20.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange20.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange20.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange20.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A58", "E58"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F58", "J58"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A59", "E59"], "Name & Address of Professional Firm");
                MergeCellsAndApplyBorders(worksheet.Range["F59", "J59"], "Name & Signature of Qualified Person who");
                MergeCellsAndApplyBorders(worksheet.Range["A60", "E60"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F60", "J60"], "prepared the calculation");

                // Merge cells A61 to E64 and set alignment
                Excel.Range mergeRange21 = worksheet.Range["A61", "E64"];
                mergeRange21.Merge();
                mergeRange21.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange21.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange21.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange21.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange21.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange21.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells F61 to J64 and set alignment
                Excel.Range mergeRange22 = worksheet.Range["F61", "J64"];
                mergeRange22.Merge();
                mergeRange22.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange22.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange22.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange22.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange22.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange22.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A65", "E65"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F65", "J65"], "");

                // Merge cells A66 to E66 and set alignment
                Excel.Range mergeRange23 = worksheet.Range["A66", "E66"];
                mergeRange23.Merge();
                mergeRange23.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange23.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange23.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange23.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange23.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells F66 to J66 and set alignment
                Excel.Range mergeRange24 = worksheet.Range["F66", "J66"];
                mergeRange24.Merge();
                mergeRange24.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange24.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange24.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange24.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange24.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A67", "E67"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F67", "J67"], "");

                worksheet.Cells[68, 1].Font.Name = "Times New Roman";
                worksheet.Cells[68, 1].Value = "Date :";
                worksheet.Cells[68, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                // Merge cells B68 to E68 and set alignment
                Excel.Range mergeRange25 = worksheet.Range["B68", "E68"];
                mergeRange25.Merge();
                mergeRange25.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange25.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange25.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange25.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange25.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange25.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                worksheet.Cells[68, 6].Font.Name = "Times New Roman";
                worksheet.Cells[68, 6].Value = "Tel. No. :";
                worksheet.Cells[68, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                // Merge cells F68 to J68 and set alignment
                Excel.Range mergeRange26 = worksheet.Range["G68", "J68"];
                mergeRange26.Merge();
                mergeRange26.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange26.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange26.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange26.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange26.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange26.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                // Merge cells A69 to E69 and set alignment
                Excel.Range mergeRange27 = worksheet.Range["A69", "E69"];
                mergeRange27.Merge();
                mergeRange27.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange27.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange27.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange27.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange27.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;


                // Merge cells F69 to J69 and set alignment
                Excel.Range mergeRange28 = worksheet.Range["F69", "J69"];
                mergeRange28.Merge();
                mergeRange28.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange28.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange28.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange28.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange28.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                ///////------------------------------------------------------------------------------------------//////

                MergeCellsAndApplyBorders(worksheet.Range["A70", "J70"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A71", "J71"], "(5) I hereby certify that the materials, methods of external shading and areas of the building envelope");
                MergeCellsAndApplyBorders(worksheet.Range["A72", "J72"], "as detailed in this submission are correct, and that the building works *shall be/have been  constructed");
                MergeCellsAndApplyBorders(worksheet.Range["A73", "J73"], "in accordance with the envelope specifications outlined herein.");

                // Merge cells A71 to J71 and set alignment
                Excel.Range mergeRange29 = worksheet.Range["A74", "J74"];
                mergeRange29.Merge();
                mergeRange29.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange29.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange29.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange29.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange29.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A75", "E75"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F75", "J75"], "");
                MergeCellsAndApplyBorders(worksheet.Range["A76", "E76"], "Name & Address of Professional Firm");
                MergeCellsAndApplyBorders(worksheet.Range["F76", "J76"], "Name & Signature of Qualified Person who");
                MergeCellsAndApplyBorders(worksheet.Range["A77", "E77"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F77", "J77"], "signed building plans");

                // Merge cells A78 to E81 and set alignment
                Excel.Range mergeRange30 = worksheet.Range["A78", "E81"];
                mergeRange30.Merge();
                mergeRange30.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange30.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange30.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange30.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange30.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange30.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells F78 to J81 and set alignment
                Excel.Range mergeRange31 = worksheet.Range["F78", "J81"];
                mergeRange31.Merge();
                mergeRange31.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange31.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange31.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange31.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange31.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange31.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A82", "E82"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F82", "J82"], "");

                // Merge cells A83 to E83 and set alignment
                Excel.Range mergeRange32 = worksheet.Range["A83", "E83"];
                mergeRange32.Merge();
                mergeRange32.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange32.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange32.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange32.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange32.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells F83 to J83 and set alignment
                Excel.Range mergeRange33 = worksheet.Range["F83", "J83"];
                mergeRange33.Merge();
                mergeRange33.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange33.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange33.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange33.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange33.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                MergeCellsAndApplyBorders(worksheet.Range["A84", "E84"], "");
                MergeCellsAndApplyBorders(worksheet.Range["F84", "J84"], "");

                worksheet.Cells[85, 1].Font.Name = "Times New Roman";
                worksheet.Cells[85, 1].Value = "Date :";
                worksheet.Cells[85, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                // Merge cells B85 to E85 and set alignment
                Excel.Range mergeRange34 = worksheet.Range["B85", "E85"];
                mergeRange34.Merge();
                mergeRange34.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange34.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange34.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange34.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange34.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange34.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                worksheet.Cells[85, 6].Font.Name = "Times New Roman";
                worksheet.Cells[85, 6].Value = "Tel. No. :";
                worksheet.Cells[85, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                // Merge cells E84 to J84 and set alignment
                Excel.Range mergeRange35 = worksheet.Range["G85", "J85"];
                mergeRange35.Merge();
                mergeRange35.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                mergeRange35.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange35.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                // Fill the merged cell with yellow color
                mergeRange35.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                mergeRange35.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                mergeRange35.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                // Merge cells A86 to E86 and set alignment
                Excel.Range mergeRange36 = worksheet.Range["A86", "E86"];
                mergeRange36.Merge();
                mergeRange36.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange36.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange36.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange36.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange36.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;


                // Merge cells F86 to J86 and set alignment
                Excel.Range mergeRange37 = worksheet.Range["F86", "J86"];
                mergeRange37.Merge();
                mergeRange37.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                // Apply bottom & right border to the merged cells
                mergeRange37.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange37.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                mergeRange37.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                mergeRange37.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Set the column width for Brief Description
                worksheet.Columns[10].ColumnWidth = 20;


                ///////------------------------------------------------------------------------------------------//////

                ///////////////////////////////////////////////////////////// for Cover Page worksheet - Ends ////////////////////////////////////////////////////////////////

                ///////////////////////////////////////////////////////////// for AP1 Worksheets - Starts ////////////////////////////////////////////////////////////////

                AddWorksheetsForAP1Tabs();
                int rowAdder1 = 6;
                //int newSepNum1 = dataGrid_Wall_N_S.Rows.Count + rowAdder1 + 1;
                //int newSepNum2 = newSepNum1 + 1;
                //int newSepNum3 = newSepNum2 + dataGridView2.Rows.Count + 3;

                //ActivateWorksheet(workbook, "AP1_North");
                // Get the active worksheet
                //Microsoft.Office.Interop.Excel.Worksheet activeSheet = excelApp.ActiveSheet;

                ///// For AP1_North Starts/////
                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_North = workbook.Sheets["AP1_North"];

                Excel.Range mergeRange_N2 = worksheet_AP1_North.Range["A2", "J2"];
                mergeRange_N2.Merge();
                mergeRange_N2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_N2.Font.Name = "Times New Roman";
                mergeRange_N2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_N2.Font.Bold = true;

                Excel.Range mergeRange_N3 = worksheet_AP1_North.Range["A3", "J3"];
                mergeRange_N3.Merge();
                mergeRange_N3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_N3.Font.Name = "Times New Roman";
                mergeRange_N3.Value = "FAÇADE ORIENTATION : N";
                mergeRange_N3.Font.Bold = true;

                //////////// for Opaque Walls
                
                Excel.Range mergeRange_N5 = worksheet_AP1_North.Range["A5", "J5"];
                mergeRange_N5.Merge();
                mergeRange_N5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_N5.Font.Name = "Times New Roman";
                mergeRange_N5.Value = "OPAQUE WALLS";
                mergeRange_N5.Font.Bold = true;                

                Excel.Range mergeRange_N_6_1 = worksheet_AP1_North.Cells[6, 1];                
                mergeRange_N_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_6_1.Value = "S/No";
                mergeRange_N_6_1.Font.Name = "Times New Roman";
                mergeRange_N_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_6_1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_6_2_5 = worksheet_AP1_North.Range["B6", "E6"];
                mergeRange_N_6_2_5.Merge();
                mergeRange_N_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_6_2_5.Value = "Brief Description";
                mergeRange_N_6_2_5.Font.Name = "Times New Roman";
                mergeRange_N_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_6_2_5.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_6_6 = worksheet_AP1_North.Cells[6, 6];
                mergeRange_N_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_6_6.Value = "Aw";
                mergeRange_N_6_6.Font.Name = "Times New Roman";
                mergeRange_N_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_6_6.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_6_7 = worksheet_AP1_North.Cells[6, 7];
                mergeRange_N_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_6_7.Value = "Uw";
                mergeRange_N_6_7.Font.Name = "Times New Roman";
                mergeRange_N_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_6_7.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_6_8_9 = worksheet_AP1_North.Range["H6", "J6"];
                mergeRange_N_6_8_9.Merge();
                mergeRange_N_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_6_8_9.Value = "12*Aw*Uw";
                mergeRange_N_6_8_9.Font.Name = "Times New Roman";
                mergeRange_N_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_6_8_9.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_N_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_N_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_N_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;

                            

                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_N = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_North.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_N.Merge();
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_N_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_N_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_N_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_N = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_North.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_N.Merge();
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_N_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //light blue ==> Wall Types (1st Column) 
                // Microsoft.Office.Interop.Excel.Range colorBackgroundFor_S_NO2 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[newSepNum2 + 1 - 7, 1], worksheet_AP1_North.Cells[newSepNum3 - 5 - 7, 1]];
                //colorBackgroundFor_S_NO2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                //colorBackgroundFor_S_NO2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                //colorBackgroundFor_S_NO2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                //colorBackgroundFor_S_NO2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;


                //light yellow ==> Wall Description, Area & U Values
                //Microsoft.Office.Interop.Excel.Range colorBackgroundForOpaqueWalls1 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[newSepNum2 + 1-7, 2], worksheet_AP1_North.Cells[newSepNum3 - 5-7, 7]];
                //colorBackgroundForOpaqueWalls1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 204));
                //colorBackgroundForOpaqueWalls1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                //colorBackgroundForOpaqueWalls1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                //colorBackgroundForOpaqueWalls1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;

                //light orange pink ==> Wall 12*Aw*Uw
                //Microsoft.Office.Interop.Excel.Range colorBackgroundForOpaqueWalls2 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[newSepNum2 + 1-7, 8], worksheet_AP1_North.Cells[newSepNum3 - 5-7, 10]];
                //colorBackgroundForOpaqueWalls2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(253, 233, 217));
                //colorBackgroundForOpaqueWalls2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                //colorBackgroundForOpaqueWalls2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                //colorBackgroundForOpaqueWalls2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;

                //light pink ==> Wall Subtotal
                //Microsoft.Office.Interop.Excel.Range colorBackgroundForSubTotal2 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[newSepNum3 - 4-7, 1], worksheet_AP1_North.Cells[newSepNum3 - 4-7, 10]];
                //colorBackgroundForSubTotal2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 255));
                //colorBackgroundForSubTotal2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                //colorBackgroundForSubTotal2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                //colorBackgroundForSubTotal2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;


                //////////// for Fenestration

                int FenSep = dataGrid_Wall_N_S.Rows.Count + rowAdder1 + 1;;

                Excel.Range mergeRange_N_Fen1 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep, 1], worksheet_AP1_North.Cells[FenSep, 10]];
                mergeRange_N_Fen1.Merge();
                mergeRange_N_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_N_Fen1.Font.Name = "Times New Roman";
                mergeRange_N_Fen1.Value = "FENESTRATION";
                mergeRange_N_Fen1.Font.Bold = true;

                Excel.Range mergeRange_N_Fen2 = worksheet_AP1_North.Cells[(FenSep + 1), 1];
                mergeRange_N_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen2.Value = "S/No";
                mergeRange_N_Fen2.Font.Name = "Times New Roman";
                mergeRange_N_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_Fen3 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep + 1, 2], worksheet_AP1_North.Cells[FenSep + 1, 3]];
                mergeRange_N_Fen3.Merge();
                mergeRange_N_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen3.Value = "Brief Description";
                mergeRange_N_Fen3.Font.Name = "Times New Roman";
                mergeRange_N_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_Fen4 = worksheet_AP1_North.Cells[(FenSep + 1), 4];
                mergeRange_N_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen4.Value = "Af";
                mergeRange_N_Fen4.Font.Name = "Times New Roman";
                mergeRange_N_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen4.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_Fen5 = worksheet_AP1_North.Cells[(FenSep + 1), 5];
                mergeRange_N_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen5.Value = "Uf";
                mergeRange_N_Fen5.Font.Name = "Times New Roman";
                mergeRange_N_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen5.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));


                Excel.Range mergeRange_N_Fen6 = worksheet_AP1_North.Cells[(FenSep + 1), 6];
                mergeRange_N_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen6.Value = "SC";
                mergeRange_N_Fen6.Font.Name = "Times New Roman";
                mergeRange_N_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen6.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_Fen7 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep + 1, 7], worksheet_AP1_North.Cells[FenSep + 1, 8]];
                mergeRange_N_Fen7.Merge();
                mergeRange_N_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen7.Value = "3.4*Af*Uf";
                mergeRange_N_Fen7.Font.Name = "Times New Roman";
                mergeRange_N_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen7.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                Excel.Range mergeRange_N_Fen8 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep + 1, 9], worksheet_AP1_North.Cells[FenSep + 1, 10]];
                mergeRange_N_Fen8.Merge();
                mergeRange_N_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_N_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_N_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_N_Fen8.Value = "211*Af*SC*CF";
                mergeRange_N_Fen8.Font.Name = "Times New Roman";
                mergeRange_N_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ///for color==> light green
                //mergeRange_N_Fen8.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

                // Set the column width for Brief Description
                worksheet_AP1_North.Columns[3].ColumnWidth = 30;

                int FenSep1 = FenSep + 2;

                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep1 + i, 2], worksheet_AP1_North.Cells[FenSep1 + i, 3]];
                            mergeRangeFenValue1.Merge();
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2] = dataGridView2.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2] = dataGridView2.Rows[i].Cells[j].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2] = dataGridView2.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep1 + i, 7], worksheet_AP1_North.Cells[FenSep1 + i, 8]];
                            mergeRangeFenValue2.Merge();
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2] = dataGridView2.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep1 + i, 9], worksheet_AP1_North.Cells[FenSep1 + i, 10]];
                            mergeRangeFenValue3.Merge();
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 3] = dataGridView2.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_North.Cells[FenSep1 + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2 = FenSep1 + dataGridView2.Rows.Count + 1;

                Excel.Range mergeRange_N_ETTV1 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep2, 2], worksheet_AP1_North.Cells[FenSep2, 3]];
                mergeRange_N_ETTV1.Merge();
                mergeRange_N_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_N_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_N_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_N_ETTV1.Font.Bold = true;
                worksheet_AP1_North.Cells[FenSep2, 4] = Lb_Area_N.Text;
                worksheet_AP1_North.Cells[FenSep2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_North.Cells[FenSep2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_North.Cells[FenSep2, 4].Font.Bold = true;
                worksheet_AP1_North.Cells[FenSep2, 5] = "m2";
                worksheet_AP1_North.Cells[FenSep2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_North.Cells[FenSep2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_North.Cells[FenSep2, 5].Font.Bold = true;


                Excel.Range mergeRange_N_ETTV2 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep2 + 2, 2], worksheet_AP1_North.Cells[FenSep2 + 2, 3]];
                mergeRange_N_ETTV2.Merge();
                mergeRange_N_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_N_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_N_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_N_ETTV2.Font.Bold = true;
                worksheet_AP1_North.Cells[FenSep2+2, 4] = Lb_HG_N.Text;
                worksheet_AP1_North.Cells[FenSep2+2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_North.Cells[FenSep2+2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_North.Cells[FenSep2 + 2, 4].Font.Bold = true;
                worksheet_AP1_North.Cells[FenSep2+2, 5] = "W";
                worksheet_AP1_North.Cells[FenSep2+2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_North.Cells[FenSep2+2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_North.Cells[FenSep2+2, 5].Font.Bold = true;


                Excel.Range mergeRange_N_ETTV3 = worksheet_AP1_North.Range[worksheet_AP1_North.Cells[FenSep2 + 4, 2], worksheet_AP1_North.Cells[FenSep2 + 4, 3]];
                mergeRange_N_ETTV3.Merge();
                mergeRange_N_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_N_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_N_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_N_ETTV3.Font.Bold = true;
                worksheet_AP1_North.Cells[FenSep2 + 4, 4] = Lb_ETTV_N.Text;
                worksheet_AP1_North.Cells[FenSep2 + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_North.Cells[FenSep2 + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_North.Cells[FenSep2 + 4, 4].Font.Bold = true;
                worksheet_AP1_North.Cells[FenSep2 + 4, 5] = "W / m2";
                worksheet_AP1_North.Cells[FenSep2 + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_North.Cells[FenSep2 + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_North.Cells[FenSep2 + 4, 5].Font.Bold = true;


                ///// For AP1_North Ends/////
                ///


                ///// For AP1_South Starts/////

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_South = workbook.Sheets["AP1_South"];

                Excel.Range mergeRange_S2 = worksheet_AP1_South.Range["A2", "J2"];
                mergeRange_S2.Merge();
                mergeRange_S2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_S2.Font.Name = "Times New Roman";
                mergeRange_S2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_S2.Font.Bold = true;

                Excel.Range mergeRange_S3 = worksheet_AP1_South.Range["A3", "J3"];
                mergeRange_S3.Merge();
                mergeRange_S3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_S3.Font.Name = "Times New Roman";
                mergeRange_S3.Value = "FAÇADE ORIENTATION : S";
                mergeRange_S3.Font.Bold = true;

                //////////// for Opaque Walls
                
                Excel.Range mergeRange_S5 = worksheet_AP1_South.Range["A5", "J5"];
                mergeRange_S5.Merge();
                mergeRange_S5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_S5.Font.Name = "Times New Roman";
                mergeRange_S5.Value = "OPAQUE WALLS";
                mergeRange_S5.Font.Bold = true;

                Excel.Range mergeRange_S_6_1 = worksheet_AP1_South.Cells[6, 1];
                mergeRange_S_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_6_1.Value = "S/No";
                mergeRange_S_6_1.Font.Name = "Times New Roman";
                mergeRange_S_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_S_6_2_5 = worksheet_AP1_South.Range["B6", "E6"];
                mergeRange_S_6_2_5.Merge();
                mergeRange_S_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_6_2_5.Value = "Brief Description";
                mergeRange_S_6_2_5.Font.Name = "Times New Roman";
                mergeRange_S_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_S_6_6 = worksheet_AP1_South.Cells[6, 6];
                mergeRange_S_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_6_6.Value = "Aw";
                mergeRange_S_6_6.Font.Name = "Times New Roman";
                mergeRange_S_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_S_6_7 = worksheet_AP1_South.Cells[6, 7];
                mergeRange_S_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_6_7.Value = "Uw";
                mergeRange_S_6_7.Font.Name = "Times New Roman";
                mergeRange_S_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_S_6_8_9 = worksheet_AP1_South.Range["H6", "J6"];
                mergeRange_S_6_8_9.Merge();
                mergeRange_S_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_6_8_9.Value = "12*Aw*Uw";
                mergeRange_S_6_8_9.Font.Name = "Times New Roman";
                mergeRange_S_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_S_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_S_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_S_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;



                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_S = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_South.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_S.Merge();
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_S_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_S_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_S_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_S = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_South.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_S.Merge();
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_S_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_S = dataGrid_Wall_S_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_S_Fen1 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep_S, 1], worksheet_AP1_South.Cells[FenSep_S, 10]];
                mergeRange_S_Fen1.Merge();
                mergeRange_S_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_S_Fen1.Font.Name = "Times New Roman";
                mergeRange_S_Fen1.Value = "FENESTRATION";
                mergeRange_S_Fen1.Font.Bold = true;


                Excel.Range mergeRange_S_Fen2 = worksheet_AP1_South.Cells[(FenSep_S + 1), 1];
                mergeRange_S_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen2.Value = "S/No";
                mergeRange_S_Fen2.Font.Name = "Times New Roman";
                mergeRange_S_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                

                Excel.Range mergeRange_S_Fen3 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep_S + 1, 2], worksheet_AP1_South.Cells[FenSep_S + 1, 3]];
                mergeRange_S_Fen3.Merge();
                mergeRange_S_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen3.Value = "Brief Description";
                mergeRange_S_Fen3.Font.Name = "Times New Roman";
                mergeRange_S_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                

                Excel.Range mergeRange_S_Fen4 = worksheet_AP1_South.Cells[(FenSep_S + 1), 4];
                mergeRange_S_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen4.Value = "Af";
                mergeRange_S_Fen4.Font.Name = "Times New Roman";
                mergeRange_S_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_S_Fen5 = worksheet_AP1_South.Cells[(FenSep_S + 1), 5];
                mergeRange_S_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen5.Value = "Uf";
                mergeRange_S_Fen5.Font.Name = "Times New Roman";
                mergeRange_S_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                

                Excel.Range mergeRange_S_Fen6 = worksheet_AP1_South.Cells[(FenSep_S + 1), 6];
                mergeRange_S_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen6.Value = "SC";
                mergeRange_S_Fen6.Font.Name = "Times New Roman";
                mergeRange_S_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                

                Excel.Range mergeRange_S_Fen7 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep_S + 1, 7], worksheet_AP1_South.Cells[FenSep_S + 1, 8]];
                mergeRange_S_Fen7.Merge();
                mergeRange_S_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen7.Value = "3.4*Af*Uf";
                mergeRange_S_Fen7.Font.Name = "Times New Roman";
                mergeRange_S_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                

                Excel.Range mergeRange_S_Fen8 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep_S + 1, 9], worksheet_AP1_South.Cells[FenSep_S + 1, 10]];
                mergeRange_S_Fen8.Merge();
                mergeRange_S_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_S_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_S_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_S_Fen8.Value = "211*Af*SC*CF";
                mergeRange_S_Fen8.Font.Name = "Times New Roman";
                mergeRange_S_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                
                // Set the column width for Brief Description
                worksheet_AP1_South.Columns[3].ColumnWidth = 30;

                int FenSep1_S = FenSep_S + 2;


                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView3.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 1] = dataGridView3.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_S = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep1_S + i, 2], worksheet_AP1_South.Cells[FenSep1_S + i, 3]];
                            mergeRangeFenValue1_S.Merge();
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 1] = dataGridView3.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2] = dataGridView3.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2] = dataGridView3.Rows[i].Cells[j].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2] = dataGridView3.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_S = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep1_S + i, 7], worksheet_AP1_South.Cells[FenSep1_S + i, 8]];
                            mergeRangeFenValue2_S.Merge();
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2] = dataGridView3.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_S = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep1_S + i, 9], worksheet_AP1_South.Cells[FenSep1_S + i, 10]];
                            mergeRangeFenValue3_S.Merge();
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 3] = dataGridView3.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_South.Cells[FenSep1_S + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_S = FenSep1_S + dataGridView3.Rows.Count + 1;

                Excel.Range mergeRange_S_ETTV1 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep2_S, 2], worksheet_AP1_South.Cells[FenSep2_S, 3]];
                mergeRange_S_ETTV1.Merge();
                mergeRange_S_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_S_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_S_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_S_ETTV1.Font.Bold = true;
                worksheet_AP1_South.Cells[FenSep2_S, 4] = Lb_Area_S.Text;
                worksheet_AP1_South.Cells[FenSep2_S, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_South.Cells[FenSep2_S, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_South.Cells[FenSep2_S, 4].Font.Bold = true;
                worksheet_AP1_South.Cells[FenSep2_S, 5] = "m2";
                worksheet_AP1_South.Cells[FenSep2_S, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_South.Cells[FenSep2_S, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_South.Cells[FenSep2_S, 5].Font.Bold = true;

                Excel.Range mergeRange_S_ETTV2 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep2_S + 2, 2], worksheet_AP1_South.Cells[FenSep2_S + 2, 3]];
                mergeRange_S_ETTV2.Merge();
                mergeRange_S_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_S_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_S_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_S_ETTV2.Font.Bold = true;
                worksheet_AP1_South.Cells[FenSep2_S + 2, 4] = Lb_HG_S.Text;
                worksheet_AP1_South.Cells[FenSep2_S + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_South.Cells[FenSep2_S + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_South.Cells[FenSep2_S + 2, 4].Font.Bold = true;
                worksheet_AP1_South.Cells[FenSep2_S + 2, 5] = "W";
                worksheet_AP1_South.Cells[FenSep2_S + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_South.Cells[FenSep2_S + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_South.Cells[FenSep2_S+2, 5].Font.Bold = true;


                Excel.Range mergeRange_S_ETTV3 = worksheet_AP1_South.Range[worksheet_AP1_South.Cells[FenSep2_S + 4, 2], worksheet_AP1_South.Cells[FenSep2_S + 4, 3]];
                mergeRange_S_ETTV3.Merge();
                mergeRange_S_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_S_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_S_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_N_ETTV3.Font.Bold = true;
                worksheet_AP1_South.Cells[FenSep2_S + 4, 4] = Lb_ETTV_S.Text;
                worksheet_AP1_South.Cells[FenSep2_S + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_South.Cells[FenSep2_S + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_South.Cells[FenSep2_S + 4, 4].Font.Bold = true;
                worksheet_AP1_South.Cells[FenSep2_S + 4, 5] = "W / m2";
                worksheet_AP1_South.Cells[FenSep2_S + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_South.Cells[FenSep2_S + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_South.Cells[FenSep2_S + 4, 5].Font.Bold = true;

                ///// For AP1_South Ends/////
                ///

                ///// For AP1_East Starts/////

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_East = workbook.Sheets["AP1_East"];

                Excel.Range mergeRange_E2 = worksheet_AP1_East.Range["A2", "J2"];
                mergeRange_E2.Merge();
                mergeRange_E2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_E2.Font.Name = "Times New Roman";
                mergeRange_E2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_E2.Font.Bold = true;

                Excel.Range mergeRange_E3 = worksheet_AP1_East.Range["A3", "J3"];
                mergeRange_E3.Merge();
                mergeRange_E3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_E3.Font.Name = "Times New Roman";
                mergeRange_E3.Value = "FAÇADE ORIENTATION : E";
                mergeRange_E3.Font.Bold = true;

                //////////// for Opaque Walls

                Excel.Range mergeRange_E5 = worksheet_AP1_East.Range["A5", "J5"];
                mergeRange_E5.Merge();
                mergeRange_E5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_E5.Font.Name = "Times New Roman";
                mergeRange_E5.Value = "OPAQUE WALLS";
                mergeRange_E5.Font.Bold = true;

                Excel.Range mergeRange_E_6_1 = worksheet_AP1_East.Cells[6, 1];
                mergeRange_E_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_6_1.Value = "S/No";
                mergeRange_E_6_1.Font.Name = "Times New Roman";
                mergeRange_E_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_E_6_2_5 = worksheet_AP1_East.Range["B6", "E6"];
                mergeRange_E_6_2_5.Merge();
                mergeRange_E_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_6_2_5.Value = "Brief Description";
                mergeRange_E_6_2_5.Font.Name = "Times New Roman";
                mergeRange_E_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_E_6_6 = worksheet_AP1_East.Cells[6, 6];
                mergeRange_E_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_6_6.Value = "Aw";
                mergeRange_E_6_6.Font.Name = "Times New Roman";
                mergeRange_E_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_E_6_7 = worksheet_AP1_East.Cells[6, 7];
                mergeRange_E_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_6_7.Value = "Uw";
                mergeRange_E_6_7.Font.Name = "Times New Roman";
                mergeRange_E_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_E_6_8_9 = worksheet_AP1_East.Range["H6", "J6"];
                mergeRange_E_6_8_9.Merge();
                mergeRange_E_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_6_8_9.Value = "12*Aw*Uw";
                mergeRange_E_6_8_9.Font.Name = "Times New Roman";
                mergeRange_E_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_E_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_E_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;

                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_E = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_East.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_E.Merge();
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_E = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_East.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_E.Merge();
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_E = dataGrid_Wall_E_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_E_Fen1 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep_E, 1], worksheet_AP1_East.Cells[FenSep_E, 10]];
                mergeRange_E_Fen1.Merge();
                mergeRange_E_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_E_Fen1.Font.Name = "Times New Roman";
                mergeRange_E_Fen1.Value = "FENESTRATION";
                mergeRange_E_Fen1.Font.Bold = true;


                Excel.Range mergeRange_E_Fen2 = worksheet_AP1_East.Cells[(FenSep_E + 1), 1];
                mergeRange_E_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen2.Value = "S/No";
                mergeRange_E_Fen2.Font.Name = "Times New Roman";
                mergeRange_E_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_E_Fen3 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep_E + 1, 2], worksheet_AP1_East.Cells[FenSep_E + 1, 3]];
                mergeRange_E_Fen3.Merge();
                mergeRange_E_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen3.Value = "Brief Description";
                mergeRange_E_Fen3.Font.Name = "Times New Roman";
                mergeRange_E_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_E_Fen4 = worksheet_AP1_East.Cells[(FenSep_E + 1), 4];
                mergeRange_E_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen4.Value = "Af";
                mergeRange_E_Fen4.Font.Name = "Times New Roman";
                mergeRange_E_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_E_Fen5 = worksheet_AP1_East.Cells[(FenSep_E + 1), 5];
                mergeRange_E_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen5.Value = "Uf";
                mergeRange_E_Fen5.Font.Name = "Times New Roman";
                mergeRange_E_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_E_Fen6 = worksheet_AP1_East.Cells[(FenSep_E + 1), 6];
                mergeRange_E_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen6.Value = "SC";
                mergeRange_E_Fen6.Font.Name = "Times New Roman";
                mergeRange_E_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_E_Fen7 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep_E + 1, 7], worksheet_AP1_East.Cells[FenSep_E + 1, 8]];
                mergeRange_E_Fen7.Merge();
                mergeRange_E_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen7.Value = "3.4*Af*Uf";
                mergeRange_E_Fen7.Font.Name = "Times New Roman";
                mergeRange_E_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_E_Fen8 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep_E + 1, 9], worksheet_AP1_East.Cells[FenSep_E + 1, 10]];
                mergeRange_E_Fen8.Merge();
                mergeRange_E_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_E_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_E_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_E_Fen8.Value = "211*Af*SC*CF";
                mergeRange_E_Fen8.Font.Name = "Times New Roman";
                mergeRange_E_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set the column width for Brief Description
                worksheet_AP1_East.Columns[3].ColumnWidth = 30;

                int FenSep1_E = FenSep_E + 2;

                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGrid_Wndw_E_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_E_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 1] = dataGrid_Wndw_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_E = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep1_E + i, 2], worksheet_AP1_East.Cells[FenSep1_E + i, 3]];
                            mergeRangeFenValue1_E.Merge();
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 1] = dataGrid_Wndw_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2] = dataGrid_Wndw_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2] = dataGrid_Wndw_E_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2] = dataGrid_Wndw_E_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_E = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep1_E + i, 7], worksheet_AP1_East.Cells[FenSep1_E + i, 8]];
                            mergeRangeFenValue2_E.Merge();
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2] = dataGrid_Wndw_E_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_E = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep1_E + i, 9], worksheet_AP1_East.Cells[FenSep1_E + i, 10]];
                            mergeRangeFenValue3_E.Merge();
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 3] = dataGrid_Wndw_E_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_East.Cells[FenSep1_E + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_E = FenSep1_E + dataGrid_Wndw_E_S.Rows.Count + 1;

                Excel.Range mergeRange_E_ETTV1 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep2_E, 2], worksheet_AP1_East.Cells[FenSep2_E, 3]];
                mergeRange_E_ETTV1.Merge();
                mergeRange_E_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_E_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_E_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_E_ETTV1.Font.Bold = true;
                worksheet_AP1_East.Cells[FenSep2_E, 4] = Lb_Area_E.Text;
                worksheet_AP1_East.Cells[FenSep2_E, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_East.Cells[FenSep2_E, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_East.Cells[FenSep2_E, 4].Font.Bold = true;
                worksheet_AP1_East.Cells[FenSep2_E, 5] = "m2";
                worksheet_AP1_East.Cells[FenSep2_E, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_East.Cells[FenSep2_E, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_South.Cells[FenSep2_S, 5].Font.Bold = true;

                Excel.Range mergeRange_E_ETTV2 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep2_E + 2, 2], worksheet_AP1_East.Cells[FenSep2_E + 2, 3]];
                mergeRange_E_ETTV2.Merge();
                mergeRange_E_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_E_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_E_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_E_ETTV2.Font.Bold = true;
                worksheet_AP1_East.Cells[FenSep2_E + 2, 4] = Lb_HG_E.Text;
                worksheet_AP1_East.Cells[FenSep2_E + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_East.Cells[FenSep2_E + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_East.Cells[FenSep2_E + 2, 4].Font.Bold = true;
                worksheet_AP1_East.Cells[FenSep2_E + 2, 5] = "W";
                worksheet_AP1_East.Cells[FenSep2_E + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_East.Cells[FenSep2_E + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_East.Cells[FenSep2_S+2, 5].Font.Bold = true;


                Excel.Range mergeRange_E_ETTV3 = worksheet_AP1_East.Range[worksheet_AP1_East.Cells[FenSep2_E + 4, 2], worksheet_AP1_East.Cells[FenSep2_E + 4, 3]];
                mergeRange_E_ETTV3.Merge();
                mergeRange_E_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_E_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_E_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_E_ETTV3.Font.Bold = true;
                worksheet_AP1_East.Cells[FenSep2_E + 4, 4] = Lb_ETTV_E.Text;
                worksheet_AP1_East.Cells[FenSep2_E + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_East.Cells[FenSep2_E + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_East.Cells[FenSep2_E + 4, 4].Font.Bold = true;
                worksheet_AP1_East.Cells[FenSep2_E + 4, 5] = "W / m2";
                worksheet_AP1_East.Cells[FenSep2_E + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_East.Cells[FenSep2_E + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_East.Cells[FenSep2_E + 4, 5].Font.Bold = true;

                ///// For AP1_East Ends/////
                ///

               
                ///// For AP1_West Starts/////

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_West = workbook.Sheets["AP1_West"];

                Excel.Range mergeRange_W2 = worksheet_AP1_West.Range["A2", "J2"];
                mergeRange_W2.Merge();
                mergeRange_W2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_W2.Font.Name = "Times New Roman";
                mergeRange_W2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_W2.Font.Bold = true;

                Excel.Range mergeRange_W3 = worksheet_AP1_West.Range["A3", "J3"];
                mergeRange_W3.Merge();
                mergeRange_W3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_W3.Font.Name = "Times New Roman";
                mergeRange_W3.Value = "FAÇADE ORIENTATION : W";
                mergeRange_W3.Font.Bold = true;

                //////////// for Opaque Walls

                Excel.Range mergeRange_W5 = worksheet_AP1_West.Range["A5", "J5"];
                mergeRange_W5.Merge();
                mergeRange_W5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_W5.Font.Name = "Times New Roman";
                mergeRange_W5.Value = "OPAQUE WALLS";
                mergeRange_W5.Font.Bold = true;

                Excel.Range mergeRange_W_6_1 = worksheet_AP1_West.Cells[6, 1];
                mergeRange_W_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_6_1.Value = "S/No";
                mergeRange_W_6_1.Font.Name = "Times New Roman";
                mergeRange_W_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_W_6_2_5 = worksheet_AP1_West.Range["B6", "E6"];
                mergeRange_W_6_2_5.Merge();
                mergeRange_W_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_6_2_5.Value = "Brief Description";
                mergeRange_W_6_2_5.Font.Name = "Times New Roman";
                mergeRange_W_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_W_6_6 = worksheet_AP1_West.Cells[6, 6];
                mergeRange_W_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_6_6.Value = "Aw";
                mergeRange_W_6_6.Font.Name = "Times New Roman";
                mergeRange_W_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_W_6_7 = worksheet_AP1_West.Cells[6, 7];
                mergeRange_W_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_6_7.Value = "Uw";
                mergeRange_W_6_7.Font.Name = "Times New Roman";
                mergeRange_W_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_W_6_8_9 = worksheet_AP1_West.Range["H6", "J6"];
                mergeRange_W_6_8_9.Merge();
                mergeRange_W_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_6_8_9.Value = "12*Aw*Uw";
                mergeRange_W_6_8_9.Font.Name = "Times New Roman";
                mergeRange_W_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_W_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_W_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_W = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_West.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_W.Merge();
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_W = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_West.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_W.Merge();
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_W = dataGrid_Wall_W_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_W_Fen1 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep_W, 1], worksheet_AP1_West.Cells[FenSep_W, 10]];
                mergeRange_W_Fen1.Merge();
                mergeRange_W_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_W_Fen1.Font.Name = "Times New Roman";
                mergeRange_W_Fen1.Value = "FENESTRATION";
                mergeRange_W_Fen1.Font.Bold = true;


                Excel.Range mergeRange_W_Fen2 = worksheet_AP1_West.Cells[(FenSep_W + 1), 1];
                mergeRange_W_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen2.Value = "S/No";
                mergeRange_W_Fen2.Font.Name = "Times New Roman";
                mergeRange_W_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_W_Fen3 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep_W + 1, 2], worksheet_AP1_West.Cells[FenSep_W + 1, 3]];
                mergeRange_W_Fen3.Merge();
                mergeRange_W_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen3.Value = "Brief Description";
                mergeRange_W_Fen3.Font.Name = "Times New Roman";
                mergeRange_W_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_W_Fen4 = worksheet_AP1_West.Cells[(FenSep_W + 1), 4];
                mergeRange_W_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen4.Value = "Af";
                mergeRange_W_Fen4.Font.Name = "Times New Roman";
                mergeRange_W_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_W_Fen5 = worksheet_AP1_West.Cells[(FenSep_W + 1), 5];
                mergeRange_W_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen5.Value = "Uf";
                mergeRange_W_Fen5.Font.Name = "Times New Roman";
                mergeRange_W_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_W_Fen6 = worksheet_AP1_West.Cells[(FenSep_W + 1), 6];
                mergeRange_W_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen6.Value = "SC";
                mergeRange_W_Fen6.Font.Name = "Times New Roman";
                mergeRange_W_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_W_Fen7 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep_W + 1, 7], worksheet_AP1_West.Cells[FenSep_W + 1, 8]];
                mergeRange_W_Fen7.Merge();
                mergeRange_W_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen7.Value = "3.4*Af*Uf";
                mergeRange_W_Fen7.Font.Name = "Times New Roman";
                mergeRange_W_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_W_Fen8 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep_W + 1, 9], worksheet_AP1_West.Cells[FenSep_W + 1, 10]];
                mergeRange_W_Fen8.Merge();
                mergeRange_W_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_W_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_W_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_W_Fen8.Value = "211*Af*SC*CF";
                mergeRange_W_Fen8.Font.Name = "Times New Roman";
                mergeRange_W_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set the column width for Brief Description
                worksheet_AP1_West.Columns[3].ColumnWidth = 30;

                int FenSep1_W = FenSep_W + 2;

                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGrid_Wndw_W_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_W_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 1] = dataGrid_Wndw_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_W = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep1_W + i, 2], worksheet_AP1_West.Cells[FenSep1_W + i, 3]];
                            mergeRangeFenValue1_W.Merge();
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 1] = dataGrid_Wndw_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2] = dataGrid_Wndw_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2] = dataGrid_Wndw_W_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2] = dataGrid_Wndw_W_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_W = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep1_W + i, 7], worksheet_AP1_West.Cells[FenSep1_W + i, 8]];
                            mergeRangeFenValue2_W.Merge();
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2] = dataGrid_Wndw_W_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_W = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep1_W + i, 9], worksheet_AP1_West.Cells[FenSep1_W + i, 10]];
                            mergeRangeFenValue3_W.Merge();
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 3] = dataGrid_Wndw_W_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_West.Cells[FenSep1_W + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_W = FenSep1_W + dataGrid_Wndw_W_S.Rows.Count + 1;

                Excel.Range mergeRange_W_ETTV1 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep2_W, 2], worksheet_AP1_West.Cells[FenSep2_W, 3]];
                mergeRange_W_ETTV1.Merge();
                mergeRange_W_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_W_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_W_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_W_ETTV1.Font.Bold = true;
                worksheet_AP1_West.Cells[FenSep2_W, 4] = Lb_Area_W.Text;
                worksheet_AP1_West.Cells[FenSep2_W, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_West.Cells[FenSep2_W, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_West.Cells[FenSep2_W, 4].Font.Bold = true;
                worksheet_AP1_West.Cells[FenSep2_W, 5] = "m2";
                worksheet_AP1_West.Cells[FenSep2_W, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_West.Cells[FenSep2_W, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_West.Cells[FenSep2_W, 5].Font.Bold = true;

                Excel.Range mergeRange_W_ETTV2 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep2_W + 2, 2], worksheet_AP1_West.Cells[FenSep2_W + 2, 3]];
                mergeRange_W_ETTV2.Merge();
                mergeRange_W_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_W_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_W_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_W_ETTV2.Font.Bold = true;
                worksheet_AP1_West.Cells[FenSep2_W + 2, 4] = Lb_HG_W.Text;
                worksheet_AP1_West.Cells[FenSep2_W + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_West.Cells[FenSep2_W + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_West.Cells[FenSep2_W + 2, 4].Font.Bold = true;
                worksheet_AP1_West.Cells[FenSep2_W + 2, 5] = "W";
                worksheet_AP1_West.Cells[FenSep2_W + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_West.Cells[FenSep2_W + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_West.Cells[FenSep2_W+2, 5].Font.Bold = true;


                Excel.Range mergeRange_W_ETTV3 = worksheet_AP1_West.Range[worksheet_AP1_West.Cells[FenSep2_W + 4, 2], worksheet_AP1_West.Cells[FenSep2_W + 4, 3]];
                mergeRange_W_ETTV3.Merge();
                mergeRange_W_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_W_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_W_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_W_ETTV3.Font.Bold = true;
                worksheet_AP1_West.Cells[FenSep2_W + 4, 4] = Lb_ETTV_W.Text;
                worksheet_AP1_West.Cells[FenSep2_W + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_West.Cells[FenSep2_W + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_West.Cells[FenSep2_W + 4, 4].Font.Bold = true;
                worksheet_AP1_West.Cells[FenSep2_W + 4, 5] = "W / m2";
                worksheet_AP1_West.Cells[FenSep2_W + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_West.Cells[FenSep2_W + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_West.Cells[FenSep2_W + 4, 5].Font.Bold = true;

                ///// For AP1_West Ends/////
                ///

                
                ///// For AP1_NorthEast Starts/////

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_NorthEast = workbook.Sheets["AP1_North East"];

                Excel.Range mergeRange_NE2 = worksheet_AP1_NorthEast.Range["A2", "J2"];
                mergeRange_NE2.Merge();
                mergeRange_NE2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NE2.Font.Name = "Times New Roman";
                mergeRange_NE2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_NE2.Font.Bold = true;

                Excel.Range mergeRange_NE3 = worksheet_AP1_NorthEast.Range["A3", "J3"];
                mergeRange_NE3.Merge();
                mergeRange_NE3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NE3.Font.Name = "Times New Roman";
                mergeRange_NE3.Value = "FAÇADE ORIENTATION : NE";
                mergeRange_NE3.Font.Bold = true;

                //////////// for Opaque Walls

                Excel.Range mergeRange_NE5 = worksheet_AP1_NorthEast.Range["A5", "J5"];
                mergeRange_NE5.Merge();
                mergeRange_NE5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NE5.Font.Name = "Times New Roman";
                mergeRange_NE5.Value = "OPAQUE WALLS";
                mergeRange_NE5.Font.Bold = true;

                Excel.Range mergeRange_NE_6_1 = worksheet_AP1_NorthEast.Cells[6, 1];
                mergeRange_NE_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_6_1.Value = "S/No";
                mergeRange_NE_6_1.Font.Name = "Times New Roman";
                mergeRange_NE_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NE_6_2_5 = worksheet_AP1_NorthEast.Range["B6", "E6"];
                mergeRange_NE_6_2_5.Merge();
                mergeRange_NE_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_6_2_5.Value = "Brief Description";
                mergeRange_NE_6_2_5.Font.Name = "Times New Roman";
                mergeRange_NE_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NE_6_6 = worksheet_AP1_NorthEast.Cells[6, 6];
                mergeRange_NE_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_6_6.Value = "Aw";
                mergeRange_NE_6_6.Font.Name = "Times New Roman";
                mergeRange_NE_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NE_6_7 = worksheet_AP1_NorthEast.Cells[6, 7];
                mergeRange_NE_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_6_7.Value = "Uw";
                mergeRange_NE_6_7.Font.Name = "Times New Roman";
                mergeRange_NE_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NE_6_8_9 = worksheet_AP1_NorthEast.Range["H6", "J6"];
                mergeRange_NE_6_8_9.Merge();
                mergeRange_NE_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_6_8_9.Value = "12*Aw*Uw";
                mergeRange_NE_6_8_9.Font.Name = "Times New Roman";
                mergeRange_NE_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_NE_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_NE_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_NE = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_NE.Merge();
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_NE = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_NE.Merge();
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_NE = dataGrid_Wall_NE_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_NE_Fen1 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep_NE, 1], worksheet_AP1_NorthEast.Cells[FenSep_NE, 10]];
                mergeRange_NE_Fen1.Merge();
                mergeRange_NE_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NE_Fen1.Font.Name = "Times New Roman";
                mergeRange_NE_Fen1.Value = "FENESTRATION";
                mergeRange_NE_Fen1.Font.Bold = true;


                Excel.Range mergeRange_NE_Fen2 = worksheet_AP1_NorthEast.Cells[(FenSep_NE + 1), 1];
                mergeRange_NE_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen2.Value = "S/No";
                mergeRange_NE_Fen2.Font.Name = "Times New Roman";
                mergeRange_NE_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NE_Fen3 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep_NE + 1, 2], worksheet_AP1_NorthEast.Cells[FenSep_NE + 1, 3]];
                mergeRange_NE_Fen3.Merge();
                mergeRange_NE_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen3.Value = "Brief Description";
                mergeRange_NE_Fen3.Font.Name = "Times New Roman";
                mergeRange_NE_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NE_Fen4 = worksheet_AP1_NorthEast.Cells[(FenSep_NE + 1), 4];
                mergeRange_NE_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen4.Value = "Af";
                mergeRange_NE_Fen4.Font.Name = "Times New Roman";
                mergeRange_NE_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NE_Fen5 = worksheet_AP1_NorthEast.Cells[(FenSep_NE + 1), 5];
                mergeRange_NE_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen5.Value = "Uf";
                mergeRange_NE_Fen5.Font.Name = "Times New Roman";
                mergeRange_NE_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NE_Fen6 = worksheet_AP1_NorthEast.Cells[(FenSep_NE + 1), 6];
                mergeRange_NE_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen6.Value = "SC";
                mergeRange_NE_Fen6.Font.Name = "Times New Roman";
                mergeRange_NE_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NE_Fen7 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep_NE + 1, 7], worksheet_AP1_NorthEast.Cells[FenSep_NE + 1, 8]];
                mergeRange_NE_Fen7.Merge();
                mergeRange_NE_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen7.Value = "3.4*Af*Uf";
                mergeRange_NE_Fen7.Font.Name = "Times New Roman";
                mergeRange_NE_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NE_Fen8 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep_NE + 1, 9], worksheet_AP1_NorthEast.Cells[FenSep_NE + 1, 10]];
                mergeRange_NE_Fen8.Merge();
                mergeRange_NE_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NE_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NE_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NE_Fen8.Value = "211*Af*SC*CF";
                mergeRange_NE_Fen8.Font.Name = "Times New Roman";
                mergeRange_NE_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set the column width for Brief Description
                worksheet_AP1_NorthEast.Columns[3].ColumnWidth = 30;

                int FenSep1_NE = FenSep_NE + 2;

                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGrid_Wndw_NE_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_NE_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 1] = dataGrid_Wndw_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_NE = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, 2], worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, 3]];
                            mergeRangeFenValue1_NE.Merge();
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 1] = dataGrid_Wndw_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2] = dataGrid_Wndw_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2] = dataGrid_Wndw_NE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2] = dataGrid_Wndw_NE_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_NE = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, 7], worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, 8]];
                            mergeRangeFenValue2_NE.Merge();
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2] = dataGrid_Wndw_NE_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_NE = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, 9], worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, 10]];
                            mergeRangeFenValue3_NE.Merge();
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 3] = dataGrid_Wndw_NE_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthEast.Cells[FenSep1_NE + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_NE = FenSep1_NE + dataGrid_Wndw_NE_S.Rows.Count + 1;

                Excel.Range mergeRange_NE_ETTV1 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep2_NE, 2], worksheet_AP1_NorthEast.Cells[FenSep2_NE, 3]];
                mergeRange_NE_ETTV1.Merge();
                mergeRange_NE_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_NE_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_NE_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_NE_ETTV1.Font.Bold = true;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE, 4] = Lb_Area_NE.Text;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthEast.Cells[FenSep2_NE, 4].Font.Bold = true;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE, 5] = "m2";
                worksheet_AP1_NorthEast.Cells[FenSep2_NE, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthEast.Cells[FenSep2_W, 5].Font.Bold = true;

                Excel.Range mergeRange_NE_ETTV2 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 2], worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 3]];
                mergeRange_NE_ETTV2.Merge();
                mergeRange_NE_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_NE_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_NE_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_NE_ETTV2.Font.Bold = true;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 4] = Lb_HG_NE.Text;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthEast.Cells[FenSep2_W + 2, 4].Font.Bold = true;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 5] = "W";
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthEast.Cells[FenSepNE_W+2, 5].Font.Bold = true;


                Excel.Range mergeRange_NE_ETTV3 = worksheet_AP1_NorthEast.Range[worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 2], worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 3]];
                mergeRange_NE_ETTV3.Merge();
                mergeRange_NE_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_NE_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_NE_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_NE_ETTV3.Font.Bold = true;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 4] = Lb_ETTV_NE.Text;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 4].Font.Bold = true;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 5] = "W / m2";
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthEast.Cells[FenSep2_NE + 4, 5].Font.Bold = true;



                ///// For AP1_NorthEast Ends/////
                ///


                ///// For AP1_NorthWest Starts/////              

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_NorthWest = workbook.Sheets["AP1_North West"];

                Excel.Range mergeRange_NW2 = worksheet_AP1_NorthWest.Range["A2", "J2"];
                mergeRange_NW2.Merge();
                mergeRange_NW2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NW2.Font.Name = "Times New Roman";
                mergeRange_NW2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_NW2.Font.Bold = true;

                Excel.Range mergeRange_NW3 = worksheet_AP1_NorthWest.Range["A3", "J3"];
                mergeRange_NW3.Merge();
                mergeRange_NW3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NW3.Font.Name = "Times New Roman";
                mergeRange_NW3.Value = "FAÇADE ORIENTATION : NW";
                mergeRange_NW3.Font.Bold = true;

                //////////// for Opaque Walls

                Excel.Range mergeRange_NW5 = worksheet_AP1_NorthWest.Range["A5", "J5"];
                mergeRange_NW5.Merge();
                mergeRange_NW5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NW5.Font.Name = "Times New Roman";
                mergeRange_NW5.Value = "OPAQUE WALLS";
                mergeRange_NW5.Font.Bold = true;

                Excel.Range mergeRange_NW_6_1 = worksheet_AP1_NorthWest.Cells[6, 1];
                mergeRange_NW_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_6_1.Value = "S/No";
                mergeRange_NW_6_1.Font.Name = "Times New Roman";
                mergeRange_NW_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NW_6_2_5 = worksheet_AP1_NorthWest.Range["B6", "E6"];
                mergeRange_NW_6_2_5.Merge();
                mergeRange_NW_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_6_2_5.Value = "Brief Description";
                mergeRange_NW_6_2_5.Font.Name = "Times New Roman";
                mergeRange_NW_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NW_6_6 = worksheet_AP1_NorthWest.Cells[6, 6];
                mergeRange_NW_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_6_6.Value = "Aw";
                mergeRange_NW_6_6.Font.Name = "Times New Roman";
                mergeRange_NW_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NW_6_7 = worksheet_AP1_NorthWest.Cells[6, 7];
                mergeRange_NW_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_6_7.Value = "Uw";
                mergeRange_NW_6_7.Font.Name = "Times New Roman";
                mergeRange_NW_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NW_6_8_9 = worksheet_AP1_NorthWest.Range["H6", "J6"];
                mergeRange_NW_6_8_9.Merge();
                mergeRange_NW_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_6_8_9.Value = "12*Aw*Uw";
                mergeRange_NW_6_8_9.Font.Name = "Times New Roman";
                mergeRange_NW_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_NW_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_NW_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_NW = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_NW.Merge();
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_NW = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_NW.Merge();
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_NW = dataGrid_Wall_NW_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_NW_Fen1 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep_NW, 1], worksheet_AP1_NorthWest.Cells[FenSep_NW, 10]];
                mergeRange_NW_Fen1.Merge();
                mergeRange_NW_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_NW_Fen1.Font.Name = "Times New Roman";
                mergeRange_NW_Fen1.Value = "FENESTRATION";
                mergeRange_NW_Fen1.Font.Bold = true;


                Excel.Range mergeRange_NW_Fen2 = worksheet_AP1_NorthWest.Cells[(FenSep_NW + 1), 1];
                mergeRange_NW_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen2.Value = "S/No";
                mergeRange_NW_Fen2.Font.Name = "Times New Roman";
                mergeRange_NW_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NW_Fen3 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep_NW + 1, 2], worksheet_AP1_NorthWest.Cells[FenSep_NW + 1, 3]];
                mergeRange_NW_Fen3.Merge();
                mergeRange_NW_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen3.Value = "Brief Description";
                mergeRange_NW_Fen3.Font.Name = "Times New Roman";
                mergeRange_NW_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NW_Fen4 = worksheet_AP1_NorthWest.Cells[(FenSep_NW + 1), 4];
                mergeRange_NW_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen4.Value = "Af";
                mergeRange_NW_Fen4.Font.Name = "Times New Roman";
                mergeRange_NW_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_NW_Fen5 = worksheet_AP1_NorthWest.Cells[(FenSep_NW + 1), 5];
                mergeRange_NW_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen5.Value = "Uf";
                mergeRange_NW_Fen5.Font.Name = "Times New Roman";
                mergeRange_NW_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NW_Fen6 = worksheet_AP1_NorthWest.Cells[(FenSep_NW + 1), 6];
                mergeRange_NW_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen6.Value = "SC";
                mergeRange_NW_Fen6.Font.Name = "Times New Roman";
                mergeRange_NW_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NW_Fen7 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep_NW + 1, 7], worksheet_AP1_NorthWest.Cells[FenSep_NW + 1, 8]];
                mergeRange_NW_Fen7.Merge();
                mergeRange_NW_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen7.Value = "3.4*Af*Uf";
                mergeRange_NW_Fen7.Font.Name = "Times New Roman";
                mergeRange_NW_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_NW_Fen8 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep_NW + 1, 9], worksheet_AP1_NorthWest.Cells[FenSep_NW + 1, 10]];
                mergeRange_NW_Fen8.Merge();
                mergeRange_NW_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_NW_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_NW_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_NW_Fen8.Value = "211*Af*SC*CF";
                mergeRange_NW_Fen8.Font.Name = "Times New Roman";
                mergeRange_NW_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set the column width for Brief Description
                worksheet_AP1_NorthWest.Columns[3].ColumnWidth = 30;

                int FenSep1_NW = FenSep_NW + 2;

                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGrid_Wndw_NW_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_NW_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 1] = dataGrid_Wndw_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_NW = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, 2], worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, 3]];
                            mergeRangeFenValue1_NW.Merge();
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 1] = dataGrid_Wndw_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2] = dataGrid_Wndw_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2] = dataGrid_Wndw_NW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2] = dataGrid_Wndw_NW_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_NW = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, 7], worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, 8]];
                            mergeRangeFenValue2_NW.Merge();
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2] = dataGrid_Wndw_NW_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_NW = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, 9], worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, 10]];
                            mergeRangeFenValue3_NW.Merge();
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 3] = dataGrid_Wndw_NW_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_NorthWest.Cells[FenSep1_NW + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_NW = FenSep1_NW + dataGrid_Wndw_NW_S.Rows.Count + 1;

                Excel.Range mergeRange_NW_ETTV1 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep2_NW, 2], worksheet_AP1_NorthWest.Cells[FenSep2_NW, 3]];
                mergeRange_NW_ETTV1.Merge();
                mergeRange_NW_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_NW_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_NW_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_NW_ETTV1.Font.Bold = true;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW, 4] = Lb_Area_NW.Text;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthWest.Cells[FenSep2_NW, 4].Font.Bold = true;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW, 5] = "m2";
                worksheet_AP1_NorthWest.Cells[FenSep2_NW, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthWest.Cells[FenSep2_NW, 5].Font.Bold = true;

                Excel.Range mergeRange_NW_ETTV2 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 2], worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 3]];
                mergeRange_NW_ETTV2.Merge();
                mergeRange_NW_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_NW_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_NW_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_NW_ETTV2.Font.Bold = true;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 4] = Lb_HG_NW.Text;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 4].Font.Bold = true;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 5] = "W";
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthEast.Cells[FenSepNE_W+2, 5].Font.Bold = true;


                Excel.Range mergeRange_NW_ETTV3 = worksheet_AP1_NorthWest.Range[worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 2], worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 3]];
                mergeRange_NW_ETTV3.Merge();
                mergeRange_NW_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_NW_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_NW_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_NW_ETTV3.Font.Bold = true;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 4] = Lb_ETTV_NW.Text;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 4].Font.Bold = true;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 5] = "W / m2";
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_NorthWest.Cells[FenSep2_NW + 4, 5].Font.Bold = true;

                ///// For AP1_NorthWest Ends/////
                ///
                              

                ///// For AP1_SouthEast Starts/////              

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_SouthEast = workbook.Sheets["AP1_South East"];

                Excel.Range mergeRange_SE2 = worksheet_AP1_SouthEast.Range["A2", "J2"];
                mergeRange_SE2.Merge();
                mergeRange_SE2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SE2.Font.Name = "Times New Roman";
                mergeRange_SE2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_SE2.Font.Bold = true;

                Excel.Range mergeRange_SE3 = worksheet_AP1_SouthEast.Range["A3", "J3"];
                mergeRange_SE3.Merge();
                mergeRange_SE3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SE3.Font.Name = "Times New Roman";
                mergeRange_SE3.Value = "FAÇADE ORIENTATION : SE";
                mergeRange_SE3.Font.Bold = true;

                //////////// for Opaque Walls

                Excel.Range mergeRange_SE5 = worksheet_AP1_SouthEast.Range["A5", "J5"];
                mergeRange_SE5.Merge();
                mergeRange_SE5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SE5.Font.Name = "Times New Roman";
                mergeRange_SE5.Value = "OPAQUE WALLS";
                mergeRange_SE5.Font.Bold = true;

                Excel.Range mergeRange_SE_6_1 = worksheet_AP1_SouthEast.Cells[6, 1];
                mergeRange_SE_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_6_1.Value = "S/No";
                mergeRange_SE_6_1.Font.Name = "Times New Roman";
                mergeRange_SE_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SE_6_2_5 = worksheet_AP1_SouthEast.Range["B6", "E6"];
                mergeRange_SE_6_2_5.Merge();
                mergeRange_SE_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_6_2_5.Value = "Brief Description";
                mergeRange_SE_6_2_5.Font.Name = "Times New Roman";
                mergeRange_SE_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SE_6_6 = worksheet_AP1_SouthEast.Cells[6, 6];
                mergeRange_SE_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_6_6.Value = "Aw";
                mergeRange_SE_6_6.Font.Name = "Times New Roman";
                mergeRange_SE_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SE_6_7 = worksheet_AP1_SouthEast.Cells[6, 7];
                mergeRange_SE_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_6_7.Value = "Uw";
                mergeRange_SE_6_7.Font.Name = "Times New Roman";
                mergeRange_SE_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SE_6_8_9 = worksheet_AP1_SouthEast.Range["H6", "J6"];
                mergeRange_SE_6_8_9.Merge();
                mergeRange_SE_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_6_8_9.Value = "12*Aw*Uw";
                mergeRange_SE_6_8_9.Font.Name = "Times New Roman";
                mergeRange_SE_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_SE_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_SE_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_SE = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_SE.Merge();
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_SE = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_SE.Merge();
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_SE = dataGrid_Wall_SE_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_SE_Fen1 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep_SE, 1], worksheet_AP1_SouthEast.Cells[FenSep_SE, 10]];
                mergeRange_SE_Fen1.Merge();
                mergeRange_SE_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SE_Fen1.Font.Name = "Times New Roman";
                mergeRange_SE_Fen1.Value = "FENESTRATION";
                mergeRange_SE_Fen1.Font.Bold = true;


                Excel.Range mergeRange_SE_Fen2 = worksheet_AP1_SouthEast.Cells[(FenSep_SE + 1), 1];
                mergeRange_SE_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen2.Value = "S/No";
                mergeRange_SE_Fen2.Font.Name = "Times New Roman";
                mergeRange_SE_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SE_Fen3 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep_SE + 1, 2], worksheet_AP1_SouthEast.Cells[FenSep_SE + 1, 3]];
                mergeRange_SE_Fen3.Merge();
                mergeRange_SE_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen3.Value = "Brief Description";
                mergeRange_SE_Fen3.Font.Name = "Times New Roman";
                mergeRange_SE_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SE_Fen4 = worksheet_AP1_SouthEast.Cells[(FenSep_SE + 1), 4];
                mergeRange_SE_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen4.Value = "Af";
                mergeRange_SE_Fen4.Font.Name = "Times New Roman";
                mergeRange_SE_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SE_Fen5 = worksheet_AP1_SouthEast.Cells[(FenSep_SE + 1), 5];
                mergeRange_SE_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen5.Value = "Uf";
                mergeRange_SE_Fen5.Font.Name = "Times New Roman";
                mergeRange_SE_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SE_Fen6 = worksheet_AP1_SouthEast.Cells[(FenSep_SE + 1), 6];
                mergeRange_SE_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen6.Value = "SC";
                mergeRange_SE_Fen6.Font.Name = "Times New Roman";
                mergeRange_SE_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SE_Fen7 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep_SE + 1, 7], worksheet_AP1_SouthEast.Cells[FenSep_SE + 1, 8]];
                mergeRange_SE_Fen7.Merge();
                mergeRange_SE_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen7.Value = "3.4*Af*Uf";
                mergeRange_SE_Fen7.Font.Name = "Times New Roman";
                mergeRange_SE_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SE_Fen8 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep_SE + 1, 9], worksheet_AP1_SouthEast.Cells[FenSep_SE + 1, 10]];
                mergeRange_SE_Fen8.Merge();
                mergeRange_SE_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SE_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SE_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SE_Fen8.Value = "211*Af*SC*CF";
                mergeRange_SE_Fen8.Font.Name = "Times New Roman";
                mergeRange_SE_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set the column width for Brief Description
                worksheet_AP1_SouthEast.Columns[3].ColumnWidth = 30;

                int FenSep1_SE = FenSep_SE + 2;

                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGrid_Wndw_SE_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_SE_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 1] = dataGrid_Wndw_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_SE = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, 2], worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, 3]];
                            mergeRangeFenValue1_SE.Merge();
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 1] = dataGrid_Wndw_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2] = dataGrid_Wndw_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2] = dataGrid_Wndw_SE_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2] = dataGrid_Wndw_SE_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_SE = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, 7], worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, 8]];
                            mergeRangeFenValue2_SE.Merge();
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2] = dataGrid_Wndw_SE_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_SE = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, 9], worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, 10]];
                            mergeRangeFenValue3_SE.Merge();
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 3] = dataGrid_Wndw_SE_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthEast.Cells[FenSep1_SE + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_SE = FenSep1_SE + dataGrid_Wndw_SE_S.Rows.Count + 1;

                Excel.Range mergeRange_SE_ETTV1 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep2_SE, 2], worksheet_AP1_SouthEast.Cells[FenSep2_SE, 3]];
                mergeRange_SE_ETTV1.Merge();
                mergeRange_SE_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_SE_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_SE_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_SE_ETTV1.Font.Bold = true;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE, 4] = Lb_Area_SE.Text;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthEast.Cells[FenSep2_SE, 4].Font.Bold = true;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE, 5] = "m2";
                worksheet_AP1_SouthEast.Cells[FenSep2_SE, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthEast.Cells[FenSep2_SE, 5].Font.Bold = true;

                Excel.Range mergeRange_SE_ETTV2 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 2], worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 3]];
                mergeRange_SE_ETTV2.Merge();
                mergeRange_SE_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_SE_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_SE_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_SE_ETTV2.Font.Bold = true;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 4] = Lb_HG_SE.Text;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 4].Font.Bold = true;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 5] = "W";
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthEast.Cells[FenSep2_SE+2, 5].Font.Bold = true;


                Excel.Range mergeRange_SE_ETTV3 = worksheet_AP1_SouthEast.Range[worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 2], worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 3]];
                mergeRange_SE_ETTV3.Merge();
                mergeRange_SE_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_SE_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_SE_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_SE_ETTV3.Font.Bold = true;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 4] = Lb_ETTV_SE.Text;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 4].Font.Bold = true;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 5] = "W / m2";
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthEast.Cells[FenSep2_SE + 4, 5].Font.Bold = true;

                ///// For AP1_SouthEast Ends/////
                ///
                                
                ///// For AP1_SouthWest Starts/////              

                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP1_SouthWest = workbook.Sheets["AP1_South West"];

                Excel.Range mergeRange_SW2 = worksheet_AP1_SouthWest.Range["A2", "J2"];
                mergeRange_SW2.Merge();
                mergeRange_SW2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SW2.Font.Name = "Times New Roman";
                mergeRange_SW2.Value = "CALCULATION OF RETV OF BUILDING ENVELOPE";
                mergeRange_SW2.Font.Bold = true;

                Excel.Range mergeRange_SW3 = worksheet_AP1_SouthWest.Range["A3", "J3"];
                mergeRange_SW3.Merge();
                mergeRange_SW3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SW3.Font.Name = "Times New Roman";
                mergeRange_SW3.Value = "FAÇADE ORIENTATION : SW";
                mergeRange_SW3.Font.Bold = true;

                //////////// for Opaque Walls
                ///

                Excel.Range mergeRange_SW5 = worksheet_AP1_SouthWest.Range["A5", "J5"];
                mergeRange_SW5.Merge();
                mergeRange_SW5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SW5.Font.Name = "Times New Roman";
                mergeRange_SW5.Value = "OPAQUE WALLS";
                mergeRange_SW5.Font.Bold = true;

                Excel.Range mergeRange_SW_6_1 = worksheet_AP1_SouthWest.Cells[6, 1];
                mergeRange_SW_6_1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_6_1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_6_1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_6_1.Value = "S/No";
                mergeRange_SW_6_1.Font.Name = "Times New Roman";
                mergeRange_SW_6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SW_6_2_5 = worksheet_AP1_SouthWest.Range["B6", "E6"];
                mergeRange_SW_6_2_5.Merge();
                mergeRange_SW_6_2_5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_6_2_5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_6_2_5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_6_2_5.Value = "Brief Description";
                mergeRange_SW_6_2_5.Font.Name = "Times New Roman";
                mergeRange_SW_6_2_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SW_6_6 = worksheet_AP1_SouthWest.Cells[6, 6];
                mergeRange_SW_6_6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_6_6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_6_6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_6_6.Value = "Aw";
                mergeRange_SW_6_6.Font.Name = "Times New Roman";
                mergeRange_SW_6_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SW_6_7 = worksheet_AP1_SouthWest.Cells[6, 7];
                mergeRange_SW_6_7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_6_7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_6_7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_6_7.Value = "Uw";
                mergeRange_SW_6_7.Font.Name = "Times New Roman";
                mergeRange_SW_6_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SW_6_8_9 = worksheet_AP1_SouthWest.Range["H6", "J6"];
                mergeRange_SW_6_8_9.Merge();
                mergeRange_SW_6_8_9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_6_8_9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_6_8_9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_6_8_9.Value = "12*Aw*Uw";
                mergeRange_SW_6_8_9.Font.Name = "Times New Roman";
                mergeRange_SW_6_8_9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////for Datagrid data to excel for opaque wall
                for (int i = 0; i < dataGrid_Wall_SW_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_SW_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 236, 255));
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            //worksheet_AP1_North.Cells[rowAdder1 + 1 + i, j + 1].Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }
                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue1_SW = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, 2], worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, 5]];
                            mergeRangeWallValue1_SW.Merge();
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 1] = dataGrid_Wall_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        }
                        if (j == 4) // j=4 prints 12*Aw*Uw
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallValue2_SW = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, 8], worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, 10]];
                            mergeRangeWallValue2_SW.Merge();
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4] = dataGrid_Wall_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[rowAdder1 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                }

                //////////// for Fenestration

                int FenSep_SW = dataGrid_Wall_SW_S.Rows.Count + rowAdder1 + 1; ;

                Excel.Range mergeRange_SW_Fen1 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep_SW, 1], worksheet_AP1_SouthWest.Cells[FenSep_SW, 10]];
                mergeRange_SW_Fen1.Merge();
                mergeRange_SW_Fen1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_SW_Fen1.Font.Name = "Times New Roman";
                mergeRange_SW_Fen1.Value = "FENESTRATION";
                mergeRange_SW_Fen1.Font.Bold = true;


                Excel.Range mergeRange_SW_Fen2 = worksheet_AP1_SouthWest.Cells[(FenSep_SW + 1), 1];
                mergeRange_SW_Fen2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen2.Value = "S/No";
                mergeRange_SW_Fen2.Font.Name = "Times New Roman";
                mergeRange_SW_Fen2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SW_Fen3 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep_SW + 1, 2], worksheet_AP1_SouthWest.Cells[FenSep_SW + 1, 3]];
                mergeRange_SW_Fen3.Merge();
                mergeRange_SW_Fen3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen3.Value = "Brief Description";
                mergeRange_SW_Fen3.Font.Name = "Times New Roman";
                mergeRange_SW_Fen3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SW_Fen4 = worksheet_AP1_SouthWest.Cells[(FenSep_SW + 1), 4];
                mergeRange_SW_Fen4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen4.Value = "Af";
                mergeRange_SW_Fen4.Font.Name = "Times New Roman";
                mergeRange_SW_Fen4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_SW_Fen5 = worksheet_AP1_SouthWest.Cells[(FenSep_SW + 1), 5];
                mergeRange_SW_Fen5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen5.Value = "Uf";
                mergeRange_SW_Fen5.Font.Name = "Times New Roman";
                mergeRange_SW_Fen5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SW_Fen6 = worksheet_AP1_SouthWest.Cells[(FenSep_SW + 1), 6];
                mergeRange_SW_Fen6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen6.Value = "SC";
                mergeRange_SW_Fen6.Font.Name = "Times New Roman";
                mergeRange_SW_Fen6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SW_Fen7 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep_SW + 1, 7], worksheet_AP1_SouthWest.Cells[FenSep_SW + 1, 8]];
                mergeRange_SW_Fen7.Merge();
                mergeRange_SW_Fen7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen7.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen7.Value = "3.4*Af*Uf";
                mergeRange_SW_Fen7.Font.Name = "Times New Roman";
                mergeRange_SW_Fen7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_SW_Fen8 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep_SW + 1, 9], worksheet_AP1_SouthWest.Cells[FenSep_SW + 1, 10]];
                mergeRange_SW_Fen8.Merge();
                mergeRange_SW_Fen8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_SW_Fen8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_SW_Fen8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_SW_Fen8.Value = "211*Af*SC*CF";
                mergeRange_SW_Fen8.Font.Name = "Times New Roman";
                mergeRange_SW_Fen8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set the column width for Brief Description
                worksheet_AP1_SouthWest.Columns[3].ColumnWidth = 30;

                int FenSep1_SW = FenSep_SW + 2;


                //////for Datagrid data to excel for fenestration
                for (int i = 0; i < dataGrid_Wndw_SW_S.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_SW_S.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 1] = dataGrid_Wndw_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue1_SW = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, 2], worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, 3]];
                            mergeRangeFenValue1_SW.Merge();
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 1] = dataGrid_Wndw_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints Area(m2)
                        {
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2] = dataGrid_Wndw_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3 prints U Values
                        {
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2] = dataGrid_Wndw_SW_S.Rows[i].Cells[j].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints SC Values
                        {
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2] = dataGrid_Wndw_SW_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SE + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints 3.4*Af*Uf Value
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue2_SW = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, 7], worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, 8]];
                            mergeRangeFenValue2_SW.Merge();
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2] = dataGrid_Wndw_SW_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints 211*Af*SC*CF 
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeFenValue3_SW = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, 9], worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, 10]];
                            mergeRangeFenValue3_SW.Merge();
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 3] = dataGrid_Wndw_SW_S.Rows[i].Cells[j + 2].Value;
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP1_SouthWest.Cells[FenSep1_SW + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                    }
                }

                int FenSep2_SW = FenSep1_SW + dataGrid_Wndw_SW_S.Rows.Count + 1;

                Excel.Range mergeRange_SW_ETTV1 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep2_SW, 2], worksheet_AP1_SouthWest.Cells[FenSep2_SW, 3]];
                mergeRange_SW_ETTV1.Merge();
                mergeRange_SW_ETTV1.Value = "Gross Area Of External Walls (Ao):";
                mergeRange_SW_ETTV1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_SW_ETTV1.Font.Name = "Times New Roman";
                //mergeRange_SW_ETTV1.Font.Bold = true;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW, 4] = Lb_Area_SW.Text;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthWest.Cells[FenSep2_SW, 4].Font.Bold = true;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW, 5] = "m2";
                worksheet_AP1_SouthWest.Cells[FenSep2_SW, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthWest.Cells[FenSep2_SW, 5].Font.Bold = true;

                Excel.Range mergeRange_SW_ETTV2 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 2], worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 3]];
                mergeRange_SW_ETTV2.Merge();
                mergeRange_SW_ETTV2.Value = "Gross Heat Gain:";
                mergeRange_SW_ETTV2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_SW_ETTV2.Font.Name = "Times New Roman";
                //mergeRange_SW_ETTV2.Font.Bold = true;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 4] = Lb_HG_SW.Text;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 4].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 4].Font.Bold = true;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 5] = "W";
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 2, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthWest.Cells[FenSep2_SW+2, 5].Font.Bold = true;


                Excel.Range mergeRange_SW_ETTV3 = worksheet_AP1_SouthWest.Range[worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 2], worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 3]];
                mergeRange_SW_ETTV3.Merge();
                mergeRange_SW_ETTV3.Value = "ETTV = Gross Heat Gain/Gross Area:";
                mergeRange_SW_ETTV3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                mergeRange_SW_ETTV3.Font.Name = "Times New Roman";
                //mergeRange_SW_ETTV3.Font.Bold = true;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 4] = Lb_ETTV_SW.Text;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 4].Font.Name = "Times New Roman";
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 4].Font.Bold = true;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 5] = "W / m2";
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 5].Font.Name = "Times New Roman";
                //worksheet_AP1_SouthWest.Cells[FenSep2_SW + 4, 5].Font.Bold = true;


                ///// For AP1_SouthWest Ends/////
                ///


                ///////////////////////////////////////////////////////////// for AP1 Worksheets - Ends ////////////////////////////////////////////////////////////////


                ///////////////////////////////////////////////////////////// for AP3 Worksheets - Starts ////////////////////////////////////////////////////////////////
                AddWorksheetsForAP3Tabs();

                ///// For AP3_Window & Wall Types Starts/////
                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP3_Wnd_Wall_Type = workbook.Sheets["AP3_Window & Wall Types"];

                Excel.Range mergeRange_WNW1 = worksheet_AP3_Wnd_Wall_Type.Range["A2", "J2"];
                mergeRange_WNW1.Merge();
                mergeRange_WNW1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_WNW1.Font.Name = "Times New Roman";
                mergeRange_WNW1.Value = "WINDOW / WALL TYPES";
                mergeRange_WNW1.Font.Bold = true;

                //////////// for Opaque Wall Types

                Excel.Range mergeRange_WNW2 = worksheet_AP3_Wnd_Wall_Type.Range["A4", "J4"];
                mergeRange_WNW2.Merge();
                mergeRange_WNW2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_WNW2.Font.Name = "Times New Roman";
                mergeRange_WNW2.Value = "OPAQUE WALL TYPES";
                mergeRange_WNW2.Font.Bold = true;

                Excel.Range mergeRange_WNW3 = worksheet_AP3_Wnd_Wall_Type.Cells[5, 1];
                mergeRange_WNW3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW3.Value = "S/No";
                mergeRange_WNW3.Font.Name = "Times New Roman";
                mergeRange_WNW3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW4 = worksheet_AP3_Wnd_Wall_Type.Range["B5", "F5"];
                mergeRange_WNW4.Merge();
                mergeRange_WNW4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW4.Value = "Brief Description";
                mergeRange_WNW4.Font.Name = "Times New Roman";
                mergeRange_WNW4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW5 = worksheet_AP3_Wnd_Wall_Type.Range["G5", "H5"];
                mergeRange_WNW5.Merge();
                mergeRange_WNW5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW5.Value = "U Values (W/m2K)";
                mergeRange_WNW5.Font.Name = "Times New Roman";
                mergeRange_WNW5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW6 = worksheet_AP3_Wnd_Wall_Type.Range["I5", "J5"];
                mergeRange_WNW6.Merge();
                mergeRange_WNW6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW6.Value = "Thickness (mm)";
                mergeRange_WNW6.Font.Name = "Times New Roman";
                mergeRange_WNW6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int rowAdder3 = 5;
                //////for Datagrid data to excel for opaque wall types
                for (int i = 0; i < dataGrid_Wall_Types.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wall_Types.Columns.Count; j++)
                    {
                        if (j == 0) // j=1 prints S/No
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 1] = dataGrid_Wall_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;                            
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallTypes1 = worksheet_AP3_Wnd_Wall_Type.Range[worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, 2], worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, 6]];
                            mergeRangeWallTypes1.Merge();
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3+ 1 + i, j + 1] = dataGrid_Wall_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2, prints U Values(W/m2K)
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallTypes2 = worksheet_AP3_Wnd_Wall_Type.Range[worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, 7], worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, 8]];
                            mergeRangeWallTypes2.Merge();
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 5] = dataGrid_Wall_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 5].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=3, prints Thickness(mm)
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallTypes3 = worksheet_AP3_Wnd_Wall_Type.Range[worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, 9], worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, 10]];
                            mergeRangeWallTypes3.Merge();
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 6] = dataGrid_Wall_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 6].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[rowAdder3 + 1 + i, j + 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }


                    }
                }


                //////////// for Window Types

                int WdnTypeSep = dataGrid_Wall_Types.Rows.Count + rowAdder3 + 1;

                // Set the column width for Brief Description
                worksheet_AP3_Wnd_Wall_Type.Columns[3].ColumnWidth = 30;
                worksheet_AP3_Wnd_Wall_Type.Columns[4].ColumnWidth = 15;
                worksheet_AP3_Wnd_Wall_Type.Columns[6].ColumnWidth = 20;

                Excel.Range mergeRange_WNW7 = worksheet_AP3_Wnd_Wall_Type.Range[worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep, 1], worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep, 10]];
                mergeRange_WNW7.Merge();
                mergeRange_WNW7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_WNW7.Font.Name = "Times New Roman";
                mergeRange_WNW7.Value = "FENESTRATION TYPES";
                mergeRange_WNW7.Font.Bold = true;

                Excel.Range mergeRange_WNW8 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 1];
                mergeRange_WNW8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW8.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW8.Value = "S/No";
                mergeRange_WNW8.Font.Name = "Times New Roman";
                mergeRange_WNW8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW9 = worksheet_AP3_Wnd_Wall_Type.Range[worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep + 1, 2], worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep + 1, 3]];
                mergeRange_WNW9.Merge();
                mergeRange_WNW9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW9.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW9.Value = "Brief Description";
                mergeRange_WNW9.Font.Name = "Times New Roman";
                mergeRange_WNW9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW10 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 4];
                mergeRange_WNW10.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW10.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW10.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW10.Value = "U Values(W/m2K)";
                mergeRange_WNW10.Font.Name = "Times New Roman";
                mergeRange_WNW10.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW11 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 5];
                mergeRange_WNW11.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW11.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW11.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW11.Value = "SC1";
                mergeRange_WNW11.Font.Name = "Times New Roman";
                mergeRange_WNW11.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW12 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 6];
                mergeRange_WNW12.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW12.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW12.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW12.Value = "Shading Type";
                mergeRange_WNW12.Font.Name = "Times New Roman";
                mergeRange_WNW12.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW13 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 7];
                mergeRange_WNW13.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW13.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW13.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW13.Value = "Angle(deg)";
                mergeRange_WNW13.Font.Name = "Times New Roman";
                mergeRange_WNW13.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW14 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 8];
                mergeRange_WNW14.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW14.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW14.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW14.Value = "P(m)";
                mergeRange_WNW14.Font.Name = "Times New Roman";
                mergeRange_WNW14.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW15 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 9];
                mergeRange_WNW15.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW15.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW15.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW15.Value = "H(m)";
                mergeRange_WNW15.Font.Name = "Times New Roman";
                mergeRange_WNW15.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WNW16 = worksheet_AP3_Wnd_Wall_Type.Cells[(WdnTypeSep + 1), 10];
                mergeRange_WNW16.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WNW16.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WNW16.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WNW16.Value = "W(m)";
                mergeRange_WNW16.Font.Name = "Times New Roman";
                mergeRange_WNW16.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int WdnTypeSep1 = WdnTypeSep + 2;


                for (int i = 0; i < dataGrid_Wndw_Types.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGrid_Wndw_Types.Columns.Count; j++)
                    {

                        if (j == 0) // j=0 prints S/No
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 1] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 1) // j=1, prints BriefDesc
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWndwTypes1 = worksheet_AP3_Wnd_Wall_Type.Range[worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, 2], worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, 3]];
                            mergeRangeWndwTypes1.Merge();
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 1] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }

                        if (j == 2) // j=2 prints U Values
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 3) // j=2 prints SC1
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 4) // j=4 prints Shading Type
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 5) // j=5 prints Angel
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 6) // j=6 prints P(m)
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 7) // j=7 prints H(m)
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        if (j == 8) // j=8 prints W(m)
                        {
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2] = dataGrid_Wndw_Types.Rows[i].Cells[j].Value;
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_Wnd_Wall_Type.Cells[WdnTypeSep1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }


                    }
                }

                ///// For AP3_Window & Wall Types Ends/////
                ///


                ///// For AP3_Window Assembly Starts/////
                Microsoft.Office.Interop.Excel.Worksheet worksheet_AP3_WallAss = workbook.Sheets["AP3_Wall Assembly"];

                // Set the column width for Brief Description
                worksheet_AP3_WallAss.Columns[3].ColumnWidth = 13;
                worksheet_AP3_WallAss.Columns[4].ColumnWidth = 13;
                worksheet_AP3_WallAss.Columns[7].ColumnWidth = 13;
                worksheet_AP3_WallAss.Columns[8].ColumnWidth = 13;
                worksheet_AP3_WallAss.Columns[11].ColumnWidth = 13;
                worksheet_AP3_WallAss.Columns[12].ColumnWidth = 13;

                Excel.Range mergeRange_WASS0 = worksheet_AP3_WallAss.Range["A2", "L2"];
                mergeRange_WASS0.Merge();
                mergeRange_WASS0.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                mergeRange_WASS0.Font.Name = "Times New Roman";
                mergeRange_WASS0.Value = "WALL ASSEMBLY";
                mergeRange_WASS0.Font.Bold = true;


                Excel.Range mergeRange_WASS1 = worksheet_AP3_WallAss.Range["A3", "B3"];
                mergeRange_WASS1.Merge();
                mergeRange_WASS1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WASS1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WASS1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WASS1.Value = "Function";
                mergeRange_WASS1.Font.Name = "Times New Roman";
                mergeRange_WASS1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WASS2 = worksheet_AP3_WallAss.Range["C3", "D3"];
                mergeRange_WASS2.Merge();
                mergeRange_WASS2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WASS2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WASS2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WASS2.Value = "Material";
                mergeRange_WASS2.Font.Name = "Times New Roman";
                mergeRange_WASS2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                Excel.Range mergeRange_WASS3 = worksheet_AP3_WallAss.Range["E3", "F3"];
                mergeRange_WASS3.Merge();
                mergeRange_WASS3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WASS3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WASS3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WASS3.Value = "Thickness(mm)";
                mergeRange_WASS3.Font.Name = "Times New Roman";
                mergeRange_WASS3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WASS4 = worksheet_AP3_WallAss.Range["G3", "H3"];
                mergeRange_WASS4.Merge();
                mergeRange_WASS4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WASS4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WASS4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WASS4.Value = "Thermal Conductivity (W/mK)";
                mergeRange_WASS4.Font.Name = "Times New Roman";
                mergeRange_WASS4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WASS5 = worksheet_AP3_WallAss.Range["I3", "J3"];
                mergeRange_WASS5.Merge();
                mergeRange_WASS5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WASS5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WASS5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WASS5.Value = "Denisity(kg/m3)";
                mergeRange_WASS5.Font.Name = "Times New Roman";
                mergeRange_WASS5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mergeRange_WASS6 = worksheet_AP3_WallAss.Range["K3", "L3"];
                mergeRange_WASS6.Merge();
                mergeRange_WASS6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                mergeRange_WASS6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                mergeRange_WASS6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                mergeRange_WASS6.Value = "Thermal Resistance R(m2K/W)";
                mergeRange_WASS6.Font.Name = "Times New Roman";
                mergeRange_WASS6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int rowAdder4 = 3;
                //////for Datagrid data to excel for opaque wall types
                for (int i = 0; i < dataGridView5.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView5.Columns.Count; j++)
                    {
                        if (j == 0) // j=0 prints Function
                        {
                            string cellValue = dataGridView5.Rows[i].Cells[j].Value?.ToString();

                            switch (cellValue)
                            {
                                case "5":
                                    cellValue = "finish1";
                                    break;
                                case "4":
                                    cellValue = "insulation";
                                    break;
                                case "3":
                                    cellValue = "membrane";
                                    break;
                                case "2":
                                    cellValue = "substrate";
                                    break;
                                case "1":
                                    cellValue = "structure";
                                    break;
                                case "100":
                                    cellValue = "finish2";
                                    break;
                            }

                            Microsoft.Office.Interop.Excel.Range mergeRangeWallAss1 = worksheet_AP3_WallAss.Range[worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 1], worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 2]];
                            mergeRangeWallAss1.Merge();
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 1] = cellValue;
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 1].Font.Name = "Times New Roman";
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            mergeRangeWallAss1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            mergeRangeWallAss1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            mergeRangeWallAss1.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }

                        if (j == 1) // j=1 prints Material
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallAss2 = worksheet_AP3_WallAss.Range[worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 3], worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 4]];
                            mergeRangeWallAss2.Merge();
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 2] = dataGridView5.Rows[i].Cells[j].Value;
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 2].Font.Name = "Times New Roman";
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            mergeRangeWallAss2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            mergeRangeWallAss2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            mergeRangeWallAss2.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }

                        if (j == 2) // j=2 prints Thickness
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallAss3 = worksheet_AP3_WallAss.Range[worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 5], worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 6]];
                            mergeRangeWallAss3.Merge();
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 3] = dataGridView5.Rows[i].Cells[j].Value;
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 3].Font.Name = "Times New Roman";
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            mergeRangeWallAss3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            mergeRangeWallAss3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            mergeRangeWallAss3.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }

                        if (j == 3) // j=3 prints Thermal Conductivity
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallAss4 = worksheet_AP3_WallAss.Range[worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 7], worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 8]];
                            mergeRangeWallAss4.Merge();
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 4] = dataGridView5.Rows[i].Cells[j].Value;
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 4].Font.Name = "Times New Roman";
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            mergeRangeWallAss4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            mergeRangeWallAss4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            mergeRangeWallAss4.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }

                        if (j == 4) // j=4 prints Denisty
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallAss5 = worksheet_AP3_WallAss.Range[worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 9], worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 10]];
                            mergeRangeWallAss5.Merge();
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 5] = dataGridView5.Rows[i].Cells[j].Value;
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 5].Font.Name = "Times New Roman";
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            mergeRangeWallAss5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            mergeRangeWallAss5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            mergeRangeWallAss5.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }

                        if (j == 5) // j=5 prints Thermal Resistance
                        {
                            Microsoft.Office.Interop.Excel.Range mergeRangeWallAss6 = worksheet_AP3_WallAss.Range[worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 11], worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, 12]];
                            mergeRangeWallAss6.Merge();
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 6] = dataGridView5.Rows[i].Cells[j].Value;
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 6].Font.Name = "Times New Roman";
                            worksheet_AP3_WallAss.Cells[rowAdder4 + 1 + i, j + 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            mergeRangeWallAss6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            mergeRangeWallAss6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                            mergeRangeWallAss6.Borders.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;
                        }

                    }
                }


                ///////////////////////////////////////////////////////////// for AP3 Worksheets - Ends ////////////////////////////////////////////////////////////////

                ///////////////////////////////////////////////////////////// for AP5 Worksheets - Starts ////////////////////////////////////////////////////////////////
                AddWorksheetsForAP5Tabs();
                int num_fc = 0;
                //int num_fc_hp = 0;
                foreach (string st in WndwTypeLst1)
                {
                   
                    Microsoft.Office.Interop.Excel.Worksheet worksheet_AP5= workbook.Sheets["AP5_"+st];

                    Excel.Range mergeRange_AP5 = worksheet_AP5.Range["A2", "M2"];
                    mergeRange_AP5.Merge();
                    mergeRange_AP5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    mergeRange_AP5.Font.Name = "Times New Roman";
                    mergeRange_AP5.Value = "SC CALCULATION";
                    mergeRange_AP5.Font.Bold = true;

                    Excel.Range mergeRange_AP5_1 = worksheet_AP5.Range["A3", "M3"];
                    mergeRange_AP5_1.Merge();
                    mergeRange_AP5_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    mergeRange_AP5_1.Font.Name = "Times New Roman";
                    mergeRange_AP5_1.Value = "FENESTRATION: "+st;
                    mergeRange_AP5_1.Font.Bold = true;

                    Excel.Range mergeRange_AP5_2 = worksheet_AP5.Cells[5,1];                    
                    mergeRange_AP5_2.Value = "SC1:";
                    mergeRange_AP5_2.Font.Name = "Times New Roman";
                    mergeRange_AP5_2.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ///////////////////////////////////////////////////
                    Excel.Range mergeRange_AP5_3 = worksheet_AP5.Cells[5, 2];
                    mergeRange_AP5_3.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString();
                    mergeRange_AP5_3.Font.Name = "Times New Roman";
                    mergeRange_AP5_3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    mergeRange_AP5_3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    mergeRange_AP5_3.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                    mergeRange_AP5_3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                    Excel.Range mergeRange_AP5_4 = worksheet_AP5.Cells[7, 1];
                    mergeRange_AP5_4.Value = "U Value:";
                    mergeRange_AP5_4.Font.Name = "Times New Roman";
                    mergeRange_AP5_4.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ///////////////////////////////////////////////////
                    Excel.Range mergeRange_AP5_5 = worksheet_AP5.Cells[7, 2];
                    mergeRange_AP5_5.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[2].Value.ToString();
                    mergeRange_AP5_5.Font.Name = "Times New Roman";
                    mergeRange_AP5_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    mergeRange_AP5_5.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    mergeRange_AP5_5.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                    mergeRange_AP5_5.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                    ///////////////////////////////////////////////////
                    Excel.Range mergeRange_AP5_6 = worksheet_AP5.Cells[7, 3];
                    mergeRange_AP5_6.Value = "W/m2K";
                    mergeRange_AP5_6.Font.Name = "Times New Roman";
                    mergeRange_AP5_6.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


                    Excel.Range mergeRange_AP5_7 = worksheet_AP5.Cells[9, 1];
                    mergeRange_AP5_7.Value = "Shades:";
                    mergeRange_AP5_7.Font.Name = "Times New Roman";
                    mergeRange_AP5_7.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ///////////////////////////////////////////////////
                    Excel.Range mergeRange_AP5_8 = worksheet_AP5.Range["B9", "C9"];
                    mergeRange_AP5_8.Merge();
                    mergeRange_AP5_8.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[4].Value.ToString();
                    mergeRange_AP5_8.Font.Name = "Times New Roman";
                    mergeRange_AP5_8.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    mergeRange_AP5_8.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    mergeRange_AP5_8.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                    mergeRange_AP5_8.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                    if (dataGrid_Wndw_Types.Rows[num_fc].Cells[4].Value.ToString() == "Horizontal Projection")
                    {
                        Excel.Range mergeRange_AP5_9 = worksheet_AP5.Range["E5", "G18"];
                        mergeRange_AP5_9.Merge();
                        string imagePath1 = BuildingCoder.Util.GetFilePath("Hor_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Hor_Pro_Wndw.png"; // Specify the path to your PNG image

                        if (File.Exists(imagePath1))
                        {
                            // Calculate the position and size of the image to fit within the merge cell
                            float left = (float)mergeRange_AP5_9.Left;
                            float top = (float)mergeRange_AP5_9.Top;
                            float width = (float)mergeRange_AP5_9.Width;
                            float height = (float)mergeRange_AP5_9.Height;
                            // Add the image to the worksheet
                            worksheet_AP5.Shapes.AddPicture(imagePath1, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
                        }
                        else
                        {
                            MessageBox.Show("Image file not found at the specified path.");
                        }

                        Excel.Range mergeRange_AP5_10 = worksheet_AP5.Cells[5, 9];
                        mergeRange_AP5_10.Value = "Angle (deg):";
                        mergeRange_AP5_10.Font.Name = "Times New Roman";
                        mergeRange_AP5_10.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_11 = worksheet_AP5.Cells[5, 10];
                        mergeRange_AP5_11.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString();
                        mergeRange_AP5_11.Font.Name = "Times New Roman";
                        mergeRange_AP5_11.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_11.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_12 = worksheet_AP5.Cells[7, 9];
                        mergeRange_AP5_12.Value = "P (m):";
                        mergeRange_AP5_12.Font.Name = "Times New Roman";
                        mergeRange_AP5_12.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_13 = worksheet_AP5.Cells[7, 10];
                        mergeRange_AP5_13.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString();
                        mergeRange_AP5_13.Font.Name = "Times New Roman";
                        mergeRange_AP5_13.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_13.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_14 = worksheet_AP5.Cells[9, 9];
                        mergeRange_AP5_14.Value = "H (m):";
                        mergeRange_AP5_14.Font.Name = "Times New Roman";
                        mergeRange_AP5_14.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_15 = worksheet_AP5.Cells[9, 10];
                        mergeRange_AP5_15.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString();
                        mergeRange_AP5_15.Font.Name = "Times New Roman";
                        mergeRange_AP5_15.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_15.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_16 = worksheet_AP5.Cells[11, 9];
                        mergeRange_AP5_16.Value = "W (m):";
                        mergeRange_AP5_16.Font.Name = "Times New Roman";
                        mergeRange_AP5_16.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_17 = worksheet_AP5.Cells[11, 10];
                        mergeRange_AP5_17.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString();
                        mergeRange_AP5_17.Font.Name = "Times New Roman";
                        mergeRange_AP5_17.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_17.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_18 = worksheet_AP5.Cells[13, 9];
                        mergeRange_AP5_18.Value = "R1 = P/H:";
                        mergeRange_AP5_18.Font.Name = "Times New Roman";
                        mergeRange_AP5_18.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_19 = worksheet_AP5.Cells[13, 10];
                        double a = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString());
                        double b = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString());
                        double c = Math.Round((a / b),4);
                        mergeRange_AP5_19.Value = c.ToString();
                        mergeRange_AP5_19.Font.Name = "Times New Roman";
                        mergeRange_AP5_19.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_19.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22 = worksheet_AP5.Range["A20", "C20"];
                        mergeRange_AP5_22.Merge();
                        mergeRange_AP5_22.Value = "North-South";
                        mergeRange_AP5_22.Font.Name = "Times New Roman";
                        mergeRange_AP5_22.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23 = worksheet_AP5.Cells[20, 4];
                        mergeRange_AP5_23.Value = "SC2:";
                        mergeRange_AP5_23.Font.Name = "Times New Roman";
                        mergeRange_AP5_23.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24 = worksheet_AP5.Cells[20, 5];
                        double Agl = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_NS = Interpolation_for_SC2(Agl, c, dataGrid_SC2_NS);
                        if (SC2_NS == 0)
                        {
                            SC2_NS = 1;
                        }
                        mergeRange_AP5_24.Value = SC2_NS.ToString("0.0000");
                        mergeRange_AP5_24.Font.Name = "Times New Roman";
                        mergeRange_AP5_24.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25 = worksheet_AP5.Range["G20", "H20"];
                        mergeRange_AP5_25.Merge();
                        mergeRange_AP5_25.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25.Font.Name = "Times New Roman";
                        mergeRange_AP5_25.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26 = worksheet_AP5.Cells[20, 9];
                        mergeRange_AP5_26.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_NS).ToString("0.0000");
                        mergeRange_AP5_26.Font.Name = "Times New Roman";
                        mergeRange_AP5_26.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        
                        Excel.Range mergeRange_AP5_22a = worksheet_AP5.Range["A22", "C22"];
                        mergeRange_AP5_22a.Merge();
                        mergeRange_AP5_22a.Value = "East-West";
                        mergeRange_AP5_22a.Font.Name = "Times New Roman";
                        mergeRange_AP5_22a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23a = worksheet_AP5.Cells[22, 4];
                        mergeRange_AP5_23a.Value = "SC2:";
                        mergeRange_AP5_23a.Font.Name = "Times New Roman";
                        mergeRange_AP5_23a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24a = worksheet_AP5.Cells[22, 5];
                        double Agl1 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_EW = Interpolation_for_SC2(Agl1, c, dataGrid_SC2_EW);
                        if (SC2_EW == 0)
                        {
                            SC2_EW = 1;
                        }
                        mergeRange_AP5_24a.Value = SC2_EW.ToString("0.0000");
                        mergeRange_AP5_24a.Font.Name = "Times New Roman";
                        mergeRange_AP5_24a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25a = worksheet_AP5.Range["G22", "H22"];
                        mergeRange_AP5_25a.Merge();
                        mergeRange_AP5_25a.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25a.Font.Name = "Times New Roman";
                        mergeRange_AP5_25a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26a = worksheet_AP5.Cells[22, 9];
                        mergeRange_AP5_26a.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_EW).ToString("0.0000");
                        mergeRange_AP5_26a.Font.Name = "Times New Roman";
                        mergeRange_AP5_26a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22b = worksheet_AP5.Range["A24", "C24"];
                        mergeRange_AP5_22b.Merge();
                        mergeRange_AP5_22b.Value = "NorthEast-NorthWest";
                        mergeRange_AP5_22b.Font.Name = "Times New Roman";
                        mergeRange_AP5_22b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23b = worksheet_AP5.Cells[24, 4];
                        mergeRange_AP5_23b.Value = "SC2:";
                        mergeRange_AP5_23b.Font.Name = "Times New Roman";
                        mergeRange_AP5_23b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24b = worksheet_AP5.Cells[24, 5];
                        double Agl2 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_NENW = Interpolation_for_SC2(Agl2, c, dataGrid_SC2_NENW);
                        if (SC2_NENW == 0)
                        {
                            SC2_NENW = 1;
                        }
                        mergeRange_AP5_24b.Value = SC2_NENW.ToString("0.0000");
                        mergeRange_AP5_24b.Font.Name = "Times New Roman";
                        mergeRange_AP5_24b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25b = worksheet_AP5.Range["G24", "H24"];
                        mergeRange_AP5_25b.Merge();
                        mergeRange_AP5_25b.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25b.Font.Name = "Times New Roman";
                        mergeRange_AP5_25b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26b = worksheet_AP5.Cells[24, 9];
                        mergeRange_AP5_26b.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_NENW).ToString("0.0000");
                        mergeRange_AP5_26b.Font.Name = "Times New Roman";
                        mergeRange_AP5_26b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;



                        Excel.Range mergeRange_AP5_22c = worksheet_AP5.Range["A26", "C26"];
                        mergeRange_AP5_22c.Merge();
                        mergeRange_AP5_22c.Value = "SouthEast-SouthWest";
                        mergeRange_AP5_22c.Font.Name = "Times New Roman";
                        mergeRange_AP5_22c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23c = worksheet_AP5.Cells[26, 4];
                        mergeRange_AP5_23c.Value = "SC2:";
                        mergeRange_AP5_23c.Font.Name = "Times New Roman";
                        mergeRange_AP5_23c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24c = worksheet_AP5.Cells[26, 5];
                        double Agl3 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_SESW = Interpolation_for_SC2(Agl3, c, dataGrid_SC2_SESW);
                        if (SC2_SESW == 0)
                        {
                            SC2_SESW = 1;
                        }
                        mergeRange_AP5_24c.Value = SC2_SESW.ToString("0.0000");
                        mergeRange_AP5_24c.Font.Name = "Times New Roman";
                        mergeRange_AP5_24c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25c = worksheet_AP5.Range["G26", "H26"];
                        mergeRange_AP5_25c.Merge();
                        mergeRange_AP5_25c.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25c.Font.Name = "Times New Roman";
                        mergeRange_AP5_25c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26c = worksheet_AP5.Cells[26, 9];
                        mergeRange_AP5_26c.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_SESW).ToString("0.0000");
                        mergeRange_AP5_26c.Font.Name = "Times New Roman";
                        mergeRange_AP5_26c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;



                        //TB_SC2_NS.Text = SC2_NS.ToString("0.0000");
                        //TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");

                        //mergeRange_AP5_24.Value = Lst_SC2_HP_NS[num_fc_hp].ToString("0.0000");
                        //mergeRange_AP5_24.Font.Name = "Times New Roman";
                        //mergeRange_AP5_24.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        //mergeRange_AP5_24.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        //mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        //mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        //Excel.Range mergeRange_AP5_25 = worksheet_AP5.Range["H20", "I20"];
                        //mergeRange_AP5_25.Merge();
                        //mergeRange_AP5_25.Value = "SC = SC1 x SC2";
                        //mergeRange_AP5_25.Font.Name = "Times New Roman";
                        ///mergeRange_AP5_25.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        //TB_SC2_NS.Text = SC2_NS.ToString("0.0000");
                        //TB_SC_NS.Text = (double.Parse(TB1.Text) * double.Parse(TB_SC2_NS.Text)).ToString("0.0000");
                        //Lst_SC2_HP_NS.Add(SC2_NS);

                        //(double.Parse(mergeRange_AP5_13.Value) / double.Parse(mergeRange_AP5_14.Value))

                        //LB11.Text = "R1 = P/H:";
                        //double R1 = Math.Round((double.Parse(TB8.Text) / double.Parse(TB9.Text)), 4);
                        //TB11.Text = R1.ToString();

                        // Set the column width for Brief Description
                        worksheet_AP5.Columns[9].ColumnWidth = 10;
                        //num_fc_hp = num_fc_hp + 1;

                    }

                    else if (dataGrid_Wndw_Types.Rows[num_fc].Cells[4].Value.ToString() == "Vertical Projection")
                    {
                        Excel.Range mergeRange_AP5_9 = worksheet_AP5.Range["E5", "I13"];
                        mergeRange_AP5_9.Merge();
                        string imagePath1 = BuildingCoder.Util.GetFilePath("Ver_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Ver_Pro_Wndw.png"; // Specify the path to your PNG image

                        if (File.Exists(imagePath1))
                        {
                            // Calculate the position and size of the image to fit within the merge cell
                            float left = (float)mergeRange_AP5_9.Left;
                            float top = (float)mergeRange_AP5_9.Top;
                            float width = (float)mergeRange_AP5_9.Width;
                            float height = (float)mergeRange_AP5_9.Height;
                            // Add the image to the worksheet
                            worksheet_AP5.Shapes.AddPicture(imagePath1, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
                        }
                        else
                        {
                            MessageBox.Show("Image file not found at the specified path.");
                        }

                        Excel.Range mergeRange_AP5_10 = worksheet_AP5.Cells[5, 11];
                        mergeRange_AP5_10.Value = "Angle (deg):";
                        mergeRange_AP5_10.Font.Name = "Times New Roman";
                        mergeRange_AP5_10.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_11 = worksheet_AP5.Cells[5, 12];
                        mergeRange_AP5_11.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString();
                        mergeRange_AP5_11.Font.Name = "Times New Roman";
                        mergeRange_AP5_11.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_11.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_12 = worksheet_AP5.Cells[7, 11];
                        mergeRange_AP5_12.Value = "P (m):";
                        mergeRange_AP5_12.Font.Name = "Times New Roman";
                        mergeRange_AP5_12.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_13 = worksheet_AP5.Cells[7, 12];
                        mergeRange_AP5_13.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString();
                        mergeRange_AP5_13.Font.Name = "Times New Roman";
                        mergeRange_AP5_13.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_13.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_14 = worksheet_AP5.Cells[9, 11];
                        mergeRange_AP5_14.Value = "H (m):";
                        mergeRange_AP5_14.Font.Name = "Times New Roman";
                        mergeRange_AP5_14.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_15 = worksheet_AP5.Cells[9, 12];
                        mergeRange_AP5_15.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString();
                        mergeRange_AP5_15.Font.Name = "Times New Roman";
                        mergeRange_AP5_15.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_15.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_16 = worksheet_AP5.Cells[11, 11];
                        mergeRange_AP5_16.Value = "W (m):";
                        mergeRange_AP5_16.Font.Name = "Times New Roman";
                        mergeRange_AP5_16.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_17 = worksheet_AP5.Cells[11, 12];
                        mergeRange_AP5_17.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString();
                        mergeRange_AP5_17.Font.Name = "Times New Roman";
                        mergeRange_AP5_17.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_17.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_18 = worksheet_AP5.Cells[13, 11];
                        mergeRange_AP5_18.Value = "R2 = P/W:";
                        mergeRange_AP5_18.Font.Name = "Times New Roman";
                        mergeRange_AP5_18.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_19 = worksheet_AP5.Cells[13, 12];
                        double a = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString());
                        double b = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString());
                        double c = Math.Round((a / b), 4);
                        mergeRange_AP5_19.Value = c.ToString();
                        mergeRange_AP5_19.Font.Name = "Times New Roman";
                        mergeRange_AP5_19.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_19.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_22 = worksheet_AP5.Range["A20", "C20"];
                        mergeRange_AP5_22.Merge();
                        mergeRange_AP5_22.Value = "North-South";
                        mergeRange_AP5_22.Font.Name = "Times New Roman";
                        mergeRange_AP5_22.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23 = worksheet_AP5.Cells[20, 4];
                        mergeRange_AP5_23.Value = "SC2:";
                        mergeRange_AP5_23.Font.Name = "Times New Roman";
                        mergeRange_AP5_23.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24 = worksheet_AP5.Cells[20, 5];
                        double Agl = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_VP_NS = Interpolation_for_SC2(Agl, c, dataGrid_SC2_VP_NS);
                        if (SC2_VP_NS == 0)
                        {
                            SC2_VP_NS = 1;
                        }
                        mergeRange_AP5_24.Value = SC2_VP_NS.ToString("0.0000");
                        mergeRange_AP5_24.Font.Name = "Times New Roman";
                        mergeRange_AP5_24.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25 = worksheet_AP5.Range["G20", "H20"];
                        mergeRange_AP5_25.Merge();
                        mergeRange_AP5_25.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25.Font.Name = "Times New Roman";
                        mergeRange_AP5_25.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26 = worksheet_AP5.Cells[20, 9];
                        mergeRange_AP5_26.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_VP_NS).ToString("0.0000");
                        mergeRange_AP5_26.Font.Name = "Times New Roman";
                        mergeRange_AP5_26.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22a = worksheet_AP5.Range["A22", "C22"];
                        mergeRange_AP5_22a.Merge();
                        mergeRange_AP5_22a.Value = "East-West";
                        mergeRange_AP5_22a.Font.Name = "Times New Roman";
                        mergeRange_AP5_22a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23a = worksheet_AP5.Cells[22, 4];
                        mergeRange_AP5_23a.Value = "SC2:";
                        mergeRange_AP5_23a.Font.Name = "Times New Roman";
                        mergeRange_AP5_23a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24a = worksheet_AP5.Cells[22, 5];
                        double Agl1 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_VP_EW = Interpolation_for_SC2(Agl1, c, dataGrid_SC2_VP_EW);
                        if (SC2_VP_EW == 0)
                        {
                            SC2_VP_EW = 1;
                        }
                        mergeRange_AP5_24a.Value = SC2_VP_EW.ToString("0.0000");
                        mergeRange_AP5_24a.Font.Name = "Times New Roman";
                        mergeRange_AP5_24a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25a = worksheet_AP5.Range["G22", "H22"];
                        mergeRange_AP5_25a.Merge();
                        mergeRange_AP5_25a.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25a.Font.Name = "Times New Roman";
                        mergeRange_AP5_25a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26a = worksheet_AP5.Cells[22, 9];
                        mergeRange_AP5_26a.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_VP_EW).ToString("0.0000");
                        mergeRange_AP5_26a.Font.Name = "Times New Roman";
                        mergeRange_AP5_26a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22b = worksheet_AP5.Range["A24", "C24"];
                        mergeRange_AP5_22b.Merge();
                        mergeRange_AP5_22b.Value = "NorthEast-NorthWest";
                        mergeRange_AP5_22b.Font.Name = "Times New Roman";
                        mergeRange_AP5_22b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23b = worksheet_AP5.Cells[24, 4];
                        mergeRange_AP5_23b.Value = "SC2:";
                        mergeRange_AP5_23b.Font.Name = "Times New Roman";
                        mergeRange_AP5_23b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24b = worksheet_AP5.Cells[24, 5];
                        double Agl2 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_VP_NENW = Interpolation_for_SC2(Agl2, c, dataGrid_SC2_VP_NENW);
                        if (SC2_VP_NENW == 0)
                        {
                            SC2_VP_NENW = 1;
                        }
                        mergeRange_AP5_24b.Value = SC2_VP_NENW.ToString("0.0000");
                        mergeRange_AP5_24b.Font.Name = "Times New Roman";
                        mergeRange_AP5_24b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25b = worksheet_AP5.Range["G24", "H24"];
                        mergeRange_AP5_25b.Merge();
                        mergeRange_AP5_25b.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25b.Font.Name = "Times New Roman";
                        mergeRange_AP5_25b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26b = worksheet_AP5.Cells[24, 9];
                        mergeRange_AP5_26b.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_VP_NENW).ToString("0.0000");
                        mergeRange_AP5_26b.Font.Name = "Times New Roman";
                        mergeRange_AP5_26b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22c = worksheet_AP5.Range["A26", "C26"];
                        mergeRange_AP5_22c.Merge();
                        mergeRange_AP5_22c.Value = "SouthEast-SouthWest";
                        mergeRange_AP5_22c.Font.Name = "Times New Roman";
                        mergeRange_AP5_22c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23c = worksheet_AP5.Cells[26, 4];
                        mergeRange_AP5_23c.Value = "SC2:";
                        mergeRange_AP5_23c.Font.Name = "Times New Roman";
                        mergeRange_AP5_23c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24c = worksheet_AP5.Cells[26, 5];
                        double Agl3 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());
                        SC2_VP_SESW = Interpolation_for_SC2(Agl3, c, dataGrid_SC2_VP_SESW);
                        if (SC2_VP_SESW == 0)
                        {
                            SC2_VP_SESW = 1;
                        }
                        mergeRange_AP5_24c.Value = SC2_VP_SESW.ToString("0.0000");
                        mergeRange_AP5_24c.Font.Name = "Times New Roman";
                        mergeRange_AP5_24c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25c = worksheet_AP5.Range["G26", "H26"];
                        mergeRange_AP5_25c.Merge();
                        mergeRange_AP5_25c.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25c.Font.Name = "Times New Roman";
                        mergeRange_AP5_25c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26c = worksheet_AP5.Cells[26, 9];
                        mergeRange_AP5_26c.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_VP_SESW).ToString("0.0000");
                        mergeRange_AP5_26c.Font.Name = "Times New Roman";
                        mergeRange_AP5_26c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        // Set the column width for Brief Description
                        worksheet_AP5.Columns[11].ColumnWidth = 10;
                    }

                    else if (dataGrid_Wndw_Types.Rows[num_fc].Cells[4].Value.ToString() == "Egg Crate Window")
                    {
                        Excel.Range mergeRange_AP5_9 = worksheet_AP5.Range["E5", "I19"];
                        mergeRange_AP5_9.Merge();
                        string imagePath1 = BuildingCoder.Util.GetFilePath("Egg_Crt_Wndw.png");  //"D:\WIP\01_API\ETTV\Egg_Crt_Wndw.png"; // Specify the path to your PNG image

                        if (File.Exists(imagePath1))
                        {
                            // Calculate the position and size of the image to fit within the merge cell
                            float left = (float)mergeRange_AP5_9.Left;
                            float top = (float)mergeRange_AP5_9.Top;
                            float width = (float)mergeRange_AP5_9.Width;
                            float height = (float)mergeRange_AP5_9.Height;
                            // Add the image to the worksheet
                            worksheet_AP5.Shapes.AddPicture(imagePath1, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
                        }
                        else
                        {
                            MessageBox.Show("Image file not found at the specified path.");
                        }

                        Excel.Range mergeRange_AP5_10 = worksheet_AP5.Cells[5, 11];
                        mergeRange_AP5_10.Value = "Angle (deg):";
                        mergeRange_AP5_10.Font.Name = "Times New Roman";
                        mergeRange_AP5_10.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_11 = worksheet_AP5.Cells[5, 12];
                        mergeRange_AP5_11.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString();
                        mergeRange_AP5_11.Font.Name = "Times New Roman";
                        mergeRange_AP5_11.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_11.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_12 = worksheet_AP5.Cells[7, 11];
                        mergeRange_AP5_12.Value = "P (m):";
                        mergeRange_AP5_12.Font.Name = "Times New Roman";
                        mergeRange_AP5_12.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_13 = worksheet_AP5.Cells[7, 12];
                        mergeRange_AP5_13.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString();
                        mergeRange_AP5_13.Font.Name = "Times New Roman";
                        mergeRange_AP5_13.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_13.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_14 = worksheet_AP5.Cells[9, 11];
                        mergeRange_AP5_14.Value = "H (m):";
                        mergeRange_AP5_14.Font.Name = "Times New Roman";
                        mergeRange_AP5_14.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_15 = worksheet_AP5.Cells[9, 12];
                        mergeRange_AP5_15.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString();
                        mergeRange_AP5_15.Font.Name = "Times New Roman";
                        mergeRange_AP5_15.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_15.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_16 = worksheet_AP5.Cells[11, 11];
                        mergeRange_AP5_16.Value = "W (m):";
                        mergeRange_AP5_16.Font.Name = "Times New Roman";
                        mergeRange_AP5_16.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_17 = worksheet_AP5.Cells[11, 12];
                        mergeRange_AP5_17.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString();
                        mergeRange_AP5_17.Font.Name = "Times New Roman";
                        mergeRange_AP5_17.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_17.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_18 = worksheet_AP5.Cells[13, 11];
                        mergeRange_AP5_18.Value = "R1 = P/H:";
                        mergeRange_AP5_18.Font.Name = "Times New Roman";
                        mergeRange_AP5_18.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_19 = worksheet_AP5.Cells[13, 12];
                        double a = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString());
                        double b = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString());
                        double c = Math.Round((a / b), 4);
                        mergeRange_AP5_19.Value = c.ToString();
                        mergeRange_AP5_19.Font.Name = "Times New Roman";
                        mergeRange_AP5_19.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_19.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_20 = worksheet_AP5.Cells[15, 11];
                        mergeRange_AP5_20.Value = "R2 = P/W:";
                        mergeRange_AP5_20.Font.Name = "Times New Roman";
                        mergeRange_AP5_20.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_21 = worksheet_AP5.Cells[15, 12];
                        double a1 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString());
                        double b1 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString());
                        double c1 = Math.Round((a1 / b1), 4);
                        mergeRange_AP5_21.Value = c1.ToString();
                        mergeRange_AP5_21.Font.Name = "Times New Roman";
                        mergeRange_AP5_21.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_21.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_21.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_21.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                                                
                        EC_R1 = c;
                        EC_R2 = c1;

                        double Agl = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString());

                        List<double> R1_EC_Lst = new List<double>
                        {
                            0.2,0.4,0.6,0.8,1.0,1.2,1.4,1.6,1.8
                        };

                        foreach (double EC_b in R1_EC_Lst)
                        {
                            if (EC_R1 < 0.2)
                            {
                                EC_b1 = 0.2;
                                EC_b2 = 0.2;
                                break;
                            }
                            if (EC_R1 >= 1.8)
                            {
                                EC_b1 = 1.8;
                                EC_b2 = 1.8;
                                break;
                            }
                            else if (EC_R1 == EC_b)
                            {
                                EC_b1 = EC_b;
                                EC_b2 = EC_b;
                                break;
                            }
                            else if (EC_R1 > 0.2 && EC_b > EC_R1)
                            {
                                EC_b2 = EC_b;
                                EC_b1 = Math.Round((EC_b - 0.2), 1);
                                break;
                            }
                        }

                        ////////////////North-South SC2_EC  
                        ////for b1
                        if (EC_b1 == 0.2)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_0);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.4)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_1);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.6)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_2);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.8)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_3);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.0)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_4);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.2)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_5);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.4)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_6);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.6)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_7);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.8)
                        {
                            SC2_EC_NS_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_8);
                            if (SC2_EC_NS_b1 == 0)
                            {
                                SC2_EC_NS_b1 = 1;
                            }
                        }

                        ////for b2
                        if (EC_b2 == 0.2)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_0);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.4)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_1);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.6)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_2);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.8)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_3);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.0)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_4);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.2)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_5);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.4)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_6);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.6)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_7);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.8)
                        {
                            SC2_EC_NS_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NS_8);
                            if (SC2_EC_NS_b2 == 0)
                            {
                                SC2_EC_NS_b2 = 1;
                            }
                        }

                        //////////////////North-South SC2_EC
                        if (SC2_EC_NS_b1 == SC2_EC_NS_b2)
                        {
                            SC2_EC_NS = SC2_EC_NS_b1;
                        }
                        else
                        {
                            SC2_EC_NS = SC2_EC_NS_b1 + ((SC2_EC_NS_b2 - SC2_EC_NS_b1) * ((EC_R1 - EC_b1) / (EC_b2 - EC_b1)));
                        }



                        Excel.Range mergeRange_AP5_22 = worksheet_AP5.Range["A20", "C20"];
                        mergeRange_AP5_22.Merge();
                        mergeRange_AP5_22.Value = "North-South";
                        mergeRange_AP5_22.Font.Name = "Times New Roman";
                        mergeRange_AP5_22.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23 = worksheet_AP5.Cells[20, 4];
                        mergeRange_AP5_23.Value = "SC2:";
                        mergeRange_AP5_23.Font.Name = "Times New Roman";
                        mergeRange_AP5_23.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24 = worksheet_AP5.Cells[20, 5];                        
                        mergeRange_AP5_24.Value = SC2_EC_NS.ToString("0.0000");
                        mergeRange_AP5_24.Font.Name = "Times New Roman";
                        mergeRange_AP5_24.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25 = worksheet_AP5.Range["G20", "H20"];
                        mergeRange_AP5_25.Merge();
                        mergeRange_AP5_25.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25.Font.Name = "Times New Roman";
                        mergeRange_AP5_25.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26 = worksheet_AP5.Cells[20, 9];
                        mergeRange_AP5_26.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_EC_NS).ToString("0.0000");
                        mergeRange_AP5_26.Font.Name = "Times New Roman";
                        mergeRange_AP5_26.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////

                        ////////////////East-West SC2_EC  
                        ////for b1
                        if (EC_b1 == 0.2)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_0);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.4)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_1);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.6)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_2);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.8)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_3);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.0)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_4);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.2)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_5);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.4)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_6);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.6)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_7);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.8)
                        {
                            SC2_EC_EW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_8);
                            if (SC2_EC_EW_b1 == 0)
                            {
                                SC2_EC_EW_b1 = 1;
                            }
                        }
                        ////for b2
                        if (EC_b2 == 0.2)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_0);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.4)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_1);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.6)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_2);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.8)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_3);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.0)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_4);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.2)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_5);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.4)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_6);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.6)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_7);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.8)
                        {
                            SC2_EC_EW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_EW_8);
                            if (SC2_EC_EW_b2 == 0)
                            {
                                SC2_EC_EW_b2 = 1;
                            }
                        }

                        if (SC2_EC_EW_b1 == SC2_EC_EW_b2)
                        {
                            SC2_EC_EW = SC2_EC_EW_b1;
                        }
                        else
                        {
                            SC2_EC_EW = SC2_EC_EW_b1 + ((SC2_EC_EW_b2 - SC2_EC_EW_b1) * ((EC_R1 - EC_b1) / (EC_b2 - EC_b1)));
                        }

                        Excel.Range mergeRange_AP5_22a = worksheet_AP5.Range["A22", "C22"];
                        mergeRange_AP5_22a.Merge();
                        mergeRange_AP5_22a.Value = "East-West";
                        mergeRange_AP5_22a.Font.Name = "Times New Roman";
                        mergeRange_AP5_22a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23a = worksheet_AP5.Cells[22, 4];
                        mergeRange_AP5_23a.Value = "SC2:";
                        mergeRange_AP5_23a.Font.Name = "Times New Roman";
                        mergeRange_AP5_23a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24a = worksheet_AP5.Cells[22, 5];
                        mergeRange_AP5_24a.Value = SC2_EC_EW.ToString("0.0000");
                        mergeRange_AP5_24a.Font.Name = "Times New Roman";
                        mergeRange_AP5_24a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25a = worksheet_AP5.Range["G22", "H22"];
                        mergeRange_AP5_25a.Merge();
                        mergeRange_AP5_25a.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25a.Font.Name = "Times New Roman";
                        mergeRange_AP5_25a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26a = worksheet_AP5.Cells[22, 9];
                        mergeRange_AP5_26a.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_EC_EW).ToString("0.0000");
                        mergeRange_AP5_26a.Font.Name = "Times New Roman";
                        mergeRange_AP5_26a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////


                        ////////////////NorthEast-NorthWest SC2_EC  
                        ////for b1
                        if (EC_b1 == 0.2)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_0);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.4)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_1);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.6)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_2);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.8)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_3);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.0)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_4);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.2)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_5);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.4)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_6);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.6)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_7);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.8)
                        {
                            SC2_EC_NENW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_8);
                            if (SC2_EC_NENW_b1 == 0)
                            {
                                SC2_EC_NENW_b1 = 1;
                            }
                        }
                        ////for b2
                        if (EC_b2 == 0.2)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_0);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.4)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_1);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.6)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_2);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.8)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_3);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.0)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_4);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.2)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_5);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.4)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_6);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.6)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_7);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.8)
                        {
                            SC2_EC_NENW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_NENW_8);
                            if (SC2_EC_NENW_b2 == 0)
                            {
                                SC2_EC_NENW_b2 = 1;
                            }
                        }

                        if (SC2_EC_NENW_b1 == SC2_EC_NENW_b2)
                        {
                            SC2_EC_NENW = SC2_EC_NENW_b1;
                        }
                        else
                        {
                            SC2_EC_NENW = SC2_EC_NENW_b1 + ((SC2_EC_NENW_b2 - SC2_EC_NENW_b1) * ((EC_R1 - EC_b1) / (EC_b2 - EC_b1)));
                        }

                        Excel.Range mergeRange_AP5_22b = worksheet_AP5.Range["A24", "C24"];
                        mergeRange_AP5_22b.Merge();
                        mergeRange_AP5_22b.Value = "NorthEast-NorthWest";
                        mergeRange_AP5_22b.Font.Name = "Times New Roman";
                        mergeRange_AP5_22b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23b = worksheet_AP5.Cells[24, 4];
                        mergeRange_AP5_23b.Value = "SC2:";
                        mergeRange_AP5_23b.Font.Name = "Times New Roman";
                        mergeRange_AP5_23b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24b = worksheet_AP5.Cells[24, 5];
                        mergeRange_AP5_24b.Value = SC2_EC_NENW.ToString("0.0000");
                        mergeRange_AP5_24b.Font.Name = "Times New Roman";
                        mergeRange_AP5_24b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25b = worksheet_AP5.Range["G24", "H24"];
                        mergeRange_AP5_25b.Merge();
                        mergeRange_AP5_25b.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25b.Font.Name = "Times New Roman";
                        mergeRange_AP5_25b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26b = worksheet_AP5.Cells[24, 9];
                        mergeRange_AP5_26b.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_EC_NENW).ToString("0.0000");
                        mergeRange_AP5_26b.Font.Name = "Times New Roman";
                        mergeRange_AP5_26b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////

                        ////////////////SouthEast-SouthWest SC2_EC  
                        ////for b1
                        if (EC_b1 == 0.2)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_0);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.4)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_1);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.6)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_2);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 0.8)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_3);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.0)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_4);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.2)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_5);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.4)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_6);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.6)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_7);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        else if (EC_b1 == 1.8)
                        {
                            SC2_EC_SESW_b1 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_8);
                            if (SC2_EC_SESW_b1 == 0)
                            {
                                SC2_EC_SESW_b1 = 1;
                            }
                        }
                        ////for b2
                        if (EC_b2 == 0.2)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_0);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.4)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_1);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.6)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_2);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 0.8)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_3);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.0)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_4);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.2)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_5);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.4)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_6);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.6)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_7);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }
                        else if (EC_b2 == 1.8)
                        {
                            SC2_EC_SESW_b2 = Interpolation_for_SC2_EC(Agl, EC_R2, dataGrid_SC2_EC_SESW_8);
                            if (SC2_EC_SESW_b2 == 0)
                            {
                                SC2_EC_SESW_b2 = 1;
                            }
                        }

                        if (SC2_EC_SESW_b1 == SC2_EC_SESW_b2)
                        {
                            SC2_EC_SESW = SC2_EC_SESW_b1;
                        }
                        else
                        {
                            SC2_EC_SESW = SC2_EC_SESW_b1 + ((SC2_EC_SESW_b2 - SC2_EC_SESW_b1) * ((EC_R1 - EC_b1) / (EC_b2 - EC_b1)));
                        }

                        Excel.Range mergeRange_AP5_22c = worksheet_AP5.Range["A26", "C26"];
                        mergeRange_AP5_22c.Merge();
                        mergeRange_AP5_22c.Value = "SouthEast-SouthWest";
                        mergeRange_AP5_22c.Font.Name = "Times New Roman";
                        mergeRange_AP5_22c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23c = worksheet_AP5.Cells[26, 4];
                        mergeRange_AP5_23c.Value = "SC2:";
                        mergeRange_AP5_23c.Font.Name = "Times New Roman";
                        mergeRange_AP5_23c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24c = worksheet_AP5.Cells[26, 5];
                        mergeRange_AP5_24c.Value = SC2_EC_SESW.ToString("0.0000");
                        mergeRange_AP5_24c.Font.Name = "Times New Roman";
                        mergeRange_AP5_24c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25c = worksheet_AP5.Range["G26", "H26"];
                        mergeRange_AP5_25c.Merge();
                        mergeRange_AP5_25c.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25c.Font.Name = "Times New Roman";
                        mergeRange_AP5_25c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26c = worksheet_AP5.Cells[26, 9];
                        mergeRange_AP5_26c.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * SC2_EC_SESW).ToString("0.0000");
                        mergeRange_AP5_26c.Font.Name = "Times New Roman";
                        mergeRange_AP5_26c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////


                        // Set the column width for Brief Description
                        worksheet_AP5.Columns[11].ColumnWidth = 10;

                    }

                    else
                    {
                        Excel.Range mergeRange_AP5_9 = worksheet_AP5.Range["E5", "G18"];
                        mergeRange_AP5_9.Merge();
                        string imagePath1 = BuildingCoder.Util.GetFilePath("Hor_Pro_Wndw.png");  //"D:\WIP\01_API\ETTV\Hor_Pro_Wndw.png"; // Specify the path to your PNG image

                        if (File.Exists(imagePath1))
                        {
                            // Calculate the position and size of the image to fit within the merge cell
                            float left = (float)mergeRange_AP5_9.Left;
                            float top = (float)mergeRange_AP5_9.Top;
                            float width = (float)mergeRange_AP5_9.Width;
                            float height = (float)mergeRange_AP5_9.Height;
                            // Add the image to the worksheet
                            worksheet_AP5.Shapes.AddPicture(imagePath1, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
                        }
                        else
                        {
                            MessageBox.Show("Image file not found at the specified path.");
                        }

                        Excel.Range mergeRange_AP5_10 = worksheet_AP5.Cells[5, 9];
                        mergeRange_AP5_10.Value = "Angle (deg):";
                        mergeRange_AP5_10.Font.Name = "Times New Roman";
                        mergeRange_AP5_10.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_11 = worksheet_AP5.Cells[5, 10];
                        mergeRange_AP5_11.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[5].Value.ToString();
                        mergeRange_AP5_11.Font.Name = "Times New Roman";
                        mergeRange_AP5_11.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_11.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_11.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_12 = worksheet_AP5.Cells[7, 9];
                        mergeRange_AP5_12.Value = "P (m):";
                        mergeRange_AP5_12.Font.Name = "Times New Roman";
                        mergeRange_AP5_12.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_13 = worksheet_AP5.Cells[7, 10];
                        mergeRange_AP5_13.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString();
                        mergeRange_AP5_13.Font.Name = "Times New Roman";
                        mergeRange_AP5_13.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_13.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_13.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_14 = worksheet_AP5.Cells[9, 9];
                        mergeRange_AP5_14.Value = "H (m):";
                        mergeRange_AP5_14.Font.Name = "Times New Roman";
                        mergeRange_AP5_14.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_15 = worksheet_AP5.Cells[9, 10];
                        mergeRange_AP5_15.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString();
                        mergeRange_AP5_15.Font.Name = "Times New Roman";
                        mergeRange_AP5_15.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_15.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_15.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_16 = worksheet_AP5.Cells[11, 9];
                        mergeRange_AP5_16.Value = "W (m):";
                        mergeRange_AP5_16.Font.Name = "Times New Roman";
                        mergeRange_AP5_16.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_17 = worksheet_AP5.Cells[11, 10];
                        mergeRange_AP5_17.Value = dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString();
                        mergeRange_AP5_17.Font.Name = "Times New Roman";
                        mergeRange_AP5_17.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_17.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_17.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_18 = worksheet_AP5.Cells[13, 9];
                        mergeRange_AP5_18.Value = "R1 = P/H:";
                        mergeRange_AP5_18.Font.Name = "Times New Roman";
                        mergeRange_AP5_18.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_19 = worksheet_AP5.Cells[13, 10];
                        double a = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString());
                        double b = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[7].Value.ToString());
                        double c = Math.Round((a / b), 4);
                        mergeRange_AP5_19.Value = c.ToString();
                        mergeRange_AP5_19.Font.Name = "Times New Roman";
                        mergeRange_AP5_19.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_19.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_19.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_20 = worksheet_AP5.Cells[15, 9];
                        mergeRange_AP5_20.Value = "R2 = P/W:";
                        mergeRange_AP5_20.Font.Name = "Times New Roman";
                        mergeRange_AP5_20.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ///////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_21 = worksheet_AP5.Cells[15, 10];
                        double a1 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[6].Value.ToString());
                        double b1 = double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[8].Value.ToString());
                        double c1 = Math.Round((a1 / b1), 4);
                        mergeRange_AP5_21.Value = c1.ToString();
                        mergeRange_AP5_21.Font.Name = "Times New Roman";
                        mergeRange_AP5_21.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_21.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_21.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_21.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_22 = worksheet_AP5.Range["A20", "C20"];
                        mergeRange_AP5_22.Merge();
                        mergeRange_AP5_22.Value = "North-South";
                        mergeRange_AP5_22.Font.Name = "Times New Roman";
                        mergeRange_AP5_22.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23 = worksheet_AP5.Cells[20, 4];
                        mergeRange_AP5_23.Value = "SC2:";
                        mergeRange_AP5_23.Font.Name = "Times New Roman";
                        mergeRange_AP5_23.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24 = worksheet_AP5.Cells[20, 5];                        
                        mergeRange_AP5_24.Value = 1;
                        mergeRange_AP5_24.Font.Name = "Times New Roman";
                        mergeRange_AP5_24.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25 = worksheet_AP5.Range["G20", "H20"];
                        mergeRange_AP5_25.Merge();
                        mergeRange_AP5_25.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25.Font.Name = "Times New Roman";
                        mergeRange_AP5_25.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26 = worksheet_AP5.Cells[20, 9];
                        mergeRange_AP5_26.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * 1).ToString("0.0000");
                        mergeRange_AP5_26.Font.Name = "Times New Roman";
                        mergeRange_AP5_26.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        Excel.Range mergeRange_AP5_22a = worksheet_AP5.Range["A22", "C22"];
                        mergeRange_AP5_22a.Merge();
                        mergeRange_AP5_22a.Value = "East-West";
                        mergeRange_AP5_22a.Font.Name = "Times New Roman";
                        mergeRange_AP5_22a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23a = worksheet_AP5.Cells[22, 4];
                        mergeRange_AP5_23a.Value = "SC2:";
                        mergeRange_AP5_23a.Font.Name = "Times New Roman";
                        mergeRange_AP5_23a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24a = worksheet_AP5.Cells[22, 5];
                        mergeRange_AP5_24a.Value = 1;
                        mergeRange_AP5_24a.Font.Name = "Times New Roman";
                        mergeRange_AP5_24a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25a = worksheet_AP5.Range["G22", "H22"];
                        mergeRange_AP5_25a.Merge();
                        mergeRange_AP5_25a.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25a.Font.Name = "Times New Roman";
                        mergeRange_AP5_25a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26a = worksheet_AP5.Cells[22, 9];
                        mergeRange_AP5_26a.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * 1).ToString("0.0000");
                        mergeRange_AP5_26a.Font.Name = "Times New Roman";
                        mergeRange_AP5_26a.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26a.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26a.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22b = worksheet_AP5.Range["A24", "C24"];
                        mergeRange_AP5_22b.Merge();
                        mergeRange_AP5_22b.Value = "NorthEast-NorthWest";
                        mergeRange_AP5_22b.Font.Name = "Times New Roman";
                        mergeRange_AP5_22b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23b = worksheet_AP5.Cells[24, 4];
                        mergeRange_AP5_23b.Value = "SC2:";
                        mergeRange_AP5_23b.Font.Name = "Times New Roman";
                        mergeRange_AP5_23b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24b = worksheet_AP5.Cells[24, 5];
                        mergeRange_AP5_24b.Value = 1;
                        mergeRange_AP5_24b.Font.Name = "Times New Roman";
                        mergeRange_AP5_24b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25b = worksheet_AP5.Range["G24", "H24"];
                        mergeRange_AP5_25b.Merge();
                        mergeRange_AP5_25b.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25b.Font.Name = "Times New Roman";
                        mergeRange_AP5_25b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26b = worksheet_AP5.Cells[24, 9];
                        mergeRange_AP5_26b.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * 1).ToString("0.0000");
                        mergeRange_AP5_26b.Font.Name = "Times New Roman";
                        mergeRange_AP5_26b.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26b.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26b.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;


                        Excel.Range mergeRange_AP5_22c = worksheet_AP5.Range["A26", "C26"];
                        mergeRange_AP5_22c.Merge();
                        mergeRange_AP5_22c.Value = "SouthEast-SouthWest";
                        mergeRange_AP5_22c.Font.Name = "Times New Roman";
                        mergeRange_AP5_22c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_23c = worksheet_AP5.Cells[26, 4];
                        mergeRange_AP5_23c.Value = "SC2:";
                        mergeRange_AP5_23c.Font.Name = "Times New Roman";
                        mergeRange_AP5_23c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_24c = worksheet_AP5.Cells[26, 5];
                        mergeRange_AP5_24c.Value = 1;
                        mergeRange_AP5_24c.Font.Name = "Times New Roman";
                        mergeRange_AP5_24c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_24c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_24c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_25c = worksheet_AP5.Range["G26", "H26"];
                        mergeRange_AP5_25c.Merge();
                        mergeRange_AP5_25c.Value = "SC = SC1 x SC2";
                        mergeRange_AP5_25c.Font.Name = "Times New Roman";
                        mergeRange_AP5_25c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        ////////////////////////////////////////////////////////////////////////////////////
                        Excel.Range mergeRange_AP5_26c = worksheet_AP5.Cells[26, 9];
                        mergeRange_AP5_26c.Value = (double.Parse(dataGrid_Wndw_Types.Rows[num_fc].Cells[3].Value.ToString()) * 1).ToString("0.0000");
                        mergeRange_AP5_26c.Font.Name = "Times New Roman";
                        mergeRange_AP5_26c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        mergeRange_AP5_26c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                        mergeRange_AP5_26c.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        // Set the column width for Brief Description
                        worksheet_AP5.Columns[9].ColumnWidth = 10;

                    }




                        num_fc = num_fc + 1;


                    //dataGrid_Wndw_Types.Rows[num1].Cells[3].Value.ToString()
                    //dataGrid_Wndw_Types.Rows[num1].Cells[2].Value.ToString();


                }



                ///////////////////////////////////////////////////////////// for AP5 Worksheets - Ends ////////////////////////////////////////////////////////////////





                /////////////////////// Continue from Cover Page starts /////////////////////////////
                //////(a)
                worksheet.Cells[31, 1].Font.Name = "Times New Roman";
                worksheet.Cells[31, 1].Value = "(a)";
                worksheet.Cells[31, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[31, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                Excel.Range rangeA = worksheet.Cells[31, 2];
                rangeA.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                rangeA.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                MergeCellsAndApplyBorders(worksheet.Range["C31", "J31"], "Sheets of Appendix 1");
                // Count the number of worksheets with names containing "AP1"
                int count1 = 0;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("AP1"))
                    {
                        count1++;
                    }
                }
                // Write the count to cell [31,2]
                worksheet.Cells[31, 2].Value = count1;
                worksheet.Cells[31, 2].Font.Name = "Times New Roman";
                worksheet.Cells[31, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////(b)
                worksheet.Cells[32, 1].Font.Name = "Times New Roman";
                worksheet.Cells[32, 1].Value = "(b)";
                worksheet.Cells[32, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[32, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                Excel.Range rangeB = worksheet.Cells[32, 2];
                rangeB.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                rangeB.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                MergeCellsAndApplyBorders(worksheet.Range["C32", "J32"], "Sheets of Appendix 2");
                // Count the number of worksheets with names containing "AP2"
                int count2 = 0;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("AP2"))
                    {
                        count2++;
                    }
                }
                // Write the count to cell [31,2]
                worksheet.Cells[32, 2].Value = count2;
                worksheet.Cells[32, 2].Font.Name = "Times New Roman";
                worksheet.Cells[32, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////(c)
                worksheet.Cells[33, 1].Font.Name = "Times New Roman";
                worksheet.Cells[33, 1].Value = "(c)";
                worksheet.Cells[33, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[33, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                Excel.Range rangeC = worksheet.Cells[33, 2];
                rangeC.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                rangeC.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                MergeCellsAndApplyBorders(worksheet.Range["C33", "J33"], "Sheets of Appendix 3");
                // Count the number of worksheets with names containing "AP3"
                int count3 = 0;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("AP3"))
                    {
                        count3++;
                    }
                }
                // Write the count to cell [33,2]
                worksheet.Cells[33, 2].Value = count3;
                worksheet.Cells[33, 2].Font.Name = "Times New Roman";
                worksheet.Cells[33, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////(d)
                worksheet.Cells[34, 1].Font.Name = "Times New Roman";
                worksheet.Cells[34, 1].Value = "(d)";
                worksheet.Cells[34, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[34, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                Excel.Range rangeD = worksheet.Cells[34, 2];
                rangeD.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                rangeD.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                MergeCellsAndApplyBorders(worksheet.Range["C34", "J34"], "Sheets of Appendix 4");
                // Count the number of worksheets with names containing "AP4"
                int count4 = 0;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("AP4"))
                    {
                        count4++;
                    }
                }
                // Write the count to cell [34,2]
                worksheet.Cells[34, 2].Value = count4;
                worksheet.Cells[34, 2].Font.Name = "Times New Roman";
                worksheet.Cells[34, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////(e)
                worksheet.Cells[35, 1].Font.Name = "Times New Roman";
                worksheet.Cells[35, 1].Value = "(e)";
                worksheet.Cells[35, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[35, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                Excel.Range rangeE = worksheet.Cells[35, 2];
                rangeE.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                rangeE.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                MergeCellsAndApplyBorders(worksheet.Range["C35", "J35"], "Sheets of detailed calculations on U-values and shading coefficients");
                // Count the number of worksheets with names containing "Summary"
                int count5 = 0;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("AP5"))
                    {
                        count5++;
                    }
                }
                // Write the count to cell [35,2]
                worksheet.Cells[35, 2].Value = count5;
                worksheet.Cells[35, 2].Font.Name = "Times New Roman";
                worksheet.Cells[35, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //////(f)
                worksheet.Cells[36, 1].Font.Name = "Times New Roman";
                worksheet.Cells[36, 1].Value = "(f)";
                worksheet.Cells[36, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[36, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                Excel.Range rangeF = worksheet.Cells[36, 2];
                rangeF.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
                rangeF.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                MergeCellsAndApplyBorders(worksheet.Range["C36", "J36"], "Sheets of drawings/sketches");
                // Count the number of worksheets with names containing "Others"
                int count6 = 0;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Contains("Others"))
                    {
                        count6++;
                    }
                }
                // Write the count to cell [36,2]
                worksheet.Cells[36, 2].Value = count6;
                worksheet.Cells[36, 2].Font.Name = "Times New Roman";
                worksheet.Cells[36, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                /////////////////////// Continue from Cover Page ends /////////////////////////////

                // Prompt the user for the file name
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = "xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    // Save the workbook to the specified file path
                    workbook.SaveAs(filePath);

                    MessageBox.Show("Excel file generated successfully!");
                }
                else
                {
                    MessageBox.Show("Operation cancelled by the user.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating Excel file: " + ex.Message);
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }


        private void AddWorksheetsForAP1Tabs()
        {
            // Get the number of existing worksheets
            int existingWorksheetsCount = workbook.Worksheets.Count;

            // Iterate through tab pages in reverse order
            for (int i = Main_TabPage.TabPages.Count - 1; i >= 0; i--)
            {
                TabPage tabPage = Main_TabPage.TabPages[i];

                // Check if the tab page name contains "AP1"
                if (tabPage.Text.Contains("AP1"))
                {
                    // Add a new worksheet for each tab page                    
                    Microsoft.Office.Interop.Excel.Worksheet worksheet1 = workbook.Worksheets.Add(After: workbook.Worksheets[existingWorksheetsCount]);
                    worksheet1.Name = tabPage.Text;

                    // Set the tab color to light blue
                    worksheet1.Tab.Color = System.Drawing.Color.LightBlue;

                    // Extract data from tab page controls and populate the worksheet
                    //ExtractDataFromTabPage(tabPage, worksheet1);
                }
            }

            ////////////////////// to activate back the Cover Page///////////////////////////////////
            // Find the worksheet with the name "Cover Page"
            Microsoft.Office.Interop.Excel.Worksheet coverPageWorksheet = null;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == "Cover Page")
                {
                    coverPageWorksheet = sheet;
                    break;
                }
            }

            // Check if the worksheet exists
            if (coverPageWorksheet != null)
            {
                // Activate the worksheet
                coverPageWorksheet.Activate();
            }


            // Save the changes to the existing Excel file
            workbook.Save();

        }

        private void AddWorksheetsForAP3Tabs()
        {
            // Get the number of existing worksheets
            int existingWorksheetsCount = workbook.Worksheets.Count;

            // Iterate through tab pages in reverse order
            for (int i = Main_TabPage.TabPages.Count - 1; i >= 0; i--)
            {
                TabPage tabPage = Main_TabPage.TabPages[i];

                // Check if the tab page name contains "AP1"
                if (tabPage.Text.Contains("AP3"))
                {
                    // Add a new worksheet for each tab page                    
                    Microsoft.Office.Interop.Excel.Worksheet worksheet3 = workbook.Worksheets.Add(After: workbook.Worksheets[existingWorksheetsCount]);
                    worksheet3.Name = tabPage.Text;

                    // Set the tab color to light blue
                    worksheet3.Tab.Color = System.Drawing.Color.LightGreen;

                    // Extract data from tab page controls and populate the worksheet
                    //ExtractDataFromTabPage(tabPage, worksheet1);
                }
            }

            ////////////////////// to activate back the Cover Page///////////////////////////////////
            // Find the worksheet with the name "Cover Page"
            Microsoft.Office.Interop.Excel.Worksheet coverPageWorksheet = null;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == "Cover Page")
                {
                    coverPageWorksheet = sheet;
                    break;
                }
            }

            // Check if the worksheet exists
            if (coverPageWorksheet != null)
            {
                // Activate the worksheet
                coverPageWorksheet.Activate();
            }


            // Save the changes to the existing Excel file
            workbook.Save();

        }

        private void AddWorksheetsForAP5Tabs()
        {
            // Get the number of existing worksheets
            int existingWorksheetsCount = workbook.Worksheets.Count;

            // Iterate through tab pages in reverse order
            for (int i = Main_TabPage.TabPages.Count - 1; i >= 0; i--)
            {
                TabPage tabPage = Main_TabPage.TabPages[i];

                // Check if the tab page name contains "AP1"
                if (tabPage.Text.Contains("AP5"))
                {
                    // Add a new worksheet for each tab page                    
                    Microsoft.Office.Interop.Excel.Worksheet worksheet5 = workbook.Worksheets.Add(After: workbook.Worksheets[existingWorksheetsCount]);
                    worksheet5.Name = tabPage.Text;

                    // Set the tab color to light blue
                    worksheet5.Tab.Color = System.Drawing.Color.DarkCyan;

                    // Extract data from tab page controls and populate the worksheet
                    //ExtractDataFromTabPage(tabPage, worksheet1);

                                        
                    

                }
            }

            ////////////////////// to activate back the Cover Page///////////////////////////////////
            // Find the worksheet with the name "Cover Page"
            Microsoft.Office.Interop.Excel.Worksheet coverPageWorksheet = null;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == "Cover Page")
                {
                    coverPageWorksheet = sheet;
                    break;
                }
            }

            // Check if the worksheet exists
            if (coverPageWorksheet != null)
            {
                // Activate the worksheet
                coverPageWorksheet.Activate();
            }


            // Save the changes to the existing Excel file
            workbook.Save();

        }

        private void MergeCellsAndApplyBorders(Excel.Range range, string value)
        {
            // Merge cells and set alignment
            range.Merge();
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Set alignment to left
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            // Set font to Times New Roman
            range.Font.Name = "Times New Roman";

            // Apply right border to the merged cells
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

            // Set value for the merged cells
            range.Value = value;
        }

        private void ActivateWorksheet(Microsoft.Office.Interop.Excel.Workbook workbook, string worksheetName)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == worksheetName)
                {
                    sheet.Activate();
                    break;
                }
            }
        }





    }

}