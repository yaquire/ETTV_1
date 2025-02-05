using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using System.Collections;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
using BuildingCoder;
using Autodesk.Revit.DB.IFC;
using Autodesk.Revit.DB.Visual;
using System.ComponentModel;
using System.Reflection;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using View = Autodesk.Revit.DB.View;
using Autodesk.Revit.UI.Events;


namespace ETTV_1
{
    [TransactionAttribute(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ETTV_1 : IExternalCommand
    {
        
        private List<Element> ExtWallLst = new List<Element>();
        
        private List<ElementId> ids = new List<ElementId>();
        private List<ElementId> ids_1 = new List<ElementId>();
        private List<ElementId> ids_2 = new List<ElementId>();
        private List<ElementId> ids_3 = new List<ElementId>();
        private List<ElementId> ids_4 = new List<ElementId>();

        const double _inchToMm = 25.4;
        const double _footToMm = 12 * _inchToMm;

        //private string Ex_Wall_Option_St;
        private int Ex_Wall_Option_Int;
        private List<Element> South_Wall_Lst = new List<Element>();
        private List<Element> North_Wall_Lst = new List<Element>();
        private List<Element> East_Wall_Lst = new List<Element>();
        private List<Element> West_Wall_Lst = new List<Element>();
        private List<Element> SouthEast_Wall_Lst = new List<Element>();
        private List<Element> SouthWest_Wall_Lst = new List<Element>();
        private List<Element> NorthEast_Wall_Lst = new List<Element>();
        private List<Element> NorthWest_Wall_Lst = new List<Element>();

        private Parameter ParaSouthWall;
        private Parameter ParaNorthWall;
        private Parameter ParaEastWall;
        private Parameter ParaWestWall;
        private Parameter ParaSouthEastWall;
        private Parameter ParaSouthWestWall;
        private Parameter ParaNorthEastWall;
        private Parameter ParaNorthWestWall;

        private int spCount;
        private int iBoundary;
        private int iSegment;
        private Element neighbour;
        private List<Element> neighbourLst;
        private Curve curve;
        private double length;
        private List<double> lengthLst;
        private List<double> lengthLst1;
        private double wallLength;
        private double area;
        private List<Element> NewWallLst;
        private List<Element> NewWallLst1;
        private List<double> wallLengthLst;
        private List<Element> ExtWallLst1 = new List<Element>();        
        private List<Element> WndwList;
        private List<Element> WndwList1;
        //private List<Element> ColnList;
        private ICollection<ElementId> intersects;
        private IList<IList<BoundarySegment>> boundaries_0;
        private List<List<BoundarySegment>> boundaries_1;
        private List<BoundarySegment> BLst;
        IList<IList<BoundarySegment>> allBoundarySegments;
        private IList<IList<BoundarySegment>> wallBoundarySegments;
        

        private List<List<Wall>> roomOuterWall = new List<List<Wall>>();

        private ICollection<ElementId> intersects_1;
        private ICollection<ElementId> intersects_2;
        private List<Wall> intersects_wall;
        private List<ElementId> intersects_wall_ids;
        
        //for windows
        private List<Element> N_WndwList;
        private List<ElementId> N_Wndw_Id_List;
        private string N_Wndw_Strg;
        private string N_Wndw_Strg1;
        private string N_Wndw_Area;
        private string N_Wndw_U;
        private string N_Wndw_SC1;

        private List<Element> S_WndwList;
        private List<ElementId> S_Wndw_Id_List;
        private string S_Wndw_Strg;
        private string S_Wndw_Strg1;
        private string S_Wndw_Area;
        private string S_Wndw_U;
        private string S_Wndw_SC1;

        private List<Element> E_WndwList;
        private List<ElementId> E_Wndw_Id_List;
        private string E_Wndw_Strg;
        private string E_Wndw_Strg1;
        private string E_Wndw_Area;
        private string E_Wndw_U;
        private string E_Wndw_SC1;

        private List<Element> W_WndwList;
        private List<ElementId> W_Wndw_Id_List;
        private string W_Wndw_Strg;
        private string W_Wndw_Strg1;
        private string W_Wndw_Area;
        private string W_Wndw_U;
        private string W_Wndw_SC1;

        private List<Element> NE_WndwList;
        private List<ElementId> NE_Wndw_Id_List;
        private string NE_Wndw_Strg;
        private string NE_Wndw_Strg1;
        private string NE_Wndw_Area;
        private string NE_Wndw_U;
        private string NE_Wndw_SC1;

        private List<Element> NW_WndwList;
        private List<ElementId> NW_Wndw_Id_List;
        private string NW_Wndw_Strg;
        private string NW_Wndw_Strg1;
        private string NW_Wndw_Area;
        private string NW_Wndw_U;
        private string NW_Wndw_SC1;

        private List<Element> SE_WndwList;
        private List<ElementId> SE_Wndw_Id_List;
        private string SE_Wndw_Strg;
        private string SE_Wndw_Strg1;
        private string SE_Wndw_Area;
        private string SE_Wndw_U;
        private string SE_Wndw_SC1;

        private List<Element> SW_WndwList;
        private List<ElementId> SW_Wndw_Id_List;
        private string SW_Wndw_Strg;
        private string SW_Wndw_Strg1;
        private string SW_Wndw_Area;
        private string SW_Wndw_U;
        private string SW_Wndw_SC1;
        //for windows

        //for walls
        private List<Element> N_Wall_List;
        private List<ElementId> N_Wall_Id_List;
        private string N_Wall_Strg;
        private string N_Wall_Strg1;
        private string N_Wall_Area;
        private List<double> N_Wall_Lgt_List;
        private List<double> N_Wall_InstArea_List;
        private string N_Wall_U;
        
        private List<Element> S_Wall_List;
        private List<ElementId> S_Wall_Id_List;
        private string S_Wall_Strg;
        private string S_Wall_Strg1;
        private string S_Wall_Area;
        private List<double> S_Wall_Lgt_List;
        private List<double> S_Wall_InstArea_List;
        private string S_Wall_U;
        

        private List<Element> E_Wall_List;
        private List<ElementId> E_Wall_Id_List;
        private string E_Wall_Strg;
        private string E_Wall_Strg1;
        private string E_Wall_Area;
        private List<double> E_Wall_Lgt_List;
        private List<double> E_Wall_InstArea_List;
        private string E_Wall_U;

        private List<Element> W_Wall_List;
        private List<ElementId> W_Wall_Id_List;
        private string W_Wall_Strg;
        private string W_Wall_Strg1;
        private string W_Wall_Area;
        private List<double> W_Wall_Lgt_List;
        private List<double> W_Wall_InstArea_List;
        private string W_Wall_U;

        private List<Element> NE_Wall_List;
        private List<ElementId> NE_Wall_Id_List;
        private string NE_Wall_Strg;
        private string NE_Wall_Strg1;
        private string NE_Wall_Area;
        private List<double> NE_Wall_Lgt_List;
        private List<double> NE_Wall_InstArea_List;
        private string NE_Wall_U;

        private List<Element> NW_Wall_List;
        private List<ElementId> NW_Wall_Id_List;
        private string NW_Wall_Strg;
        private string NW_Wall_Strg1;
        private string NW_Wall_Area;
        private List<double> NW_Wall_Lgt_List;
        private List<double> NW_Wall_InstArea_List;
        private string NW_Wall_U;

        private List<Element> SE_Wall_List;
        private List<ElementId> SE_Wall_Id_List;
        private string SE_Wall_Strg;
        private string SE_Wall_Strg1;
        private string SE_Wall_Area;
        private List<double> SE_Wall_Lgt_List;
        private List<double> SE_Wall_InstArea_List;
        private string SE_Wall_U;

        private List<Element> SW_Wall_List;
        private List<ElementId> SW_Wall_Id_List;
        private string SW_Wall_Strg;
        private string SW_Wall_Strg1;
        private string SW_Wall_Area;
        private List<double> SW_Wall_Lgt_List;
        private List<double> SW_Wall_InstArea_List;
        private string SW_Wall_U;
        //for walls

        private int num1;
        //private int num2;
        private int num3;
        private int num4;
        private IList<ElementId> inserts;
        private List<Element> InsertLst;
        private List<Element> InsertLst1;
        private double TotalInstArea;
        private IList<ElementId> inserts_0;
        private IList<Element> MainWndwList;
        private List<ElementId> WndwTypeId_1 = new List<ElementId>();
        private List<ElementId> WndwTypeId_2 = new List<ElementId>();
        private List<ElementType> WndwTypeLst = new List<ElementType>();
        private List<ElementType> WndwTypeLst1 = new List<ElementType>();

        private List<ElementId> WallTypeId_1 = new List<ElementId>();
        private List<ElementId> WallTypeId_2 = new List<ElementId>();
        private List<ElementType> WallTypeLst = new List<ElementType>();

        private int A;
        private int B;
        private int C;

        


        ////////////////////////////////////////////////////////////////////
        public class RoomBoundarySegmentObject
        {
            public List<List<Curve>> curves { get; set; }
            public List<List<BoundarySegment>> boundarySegments { get; set; }
        }

        public static List<BoundarySegment> GetRoomOuterBoundaryCurves(Room room)
        {
            List<BoundarySegment> bss = new List<BoundarySegment>();
            RoomBoundarySegmentObject rbso = GetRoomBoundaryCurves(room);

            if (rbso.curves == null || rbso.curves.Count == 0)
            {
                return bss;
            }

            List<double> curveLoopAreas = new List<double>();

            try
            {
                double maxArea = 0.0;
                int maxIndex = -1;

                for (int i = 0; i < rbso.curves.Count; i++)
                {
                    CurveLoop curveLoop = CurveLoop.Create(rbso.curves[i]);
                    IList<CurveLoop> cLoop = new List<CurveLoop> { curveLoop };
                    double curveArea = ExporterIFCUtils.ComputeAreaOfCurveLoops(cLoop);
                    curveLoopAreas.Add(curveArea);

                    if (curveArea > maxArea)
                    {
                        maxArea = curveArea;
                        maxIndex = i;
                    }
                }

                if (maxIndex != -1)
                {
                    bss = rbso.boundarySegments[maxIndex];
                }
            }
            catch (Exception ex)
            {

            }
            return bss;
        }

        public static RoomBoundarySegmentObject GetRoomBoundaryCurves(Room room)
        {
            List<List<Curve>> boundaryCurves = new List<List<Curve>>();
            List<List<BoundarySegment>> bss = new List<List<BoundarySegment>>();

            // Retrieve the boundary segments of the room
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            IList<IList<BoundarySegment>> boundarySegments = room.GetBoundarySegments(options);

            // Extract curves from the boundary segments and maintain the nested list structure
            if (boundarySegments != null)
            {
                foreach (IList<BoundarySegment> loop in boundarySegments)
                {
                    List<Curve> loopCurves = new List<Curve>();
                    List<BoundarySegment> bs = new List<BoundarySegment>();
                    foreach (BoundarySegment segment in loop)
                    {
                        Curve curve = segment.GetCurve();
                        loopCurves.Add(curve);
                        bs.Add(segment);
                    }
                    boundaryCurves.Add(loopCurves);
                    bss.Add(bs);
                }
            }
            RoomBoundarySegmentObject roomBoundarySegmentObject = new RoomBoundarySegmentObject();
            roomBoundarySegmentObject.curves = boundaryCurves;
            roomBoundarySegmentObject.boundarySegments = bss;

            return roomBoundarySegmentObject;
        }


        ///////////////////////////////////////////////////////////////////


        //public List<ElementType> MyList
        //{
        //get { return WallTypeLst; }
        //}


        //public List<ElementType> GetList()
        //{
        //return WallTypeLst;
        //}

        /// <summary>
        /// Return a bounding box around all the 
        /// walls in the entire model; for just a
        /// building, or several buildings, this is 
        /// obviously equal to the model extents.
        /// </summary>
        static BoundingBoxXYZ GetBoundingBoxAroundAllWalls(Document doc, View view = null)
        {
            BoundingBoxXYZ bb = new BoundingBoxXYZ();
            FilteredElementCollector walls = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            foreach (Wall wall in walls)
            {
                bb.ExpandToContain(wall.get_BoundingBox(null));
            }
            return bb;
        }

        /// <summary>
        /// Filter out the required walls --
        /// Return all walls that are generating boundary
        /// segments for the given room. Includes debug
        /// code to compare wall lengths and wall areas.
        /// </summary>
        static List<ElementId> RetrieveWallsGeneratingRoomBoundaries(Document doc, Room room)
        {
            List<ElementId> ids = new List<ElementId>();

            IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

            int n = boundaries.Count;

            int iBoundary = 0, iSegment;

            foreach (IList<BoundarySegment> b in boundaries)
            {
                ++iBoundary;
                iSegment = 0;
                foreach (BoundarySegment s in b)
                {
                    ++iSegment;

                    // Retrieve the id of the element that 
                    // produces this boundary segment

                    Element neighbour = doc.GetElement(s.ElementId);

                    Curve curve = s.GetCurve();
                    double length = curve.Length;

                    if (neighbour is Wall)
                    {
                        Wall wall = neighbour as Wall;

                        Parameter p = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);

                        double area = p.AsDouble();

                        LocationCurve lc = wall.Location as LocationCurve;

                        double wallLength = lc.Curve.Length;

                        ids.Add(wall.Id);
                    }
                }
            }
            return ids;
        }

        /// <summary>
        /// 获取当前模型指定视图内的所有最外层的墙体
        /// Get all the outermost walls in the 
        /// specified view of the current model
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="view">视图,默认是当前激活的视图 
        /// View, default is currently active view</param>

        public static List<ElementId> GetOutermostWalls(Document doc,View view = null)
        {
            Level lowestLevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList()
                .OrderBy(level => level.Elevation)
                .ToList().First();

            double offset = Util.MmToFoot(1000);

            if (view == null)
            {
                view = doc.ActiveView;
            }

            List<View> views = new FilteredElementCollector(doc).OfClass(typeof(View)).ToList().Cast<View>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.FloorPlan).ToList();

            List<ElementId> returnIds = new List<ElementId>();
            foreach (View v in views)
            {
                view = v;

                BoundingBoxXYZ bb = GetBoundingBoxAroundAllWalls(doc, view);

                XYZ voffset = offset * (XYZ.BasisX + XYZ.BasisY);
                bb.Min -= voffset;
                bb.Max += voffset;

                XYZ[] bottom_corners = Util.GetBottomCorners(bb, 0);

                CurveArray curves = new CurveArray();
                for (int i = 0; i < 4; ++i)
                {
                    int j = i < 3 ? i + 1 : 0;
                    curves.Append(Line.CreateBound(bottom_corners[i], bottom_corners[j]));
                }

                using (TransactionGroup group = new TransactionGroup(doc))
                {
                    Room newRoom = null;

                    group.Start("Find Outermost Walls");

                    using (Transaction transaction = new Transaction(doc))
                    {
                        transaction.Start("Create New Room Boundary Lines");

                        SketchPlane sketchPlane = SketchPlane.Create(doc, view.GenLevel.Id);

                        ModelCurveArray modelCaRoomBoundaryLines = doc.Create.NewRoomBoundaryLines(sketchPlane, curves, view);

                        // 创建房间的坐标点 -- Create room coordinates

                        double d = Util.MmToFoot(600);
                        UV point = new UV(bb.Min.X + d, bb.Min.Y + d);

                        // 根据选中点，创建房间 当前视图的楼层 doc.ActiveView.GenLevel
                        // Create room at selected point on the current view level

                        newRoom = doc.Create.NewRoom(view.GenLevel, point);

                        if (newRoom == null)
                        {
                            string msg = "创建房间失败。";
                            TaskDialog.Show("xx", msg);
                            transaction.RollBack();
                            return null;
                        }

                        //newRoom.LookupParameter("Limit Offset").Set(bb.Max.Z - bb.Min.Z);

                        RoomTag tag = doc.Create.NewRoomTag(new LinkElementId(newRoom.Id), point, view.Id);

                        transaction.Commit();
                    }

                    //获取房间的墙体 -- Get the room walls

                    List<ElementId> ids = RetrieveWallsGeneratingRoomBoundaries(doc, newRoom);
                    returnIds.AddRange(ids);

                    group.RollBack();

                    //group.Commit();

                    //return ids;
                }
            }
            return returnIds;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Assembly assembly = Assembly.GetExecutingAssembly();
            string assemblyPath = assembly.Location;

            ETTV_F1 F1 = new ETTV_F1(commandData);
            F1.ShowDialog();

            #region Progress

            // Initialize and show the progress dialog
            ProgressDialog progressDialog = new ProgressDialog();
            progressDialog.Show();

            void UpdateProgress(int progressPercentage)
            {
                progressDialog.UpdateProgress(progressPercentage, $"{progressPercentage}%"); 
            }

            UpdateProgress(0);
            System.Threading.Thread.Sleep(500);

            #endregion

            Ex_Wall_Option_Int = F1.Ex_Wall_Op;

            //for Auto Pick Exterior Walls
            if (Ex_Wall_Option_Int == 2) //for Auto Pick Exterior Walls
            {
                //Getting External Wall Elements from Id List
                ids = GetOutermostWalls(doc); //Id List
                ids_1 = ids.Distinct().ToList(); // remove duplicate from Id List

                // Selecting Walls 
                //uidoc.Selection.SetElementIds(ids_1);

                foreach (ElementId I in ids_1)
                {
                    ExtWallLst.Add(doc.GetElement(I));
                }
            }

            //for Manual Pick Exterior Walls
            else if (Ex_Wall_Option_Int == 1) //for Manual Pick Exterior Walls
            {
                //Getting walls by selection
                IList<Reference> refe = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element,
                    "Pick Exterior Walls");

                foreach (Reference P in refe)
                {
                    ExtWallLst.Add(doc.GetElement(P.ElementId));
                }

                foreach (Element E in ExtWallLst)
                {
                    ids_2.Add(E.Id);
                }
                // Selecting Walls 
                //uidoc.Selection.SetElementIds(ids_2);
            }

            // Initialise the wall orientation, set all to uncheck
            foreach (Wall WE in ExtWallLst)
            {
                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    WE.LookupParameter("North").Set(0);
                    WE.LookupParameter("South").Set(0);
                    WE.LookupParameter("East").Set(0);
                    WE.LookupParameter("West").Set(0);
                    WE.LookupParameter("NorthEast").Set(0);
                    WE.LookupParameter("NorthWest").Set(0);
                    WE.LookupParameter("SouthEast").Set(0);
                    WE.LookupParameter("SouthWest").Set(0);

                    trans.Commit();
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(10);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Initialise the Window orientation, set all to uncheck           
            FilteredElementCollector collector_wndw = new FilteredElementCollector(doc);
            ElementCategoryFilter filter_wndw = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            IList<Element> All_Wndw_Lst =
                collector_wndw.WherePasses(filter_wndw).WhereElementIsNotElementType().ToElements();
            foreach (Element wndw in All_Wndw_Lst)
            {
                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();
                    wndw.LookupParameter("North").Set(0);
                    wndw.LookupParameter("South").Set(0);
                    wndw.LookupParameter("East").Set(0);
                    wndw.LookupParameter("West").Set(0);
                    wndw.LookupParameter("NorthEast").Set(0);
                    wndw.LookupParameter("NorthWest").Set(0);
                    wndw.LookupParameter("SouthEast").Set(0);
                    wndw.LookupParameter("SouthWest").Set(0);
                    trans.Commit();
                }
            }

            // Check orientation of Walls and Windows and Set the parameter values accordingly
            foreach (Wall WE in ExtWallLst)
            {
                XYZ exteriorDirection = GetExteriorWallDirection(WE);

                exteriorDirection = TransformByProjectLocation(exteriorDirection);

                //Checking for South Facing Walls & Windows
                bool isSouthFacing = IsSouthFacing(exteriorDirection);
                if (isSouthFacing)
                {
                    South_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaSouthWall = WE.LookupParameter("South");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaSouthWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("South").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for North Facing Walls & Windows
                bool isNorthFacing = IsNorthFacing(exteriorDirection);
                if (isNorthFacing)
                {
                    North_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaNorthWall = WE.LookupParameter("North");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaNorthWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("North").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for West Facing Walls & Windows
                bool isWestFacing = IsWestFacing(exteriorDirection);
                if (isWestFacing)
                {
                    West_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaWestWall = WE.LookupParameter("West");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaWestWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("West").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for East Facing Walls & Windows
                bool isEastFacing = IsEastFacing(exteriorDirection);
                if (isEastFacing)
                {
                    East_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaEastWall = WE.LookupParameter("East");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaEastWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("East").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for SouthEast Facing Walls & Windows
                bool isSouthEastFacing = IsSouthEastFacing(exteriorDirection);
                if (isSouthEastFacing)
                {
                    SouthEast_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaSouthEastWall = WE.LookupParameter("SouthEast");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaSouthEastWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("SouthEast").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for SouthWest Facing Walls & Windows
                bool isSouthWestFacing = IsSouthWestFacing(exteriorDirection);
                if (isSouthWestFacing)
                {
                    SouthWest_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaSouthWestWall = WE.LookupParameter("SouthWest");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaSouthWestWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("SouthWest").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for NorthEast Facing Walls & Windows
                bool isNorthEastFacing = IsNorthEastFacing(exteriorDirection);
                if (isNorthEastFacing)
                {
                    NorthEast_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaNorthEastWall = WE.LookupParameter("NorthEast");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaNorthEastWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("NorthEast").Set(1);
                        }
                        trans.Commit();
                    }
                }

                //Checking for NorthWest Facing Walls & Windows
                bool isNorthWestFacing = IsNorthWestFacing(exteriorDirection);
                if (isNorthWestFacing)
                {
                    NorthWest_Wall_Lst.Add(WE);

                    inserts_0 = (WE as HostObject).FindInserts(true, true, true, true);

                    ParaNorthWestWall = WE.LookupParameter("NorthWest");
                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();
                        ParaNorthWestWall.Set(1);

                        foreach (ElementId emtid in inserts_0)
                        {
                            Element emt = doc.GetElement(emtid);
                            emt.LookupParameter("NorthWest").Set(1);
                        }
                        trans.Commit();
                    }
                }
            }

            //////////////////////////// check how many types of Ext walls we have in this project //////////////////////

            //Check TypeIDs of Ext walls
            foreach (Element w in ExtWallLst)
            {
                ElementId typeId_w = w.GetTypeId();
                WallTypeId_1.Add(typeId_w);
            }

            WallTypeId_2 = WallTypeId_1.Distinct().ToList();

            //Check Types of walls
            foreach (ElementId WallTypeId in WallTypeId_2)
            {
                ElementType type = doc.GetElement(WallTypeId) as ElementType;
                WallTypeLst.Add(type);
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(20);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////////////////////////// check how many types of windows we have in this project //////////////////////

            FilteredElementCollector collector_Wndw = new FilteredElementCollector(doc);
            ElementCategoryFilter filter_Wndw = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            MainWndwList = collector_Wndw.WherePasses(filter_Wndw).WhereElementIsNotElementType().ToElements();

            //Check TypeIDs of windows
            foreach (Element Wndw in MainWndwList)
            {
                ElementId typeId = Wndw.GetTypeId();
                WndwTypeId_1.Add(typeId);
            }

            WndwTypeId_2 = WndwTypeId_1.Distinct().ToList();

            //Check Types of windows
            foreach (ElementId WndwId in WndwTypeId_2)
            {
                ElementType type = doc.GetElement(WndwId) as ElementType;
                WndwTypeLst.Add(type);
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Get Exterior Wall Direction
            XYZ GetExteriorWallDirection(Wall wall)
            {
                LocationCurve locationCurve = wall.Location as LocationCurve;
                XYZ exteriorDirection = XYZ.BasisZ;

                if (locationCurve != null)
                {
                    Curve curve = locationCurve.Curve;

                    //Write("Wall line endpoints: ", curve);

                    XYZ direction = XYZ.BasisX;
                    if (curve is Line)
                    {
                        // Obtains the tangent vector of the wall.
                        direction = curve.ComputeDerivatives(0, true).BasisX.Normalize();
                    }
                    else
                    {
                        // An assumption, for non-linear walls, that the "tangent vector" is the direction
                        // from the start of the wall to the end.
                        direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
                    }

                    // Calculate the normal vector via cross product.
                    exteriorDirection = XYZ.BasisZ.CrossProduct(direction);

                    // Flipped walls need to reverse the calculated direction
                    if (wall.Flipped) exteriorDirection = -exteriorDirection;
                }
                return exteriorDirection;
            }

            // Obtain the active project location's position and transform
            XYZ TransformByProjectLocation(XYZ direction)
            {
                // Obtain the active project location's position.
                ProjectPosition position = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
                // Construct a rotation transform from the position angle.
                /* If I cared about transforming points as well as vectors,
                 I would need to concatenate two different transformations:
                    //Obtain a rotation transform for the angle about true north
                    Transform rotationTransform = Transform.get_Rotation(
                      XYZ.Zero, XYZ.BasisZ, pna );

                    //Obtain a translation vector for the offsets
                    XYZ translationVector = new XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation);

                    Transform translationTransform = Transform.CreateTranslation(translationVector);

                    //Combine the transforms into one.
                    Transform finalTransform = translationTransform.Multiply(rotationTransform);
                */
                Transform transform = Transform.CreateRotation(XYZ.BasisZ, position.Angle);
                // Rotate the input direction by the transform
                XYZ rotatedDirection = transform.OfVector(direction);
                return rotatedDirection;
            }

            //Checking for South Facing Direction
            bool IsSouthFacing(XYZ direction)
            {
                double angleToSouth = direction.AngleTo(-XYZ.BasisY);

                return Math.Abs(angleToSouth) < Math.PI / 8;
            }

            //Checking for North Facing Direction
            bool IsNorthFacing(XYZ direction)
            {
                double angleToNorth = direction.AngleTo(XYZ.BasisY);

                return Math.Abs(angleToNorth) < Math.PI / 8;
            }

            //Checking for West Facing Direction
            bool IsWestFacing(XYZ direction)
            {
                double angleToWest = direction.AngleTo(-XYZ.BasisX);

                return Math.Abs(angleToWest) < Math.PI / 8;
            }

            //Checking for East Facing Direction
            bool IsEastFacing(XYZ direction)
            {
                double angleToEast = direction.AngleTo(XYZ.BasisX);

                return Math.Abs(angleToEast) < Math.PI / 8;
            }

            //Checking for SouthEast Facing Direction
            bool IsSouthEastFacing(XYZ direction)
            {
                double angleToEast = direction.AngleTo(XYZ.BasisX);
                double angleToSouth = direction.AngleTo(-XYZ.BasisY);

                return ((Math.Abs(angleToEast) > Math.PI / 8) & (Math.Abs(angleToEast) < (Math.PI * 3 / 8))) &
                       ((Math.Abs(angleToSouth) > Math.PI / 8) & (Math.Abs(angleToSouth) < (Math.PI * 3 / 8)));
            }

            //Checking for SouthWest Facing Direction
            bool IsSouthWestFacing(XYZ direction)
            {
                double angleToWest = direction.AngleTo(-XYZ.BasisX);
                double angleToSouth = direction.AngleTo(-XYZ.BasisY);

                return ((Math.Abs(angleToWest) > Math.PI / 8) & (Math.Abs(angleToWest) < (Math.PI * 3 / 8))) &
                       ((Math.Abs(angleToSouth) > Math.PI / 8) & (Math.Abs(angleToSouth) < (Math.PI * 3 / 8)));
            }

            //Checking for NorthEast Facing Direction
            bool IsNorthEastFacing(XYZ direction)
            {
                double angleToEast = direction.AngleTo(XYZ.BasisX);
                double angleToNorth = direction.AngleTo(XYZ.BasisY);

                return ((Math.Abs(angleToEast) > Math.PI / 8) & (Math.Abs(angleToEast) < (Math.PI * 3 / 8))) &
                       ((Math.Abs(angleToNorth) > Math.PI / 8) & (Math.Abs(angleToNorth) < (Math.PI * 3 / 8)));
            }

            //Checking for NorthWest Facing Direction
            bool IsNorthWestFacing(XYZ direction)
            {
                double angleToWest = direction.AngleTo(-XYZ.BasisX);
                double angleToNorth = direction.AngleTo(XYZ.BasisY);

                return ((Math.Abs(angleToWest) > Math.PI / 8) & (Math.Abs(angleToWest) < (Math.PI * 3 / 8))) &
                       ((Math.Abs(angleToNorth) > Math.PI / 8) & (Math.Abs(angleToNorth) < (Math.PI * 3 / 8)));
            }

            // Select Space Based on Ventilation==> AC 
            FilteredElementCollector spCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter spFilter = new ElementCategoryFilter(BuiltInCategory.OST_Rooms);
            IList<Element> spList = spCollector.WherePasses(spFilter).WhereElementIsNotElementType().ToElements();
            spCount = spList.Count;

            IList<Element> ACSPList = spList.Where(SP => SP.LookupParameter("AC_Space")
                .AsInteger() == 1).Cast<Element>().ToList();

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(30);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////


            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(40);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////MAIN 1 for Walls ==> Find Wall Adjacency from AC Spaces
            foreach (Room room in ACSPList)
            {
                lengthLst = new List<double>();
                lengthLst1 = new List<double>();
                neighbourLst = new List<Element>();
                NewWallLst = new List<Element>();
                NewWallLst1 = new List<Element>();
                wallLengthLst = new List<double>();

                // Initialise the AC Room's parameters to 0mm
                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WallTypeLst_N").Set("");
                    room.LookupParameter("WallTypeLst_NE").Set("");
                    room.LookupParameter("WallTypeLst_E").Set("");
                    room.LookupParameter("WallTypeLst_SE").Set("");
                    room.LookupParameter("WallTypeLst_S").Set("");
                    room.LookupParameter("WallTypeLst_SW").Set("");
                    room.LookupParameter("WallTypeLst_W").Set("");
                    room.LookupParameter("WallTypeLst_NW").Set("");

                    room.LookupParameter("WallTypeLst1_N").Set("");
                    room.LookupParameter("WallTypeLst1_S").Set("");
                    room.LookupParameter("WallTypeLst1_E").Set("");
                    room.LookupParameter("WallTypeLst1_W").Set("");
                    room.LookupParameter("WallTypeLst1_NE").Set("");
                    room.LookupParameter("WallTypeLst1_NW").Set("");
                    room.LookupParameter("WallTypeLst1_SE").Set("");
                    room.LookupParameter("WallTypeLst1_SW").Set("");

                    room.LookupParameter("WallAreaLst_N").Set("");
                    room.LookupParameter("WallAreaLst_S").Set("");
                    room.LookupParameter("WallAreaLst_E").Set("");
                    room.LookupParameter("WallAreaLst_W").Set("");
                    room.LookupParameter("WallAreaLst_NE").Set("");
                    room.LookupParameter("WallAreaLst_NW").Set("");
                    room.LookupParameter("WallAreaLst_SE").Set("");
                    room.LookupParameter("WallAreaLst_SW").Set("");

                    room.LookupParameter("WallUValueLst_N").Set("");
                    room.LookupParameter("WallUValueLst_S").Set("");
                    room.LookupParameter("WallUValueLst_E").Set("");
                    room.LookupParameter("WallUValueLst_W").Set("");
                    room.LookupParameter("WallUValueLst_NE").Set("");
                    room.LookupParameter("WallUValueLst_NW").Set("");
                    room.LookupParameter("WallUValueLst_SE").Set("");
                    room.LookupParameter("WallUValueLst_SW").Set("");


                    //////////////////////////////////////////

                    room.LookupParameter("WndwTypeLst_N").Set("");
                    room.LookupParameter("WndwTypeLst_NE").Set("");
                    room.LookupParameter("WndwTypeLst_E").Set("");
                    room.LookupParameter("WndwTypeLst_SE").Set("");
                    room.LookupParameter("WndwTypeLst_S").Set("");
                    room.LookupParameter("WndwTypeLst_SW").Set("");
                    room.LookupParameter("WndwTypeLst_W").Set("");
                    room.LookupParameter("WndwTypeLst_NW").Set("");

                    room.LookupParameter("WndwTypeLst1_N").Set("");
                    room.LookupParameter("WndwTypeLst1_S").Set("");
                    room.LookupParameter("WndwTypeLst1_E").Set("");
                    room.LookupParameter("WndwTypeLst1_W").Set("");
                    room.LookupParameter("WndwTypeLst1_NE").Set("");
                    room.LookupParameter("WndwTypeLst1_NW").Set("");
                    room.LookupParameter("WndwTypeLst1_SE").Set("");
                    room.LookupParameter("WndwTypeLst1_SW").Set("");

                    room.LookupParameter("WndwAreaLst_N").Set("");
                    room.LookupParameter("WndwAreaLst_S").Set("");
                    room.LookupParameter("WndwAreaLst_E").Set("");
                    room.LookupParameter("WndwAreaLst_W").Set("");
                    room.LookupParameter("WndwAreaLst_NE").Set("");
                    room.LookupParameter("WndwAreaLst_NW").Set("");
                    room.LookupParameter("WndwAreaLst_SE").Set("");
                    room.LookupParameter("WndwAreaLst_SW").Set("");

                    room.LookupParameter("WndwUValueLst_N").Set("");
                    room.LookupParameter("WndwUValueLst_S").Set("");
                    room.LookupParameter("WndwUValueLst_E").Set("");
                    room.LookupParameter("WndwUValueLst_W").Set("");
                    room.LookupParameter("WndwUValueLst_NE").Set("");
                    room.LookupParameter("WndwUValueLst_NW").Set("");
                    room.LookupParameter("WndwUValueLst_SE").Set("");
                    room.LookupParameter("WndwUValueLst_SW").Set("");

                    room.LookupParameter("WndwSC1ValueLst_N").Set("");
                    room.LookupParameter("WndwSC1ValueLst_S").Set("");
                    room.LookupParameter("WndwSC1ValueLst_E").Set("");
                    room.LookupParameter("WndwSC1ValueLst_W").Set("");
                    room.LookupParameter("WndwSC1ValueLst_NE").Set("");
                    room.LookupParameter("WndwSC1ValueLst_NW").Set("");
                    room.LookupParameter("WndwSC1ValueLst_SE").Set("");
                    room.LookupParameter("WndwSC1ValueLst_SW").Set("");

                    trans.Commit();
                }
                DetermineAdjacentElementLengthsAndWallAreas(room);
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(60);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(70);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////MAIN for Windows ==> Check Any External Window for AC Spaces
            foreach (Room room in ACSPList)
            {
                WndwList = new List<Element>();

                //Creating bounding box from room
                BoundingBoxXYZ bb = room.get_BoundingBox(null);
                Outline outline = new Outline(bb.Min,bb.Max);
                BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);

                //Clash check room bb with windows
                FilteredElementCollector bbFltElemCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .WhereElementIsNotElementType()
                    .WhereElementIsViewIndependent()
                    .OfClass(typeof(FamilyInstance));

                //Create window list
                if(room.LookupParameter("AC_Space").AsInteger()==1)
                {
                foreach (Element Wndw in bbFltElemCollector)
                {
                    //TaskDialog.Show("A",$"{room.Name}~{Wndw.LookupParameter("ETTV_Room").AsValueString()}");
                    if (Wndw.LookupParameter("ETTV_Room").AsValueString() == $"{room.Name}")
                        WndwList.Add(Wndw);
                }
                }
                //Runns windows, it attaches the window to the room
                GetWindows makesWin_Room = new GetWindows();
                Result GetWindowResult = makesWin_Room.Execute(commandData,ref message,elements);
                if (GetWindowResult == Result.Succeeded)
                {
                    //TaskDialog.Show("Success","GetWindows executed successfully.");
                }
                else
                {
                    TaskDialog.Show("Error","GetWindows failed to execute.");
                }

                //Separate window list based on orientation 
                //foreach (Element Wndw in WndwList)
                //{
                //    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                //    {
                //        trans.Start();
                //        Wndw.LookupParameter("ETTV_Room").Set(room.Name);
                //        trans.Commit();
                //    }
                //}

                //Initialise parameters for window orientation 
                N_WndwList = new List<Element>();
                N_Wndw_Id_List = new List<ElementId>();
                S_WndwList = new List<Element>();
                S_Wndw_Id_List = new List<ElementId>();
                E_WndwList = new List<Element>();
                E_Wndw_Id_List = new List<ElementId>();
                W_WndwList = new List<Element>();
                W_Wndw_Id_List = new List<ElementId>();
                NE_WndwList = new List<Element>();
                NE_Wndw_Id_List = new List<ElementId>();
                NW_WndwList = new List<Element>();
                NW_Wndw_Id_List = new List<ElementId>();
                SE_WndwList = new List<Element>();
                SE_Wndw_Id_List = new List<ElementId>();
                SW_WndwList = new List<Element>();
                SW_Wndw_Id_List = new List<ElementId>();

                //separate windows according to orientations
                foreach (Element Wndw in WndwList)
                {
                    if (Wndw.LookupParameter("North").AsInteger() == 1)
                    {
                        N_WndwList.Add(Wndw);
                        N_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("South").AsInteger() == 1)
                    {
                        S_WndwList.Add(Wndw);
                        S_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("East").AsInteger() == 1)
                    {
                        E_WndwList.Add(Wndw);
                        E_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("West").AsInteger() == 1)
                    {
                        W_WndwList.Add(Wndw);
                        W_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("NorthEast").AsInteger() == 1)
                    {
                        NE_WndwList.Add(Wndw);
                        NE_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("NorthWest").AsInteger() == 1)
                    {
                        NW_WndwList.Add(Wndw);
                        NW_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("SouthEast").AsInteger() == 1)
                    {
                        SE_WndwList.Add(Wndw);
                        SE_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }

                    else if (Wndw.LookupParameter("SouthWest").AsInteger() == 1)
                    {
                        SW_WndwList.Add(Wndw);
                        SW_Wndw_Id_List.Add(Wndw.GetTypeId());
                    }
                }

                /////////////////////////////////////////////////////////////////////////////////////
                UpdateProgress(80);
                System.Threading.Thread.Sleep(500);
                /////////////////////////////////////////////////////////////////////////////////////

                ////////////////////////////////// North Window /////////////////////////////////////
                N_Wndw_Strg = "";
                N_Wndw_Strg1 = "";
                N_Wndw_Area = "";
                N_Wndw_U = "";
                N_Wndw_SC1 = "";
                foreach (Element NWndw in N_WndwList)
                {

                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(NWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (NWndw.GetTypeId() == WNId)
                        {
                            N_Wndw_Strg1 = N_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            N_Wndw_Strg = N_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                          "\r\n";
                            N_Wndw_Area = N_Wndw_Area + WndArea + "\r\n";
                            N_Wndw_U = N_Wndw_U + WndUValue + "\r\n";
                            N_Wndw_SC1 = N_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_N").Set(N_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_N").Set(N_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_N").Set(N_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_N").Set(N_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_N").Set(N_Wndw_SC1);

                    trans.Commit();
                }

                ////////////////////////////////// North Window /////////////////////////////////////

                ////////////////////////////////// South Window /////////////////////////////////////
                S_Wndw_Strg = "";
                S_Wndw_Strg1 = "";
                S_Wndw_Area = "";
                S_Wndw_U = "";
                S_Wndw_SC1 = "";
                foreach (Element SWndw in S_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(SWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (SWndw.GetTypeId() == WNId)
                        {
                            S_Wndw_Strg1 = S_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            S_Wndw_Strg = S_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                          "\r\n";
                            S_Wndw_Area = S_Wndw_Area + WndArea + "\r\n";
                            S_Wndw_U = S_Wndw_U + WndUValue + "\r\n";
                            S_Wndw_SC1 = S_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_S").Set(S_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_S").Set(S_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_S").Set(S_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_S").Set(S_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_S").Set(S_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// South Window /////////////////////////////////////

                ////////////////////////////////// East Window /////////////////////////////////////
                E_Wndw_Strg = "";
                E_Wndw_Strg1 = "";
                E_Wndw_Area = "";
                E_Wndw_U = "";
                E_Wndw_SC1 = "";
                foreach (Element EWndw in E_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(EWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (EWndw.GetTypeId() == WNId)
                        {
                            E_Wndw_Strg1 = E_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            E_Wndw_Strg = E_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                          "\r\n";
                            E_Wndw_Area = E_Wndw_Area + WndArea + "\r\n";
                            E_Wndw_U = E_Wndw_U + WndUValue + "\r\n";
                            E_Wndw_SC1 = E_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_E").Set(E_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_E").Set(E_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_E").Set(E_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_E").Set(E_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_E").Set(E_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// East Window /////////////////////////////////////

                ////////////////////////////////// West Window /////////////////////////////////////
                W_Wndw_Strg = "";
                W_Wndw_Strg1 = "";
                W_Wndw_Area = "";
                W_Wndw_U = "";
                W_Wndw_SC1 = "";
                foreach (Element WWndw in W_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(WWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (WWndw.GetTypeId() == WNId)
                        {
                            W_Wndw_Strg1 = W_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            W_Wndw_Strg = W_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                          "\r\n";
                            W_Wndw_Area = W_Wndw_Area + WndArea + "\r\n";
                            W_Wndw_U = W_Wndw_U + WndUValue + "\r\n";
                            W_Wndw_SC1 = W_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_W").Set(W_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_W").Set(W_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_W").Set(W_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_W").Set(W_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_W").Set(W_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// West Window /////////////////////////////////////

                ////////////////////////////////// NorthEast Window /////////////////////////////////////
                NE_Wndw_Strg = "";
                NE_Wndw_Strg1 = "";
                NE_Wndw_Area = "";
                NE_Wndw_U = "";
                NE_Wndw_SC1 = "";
                foreach (Element NEWndw in NE_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(NEWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (NEWndw.GetTypeId() == WNId)
                        {
                            NE_Wndw_Strg1 = NE_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            NE_Wndw_Strg = NE_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                           "\r\n";
                            NE_Wndw_Area = NE_Wndw_Area + WndArea + "\r\n";
                            NE_Wndw_U = NE_Wndw_U + WndUValue + "\r\n";
                            NE_Wndw_SC1 = NE_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_NE").Set(NE_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_NE").Set(NE_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_NE").Set(NE_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_NE").Set(NE_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_NE").Set(NE_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// NorthEast Window /////////////////////////////////////

                ////////////////////////////////// NorthWest Window /////////////////////////////////////
                NW_Wndw_Strg = "";
                NW_Wndw_Strg1 = "";
                NW_Wndw_Area = "";
                NW_Wndw_U = "";
                NW_Wndw_SC1 = "";
                foreach (Element NWWndw in NW_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(NWWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (NWWndw.GetTypeId() == WNId)
                        {
                            NW_Wndw_Strg1 = NW_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            NW_Wndw_Strg = NW_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                           "\r\n";
                            NW_Wndw_Area = NW_Wndw_Area + WndArea + "\r\n";
                            NW_Wndw_U = NW_Wndw_U + WndUValue + "\r\n";
                            NW_Wndw_SC1 = NW_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_NW").Set(NW_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_NW").Set(NW_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_NW").Set(NW_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_NW").Set(NW_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_NW").Set(NW_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// NorthWest Window /////////////////////////////////////

                ////////////////////////////////// SouthEast Window /////////////////////////////////////
                SE_Wndw_Strg = "";
                SE_Wndw_Strg1 = "";
                SE_Wndw_Area = "";
                SE_Wndw_U = "";
                SE_Wndw_SC1 = "";
                foreach (Element SEWndw in SE_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(SEWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////


                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (SEWndw.GetTypeId() == WNId)
                        {
                            SE_Wndw_Strg1 = SE_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            SE_Wndw_Strg = SE_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                           "\r\n";
                            SE_Wndw_Area = SE_Wndw_Area + WndArea + "\r\n";
                            SE_Wndw_U = SE_Wndw_U + WndUValue + "\r\n";
                            SE_Wndw_SC1 = SE_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_SE").Set(SE_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_SE").Set(SE_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_SE").Set(SE_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_SE").Set(SE_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_SE").Set(SE_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// SouthEast Window /////////////////////////////////////

                ////////////////////////////////// SouthWest Window /////////////////////////////////////
                SW_Wndw_Strg = "";
                SW_Wndw_Strg1 = "";
                SW_Wndw_Area = "";
                SW_Wndw_U = "";
                SW_Wndw_SC1 = "";
                
                foreach (Element SWWndw in SW_WndwList)
                {
                    //////to get H/W/Area/U values and SC values//////
                    ElementType WndType = doc.GetElement(SWWndw.GetTypeId()) as ElementType;
                    Parameter h = WndType.LookupParameter("Height");
                    double WndHgt = (h.AsDouble() * _footToMm) / 1000;
                    Parameter w = WndType.LookupParameter("Width");
                    double WndWdt = (w.AsDouble() * _footToMm) / 1000;
                    double WndArea = WndWdt * WndHgt;
                    Parameter U = WndType.LookupParameter("Heat Transfer Coefficient (U)");
                    if (U == null)
                    {
                        continue;
                    }

                    double R = 1 / U.AsDouble();
                    R += 0.044 + 0.12;
                    double WndUValue = 1 / R;
                    Parameter SC = WndType.LookupParameter("Solar Heat Gain Coefficient");
                    double WndSCValue = (SC.AsDouble()) / 0.87;
                    //////to get H/W/Area/U values and SC values//////

                    num3 = 0;
                    foreach (ElementId WNId in WndwTypeId_2)
                    {
                        if (SWWndw.GetTypeId() == WNId)
                        {
                            SW_Wndw_Strg1 = SW_Wndw_Strg1 + "F" + ((num3 + 1).ToString()) + "\r\n";
                            SW_Wndw_Strg = SW_Wndw_Strg + WndwTypeLst[num3].FamilyName + " " + WndwTypeLst[num3].Name +
                                           "\r\n";
                            SW_Wndw_Area = SW_Wndw_Area + WndArea + "\r\n";
                            SW_Wndw_U = SW_Wndw_U + WndUValue + "\r\n";
                            SW_Wndw_SC1 = SW_Wndw_SC1 + WndSCValue.ToString("0.00") + "\r\n";
                        }
                        num3 = num3 + 1;
                    }
                }

                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                {
                    trans.Start();

                    room.LookupParameter("WndwTypeLst1_SW").Set(SW_Wndw_Strg1);
                    room.LookupParameter("WndwTypeLst_SW").Set(SW_Wndw_Strg);
                    room.LookupParameter("WndwAreaLst_SW").Set(SW_Wndw_Area);
                    room.LookupParameter("WndwUValueLst_SW").Set(SW_Wndw_U);
                    room.LookupParameter("WndwSC1ValueLst_SW").Set(SW_Wndw_SC1);

                    trans.Commit();
                }
                ////////////////////////////////// SouthWest Window /////////////////////////////////////
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(90);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////MAIN 2 for Walls ==> For Walls Processing from AC Rooms
            void DetermineAdjacentElementLengthsAndWallAreas(Room room)
            {
                using (TransactionGroup transGroup = new TransactionGroup(doc))
                {
                    transGroup.Start("Transaction Group");

                    //////////////TO DELETE LOOSE WALLS//////////////////////////////// 

                    #region Delet Loose Walls

                    //bool wallsInsideRoomExist = true;

                    //while (wallsInsideRoomExist)
                    //{
                    //    List<Wall> intersects_wall = new List<Wall>();
                    //    BoundingBoxXYZ bb_2 = room.get_BoundingBox(null);
                    //    Outline outline_2 = new Outline(bb_2.Min, bb_2.Max);
                    //    BoundingBoxIntersectsFilter bbfilter_2 = new BoundingBoxIntersectsFilter(outline_2);

                    //    FilteredElementCollector bbFltElemCollector_2 = new FilteredElementCollector(doc);
                    //    ICollection<ElementId> intersects_2 = bbFltElemCollector_2.OfCategory(BuiltInCategory.OST_Walls)
                    //        .WhereElementIsNotElementType()
                    //        .WhereElementIsViewIndependent()
                    //        .OfClass(typeof(Wall))
                    //        .WherePasses(bbfilter_2).ToElementIds();

                    //    BoundingBoxXYZ roomBoundary = room.get_BoundingBox(null);

                    //    if (roomBoundary == null)
                    //    {
                    //        TaskDialog.Show("Error", "Room boundary not found.");
                    //        return;
                    //    }

                    //    List<Wall> walls = new List<Wall>();
                    //    foreach (ElementId wallId in intersects_2)
                    //    {
                    //        Element wallElement = doc.GetElement(wallId);
                    //        if (wallElement != null && wallElement is Wall)
                    //        {
                    //            walls.Add((Wall)wallElement);
                    //        }
                    //    }

                    //    foreach (Wall wall in walls)
                    //    {
                    //        if (doc.GetElement(wall.Id) == null)
                    //        {
                    //            continue;
                    //        }

                    //        LocationCurve wallLocation = wall.Location as LocationCurve;
                    //        if (wallLocation != null)
                    //        {
                    //            Curve wallCurve = wallLocation.Curve;
                    //            XYZ wallStart = wallCurve.GetEndPoint(0);
                    //            XYZ wallEnd = wallCurve.GetEndPoint(1);
                    //            XYZ wallMidpoint = (wallStart + wallEnd) / 2;

                    //            bool isWallStartInRoom = room.IsPointInRoom(wallStart);
                    //            bool isWallEndInRoom = room.IsPointInRoom(wallEnd);
                    //            bool isWallMidpointInRoom = room.IsPointInRoom(wallMidpoint);

                    //            if (isWallStartInRoom || isWallEndInRoom || isWallMidpointInRoom)
                    //            {
                    //                intersects_wall.Add(wall);
                    //            }
                    //        }
                    //    }

                    //    if (intersects_wall.Count == 0)
                    //    {
                    //        wallsInsideRoomExist = false;
                    //    }
                    //    else
                    //    {
                    //        List<ElementId> intersects_wall_ids = intersects_wall.Select(wall => wall.Id).ToList();

                    //        using (Transaction trans2 = new Transaction(doc, "ETTV_1"))
                    //        {
                    //            trans2.Start();
                    //            doc.Delete(intersects_wall_ids);
                    //            trans2.Commit();
                    //        }
                    //    }
                    //}

                    #endregion

                    //////////////LOOSE WALLS DELETED////////////////////////////////

                    ///////////////////////////////////// NEW CODES START HERE ///////////////////////////////////////////////                                     
                    ///////////////////////////////////// NEW CODES END HERE ///////////////////////////////////////////////

                    #region Delet Column

                    //////////////TO DELETE COLUMNS////////////////////////////////
                    //using (Transaction trans1 = new Transaction(doc, "Delete Columns"))
                    //{
                    //    trans1.Start();

                    //    //Creating bounding box from room
                    //    BoundingBoxXYZ bb_1 = room.get_BoundingBox(null);
                    //    Outline outline_1 = new Outline(bb_1.Min, bb_1.Max);
                    //    BoundingBoxIntersectsFilter bbfilter_1 = new BoundingBoxIntersectsFilter(outline_1);

                    //    //Clash check room bb_1 with columns
                    //    FilteredElementCollector bbFltElemCollector_1 = new FilteredElementCollector(doc);
                    //    intersects_1 = bbFltElemCollector_1.OfCategory(BuiltInCategory.OST_Columns)
                    //        .WhereElementIsNotElementType()
                    //        .WhereElementIsViewIndependent()
                    //        .OfClass(typeof(FamilyInstance))
                    //        .WherePasses(bbfilter_1).ToElementIds();

                    //    doc.Delete(intersects_1);

                    //    trans1.Commit();
                    //}
                    //////////////COLUMNS DELETED////////////////////////////////

                    #endregion

                    transGroup.RollBack(); //to restore back the columns
                }

                //////////////////////////////////Original Code Starts////////////////////////////////////////////             

                boundaries_0 = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

                //boundaries_1 = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

                boundaries_1 = new List<List<BoundarySegment>>();
                BLst = new List<BoundarySegment>();

                foreach (List<BoundarySegment> boundarySegmentList in boundaries_0)
                {
                    foreach (BoundarySegment boundarySegment in boundarySegmentList)
                    {
                        // Get the curve of the boundary segment
                        Curve curve = boundarySegment.GetCurve();

                        // Use the curve to find the associated element
                        Element element = room.Document.GetElement(boundarySegment.ElementId);

                        if (element is Wall)
                        {
                            BLst.Add(boundarySegment);

                            // This boundary segment is from a room
                            // You can further analyze the room properties if needed
                            // For example: string roomFunction = ((Room)element).get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                        }
                    }
                }

                #region Test Boundary Segment1

                //A = 0;
                //B = 0;
                //C = 0;                
                //foreach (var boundarySegmentList in boundaries_0)
                //{
                //    foreach (var boundarySegment in boundarySegmentList)
                //    {
                //        // Get the curve of the boundary segment
                //        Curve curve = boundarySegment.GetCurve();

                //        // Use the curve to find the associated element
                //        Element element = room.Document.GetElement(boundarySegment.ElementId);

                //        if (element is Wall)
                //        {
                //            boundaries_0.add(element);
                //            // This boundary segment is from a room
                //            // You can further analyze the room properties if needed
                //            // For example: string roomFunction = ((Room)element).get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                //        }
                //        else if (element is FamilyInstance && ((FamilyInstance)element).Symbol.FamilyName == "Rectangular Column")
                //        {
                //            B = B + 1;
                //            // This boundary segment is from a column
                //            // You can further analyze the column properties if needed
                //            // For example: string columnType = ((FamilyInstance)element).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString();
                //        }
                //        else
                //        {
                //            C = C + 1;
                //            // The boundary segment is associated with another type of element
                //        }
                //    }
                //}

                #endregion

                boundaries_1.Add(BLst);
                int n = boundaries_1.Count; /// separated from Original
                iBoundary = 0; /// separated from Original    
                
                //////////////////////////////////////////////////////////////////////////////////////////
                UpdateProgress(50);
                System.Threading.Thread.Sleep(500);
                //////////////////////////////////////////////////////////////////////////////////////////

                foreach (List<BoundarySegment> b in boundaries_1)
                {

                    ++iBoundary;
                    iSegment = 0;
                    foreach (BoundarySegment s in b)
                    {
                        ++iSegment;
                        var element = doc.GetElement(s.ElementId);
                        if (element != null && element is Wall) // To get the length  of the boundary segment
                        {
                            neighbour = doc.GetElement(s.ElementId);
                            neighbourLst.Add(neighbour);

                            curve = s.GetCurve();
                            length = curve.Length;
                            lengthLst.Add(length);
                        }
                    }

                    foreach (Element e in neighbourLst)
                    {
                        Wall wall = e as Wall;

                        Parameter p = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                        area = p.AsDouble();

                        LocationCurve lc = wall.Location as LocationCurve;
                        wallLength = lc.Curve.Length;

                        NewWallLst.Add(wall);
                        wallLengthLst.Add(wallLength);
                    }
                    
                    num1 = 0;
                    //inserts = null;
                    N_Wall_List = new List<Element>();
                    N_Wall_Id_List = new List<ElementId>();
                    N_Wall_Lgt_List = new List<double>();
                    N_Wall_InstArea_List = new List<double>();

                    S_Wall_List = new List<Element>();
                    S_Wall_Id_List = new List<ElementId>();
                    S_Wall_Lgt_List = new List<double>();
                    S_Wall_InstArea_List = new List<double>();

                    E_Wall_List = new List<Element>();
                    E_Wall_Id_List = new List<ElementId>();
                    E_Wall_Lgt_List = new List<double>();
                    E_Wall_InstArea_List = new List<double>();

                    W_Wall_List = new List<Element>();
                    W_Wall_Id_List = new List<ElementId>();
                    W_Wall_Lgt_List = new List<double>();
                    W_Wall_InstArea_List = new List<double>();

                    NE_Wall_List = new List<Element>();
                    NE_Wall_Id_List = new List<ElementId>();
                    NE_Wall_Lgt_List = new List<double>();
                    NE_Wall_InstArea_List = new List<double>();

                    NW_Wall_List = new List<Element>();
                    NW_Wall_Id_List = new List<ElementId>();
                    NW_Wall_Lgt_List = new List<double>();
                    NW_Wall_InstArea_List = new List<double>();

                    SE_Wall_List = new List<Element>();
                    SE_Wall_Id_List = new List<ElementId>();
                    SE_Wall_Lgt_List = new List<double>();
                    SE_Wall_InstArea_List = new List<double>();

                    SW_Wall_List = new List<Element>();
                    SW_Wall_Id_List = new List<ElementId>();
                    SW_Wall_Lgt_List = new List<double>();
                    SW_Wall_InstArea_List = new List<double>();

                    /////////// Assigning the information of walls to Room's orentation related parameters /////////////
                    
                    ////////// Check the adjacent wall list with external wall list and assign to different lists based on orientations///////
                    foreach (Element e in NewWallLst)
                    {
                        foreach (Element w in ExtWallLst)
                        {
                            if (e.Id == w.Id)
                            {
                                //num2 = 0;
                                lengthLst1.Add(lengthLst[num1]);
                                NewWallLst1.Add(w);
                                ExtWallLst1.Add(w);
                                ids_3.Add(w.Id); // For Selection of AC External Walls for checking purpose

                                ///////////////////////////to calculate the insert areas ////////////////////////////
                                inserts = (e as HostObject).FindInserts(true, true, true, true);
                                InsertLst = new List<Element>();
                                InsertLst1 = new List<Element>();
                                TotalInstArea = 0;
                                foreach (ElementId emtid in inserts)
                                {
                                    Element emt = doc.GetElement(emtid);
                                    InsertLst.Add(emt);
                                }

                                WndwList1 = new List<Element>();
                                //Creating bounding box from room
                                BoundingBoxXYZ bb = room.get_BoundingBox(null);
                                Outline outline = new Outline(bb.Min, bb.Max);
                                BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);

                                //Clash check room bb with windows
                                FilteredElementCollector bbFltElemCollector = new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_Windows)
                                    .WhereElementIsNotElementType()
                                    .WhereElementIsViewIndependent()
                                    .OfClass(typeof(FamilyInstance))
                                    .WherePasses(bbfilter);

                                //Create window list that clash with the room
                                foreach (Element Wndw in bbFltElemCollector)
                                {
                                    WndwList1.Add(Wndw);
                                }

                                foreach (Element elmt in WndwList1)
                                {
                                    foreach (Element inst in InsertLst)
                                    {
                                        if (elmt.Id == inst.Id)
                                        {
                                            InsertLst1.Add(inst);
                                        }
                                    }
                                }

                                foreach (Element emt in InsertLst1)
                                {
                                    ElementType WndType = doc.GetElement(emt.GetTypeId()) as ElementType;
                                    Parameter ht = WndType.LookupParameter("Height");
                                    double WndHgt = (ht.AsDouble() * _footToMm) / 1000;
                                    Parameter wd = WndType.LookupParameter("Width");
                                    double WndWdt = (wd.AsDouble() * _footToMm) / 1000;
                                    double WndArea = WndWdt * WndHgt;

                                    TotalInstArea = TotalInstArea + WndArea;
                                }
                                ///////////////////////////to calculate the insert areas ////////////////////////////


                                /////// Check the new walls (adjacent walls), create the lists of these new walls according to the orientation ///////////
                                using (Transaction trans = new Transaction(doc, "ETTV_1"))
                                {
                                    trans.Start();

                                    if (e.LookupParameter("North").AsInteger() == 1)
                                    {
                                        N_Wall_List.Add(e);
                                        N_Wall_Id_List.Add(e.GetTypeId());
                                        N_Wall_Lgt_List.Add(lengthLst[num1]);
                                        N_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("NorthEast").AsInteger() == 1)
                                    {
                                        NE_Wall_List.Add(e);
                                        NE_Wall_Id_List.Add(e.GetTypeId());
                                        NE_Wall_Lgt_List.Add(lengthLst[num1]);
                                        NE_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("East").AsInteger() == 1)
                                    {
                                        E_Wall_List.Add(e);
                                        E_Wall_Id_List.Add(e.GetTypeId());
                                        E_Wall_Lgt_List.Add(lengthLst[num1]);
                                        E_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("SouthEast").AsInteger() == 1)
                                    {
                                        SE_Wall_List.Add(e);
                                        SE_Wall_Id_List.Add(e.GetTypeId());
                                        SE_Wall_Lgt_List.Add(lengthLst[num1]);
                                        SE_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("South").AsInteger() == 1)
                                    {
                                        S_Wall_List.Add(e);
                                        S_Wall_Id_List.Add(e.GetTypeId());
                                        S_Wall_Lgt_List.Add(lengthLst[num1]);
                                        S_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("SouthWest").AsInteger() == 1)
                                    {
                                        SW_Wall_List.Add(e);
                                        SW_Wall_Id_List.Add(e.GetTypeId());
                                        SW_Wall_Lgt_List.Add(lengthLst[num1]);
                                        SW_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("West").AsInteger() == 1)
                                    {
                                        W_Wall_List.Add(e);
                                        W_Wall_Id_List.Add(e.GetTypeId());
                                        W_Wall_Lgt_List.Add(lengthLst[num1]);
                                        W_Wall_InstArea_List.Add(TotalInstArea);
                                    }

                                    else if (e.LookupParameter("NorthWest").AsInteger() == 1)
                                    {
                                        NW_Wall_List.Add(e);
                                        NW_Wall_Id_List.Add(e.GetTypeId());
                                        NW_Wall_Lgt_List.Add(lengthLst[num1]);
                                        NW_Wall_InstArea_List.Add(TotalInstArea);
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        num1 = num1 + 1;
                    }

                    ////////Enter wall length (which is retrieved from boundary segment length) into the AC Room's parameters////////
                    
                    ////////////////////////////////// North Wall /////////////////////////////////////
                    N_Wall_Strg = "";
                    N_Wall_Strg1 = "";
                    N_Wall_Area = "";
                    N_Wall_U = "";

                    num4 = 0;
                    foreach (Element NWall in N_Wall_List)
                    {

                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(NWall.GetTypeId()) as ElementType;

                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values//////                  


                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (NWall.GetTypeId() == WallId)
                            {
                                N_Wall_Strg1 = N_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                N_Wall_Strg = N_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                              WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(N_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (N_Wall_InstArea_List[num4]);
                                N_Wall_Area = N_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                N_Wall_U = N_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";
                                
                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_N").Set(N_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_N").Set(N_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_N").Set(N_Wall_Area);
                        room.LookupParameter("WallUValueLst_N").Set(N_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// North Wall /////////////////////////////////////

                    ////////////////////////////////// South Wall /////////////////////////////////////
                    S_Wall_Strg = "";
                    S_Wall_Strg = "";
                    S_Wall_Strg1 = "";
                    S_Wall_U = "";

                    num4 = 0;
                    foreach (Element SWall in S_Wall_List)
                    {

                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(SWall.GetTypeId()) as ElementType;

                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values////// 

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (SWall.GetTypeId() == WallId)
                            {
                                S_Wall_Strg1 = S_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                S_Wall_Strg = S_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                              WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(S_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (S_Wall_InstArea_List[num4]);
                                S_Wall_Area = S_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                S_Wall_U = S_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_S").Set(S_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_S").Set(S_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_S").Set(S_Wall_Area);
                        room.LookupParameter("WallUValueLst_S").Set(S_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// South Wall /////////////////////////////////////

                    ////////////////////////////////// East Wall /////////////////////////////////////
                    E_Wall_Strg = "";
                    E_Wall_Strg1 = "";
                    E_Wall_Area = "";
                    E_Wall_U = "";

                    num4 = 0;
                    foreach (Element EWall in E_Wall_List)
                    {
                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(EWall.GetTypeId()) as ElementType;

                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values////// 

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (EWall.GetTypeId() == WallId)
                            {
                                E_Wall_Strg1 = E_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                E_Wall_Strg = E_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                              WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(E_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (E_Wall_InstArea_List[num4]);
                                E_Wall_Area = E_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                E_Wall_U = E_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_E").Set(E_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_E").Set(E_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_E").Set(E_Wall_Area);
                        room.LookupParameter("WallUValueLst_E").Set(E_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// East Wall /////////////////////////////////////

                    ////////////////////////////////// West Wall /////////////////////////////////////
                    W_Wall_Strg = "";
                    W_Wall_Strg1 = "";
                    W_Wall_Area = "";
                    W_Wall_U = "";

                    num4 = 0;
                    foreach (Element WWall in W_Wall_List)
                    {
                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(WWall.GetTypeId()) as ElementType;

                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values////// 

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (WWall.GetTypeId() == WallId)
                            {
                                W_Wall_Strg1 = W_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                W_Wall_Strg = W_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                              WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(W_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (W_Wall_InstArea_List[num4]);
                                W_Wall_Area = W_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                W_Wall_U = W_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_W").Set(W_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_W").Set(W_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_W").Set(W_Wall_Area);
                        room.LookupParameter("WallUValueLst_W").Set(W_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// West Wall /////////////////////////////////////
                    
                    ////////////////////////////////// NorthEast Wall /////////////////////////////////////
                    NE_Wall_Strg = "";
                    NE_Wall_Strg1 = "";
                    NE_Wall_Area = "";
                    NE_Wall_U = "";

                    num4 = 0;
                    foreach (Element NEWall in NE_Wall_List)
                    {
                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(NEWall.GetTypeId()) as ElementType;

                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values////// 

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (NEWall.GetTypeId() == WallId)
                            {
                                NE_Wall_Strg1 = NE_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                NE_Wall_Strg = NE_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                               WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(NE_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (NE_Wall_InstArea_List[num4]);
                                NE_Wall_Area = NE_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                NE_Wall_U = NE_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_NE").Set(NE_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_NE").Set(NE_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_NE").Set(NE_Wall_Area);
                        room.LookupParameter("WallUValueLst_NE").Set(NE_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// NorthEast Wall /////////////////////////////////////

                    ////////////////////////////////// NorthWest Wall /////////////////////////////////////
                    NW_Wall_Strg = "";
                    NW_Wall_Strg1 = "";
                    NW_Wall_Area = "";
                    NW_Wall_U = "";

                    num4 = 0;
                    foreach (Element NWWall in NW_Wall_List)
                    {
                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(NWWall.GetTypeId()) as ElementType;
                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values////// 

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (NWWall.GetTypeId() == WallId)
                            {
                                NW_Wall_Strg1 = NW_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                NW_Wall_Strg = NW_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                               WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(NW_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (NW_Wall_InstArea_List[num4]);
                                NW_Wall_Area = NW_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                NW_Wall_U = NW_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_NW").Set(NW_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_NW").Set(NW_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_NW").Set(NW_Wall_Area);
                        room.LookupParameter("WallUValueLst_NW").Set(NW_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// NorthWest Wall /////////////////////////////////////

                    ////////////////////////////////// SouthEast Wall /////////////////////////////////////
                    SE_Wall_Strg = "";
                    SE_Wall_Strg1 = "";
                    SE_Wall_Area = "";
                    SE_Wall_U = "";

                    num4 = 0;
                    foreach (Element SEWall in SE_Wall_List)
                    {
                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(SEWall.GetTypeId()) as ElementType;
                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z-rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;

                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values//////

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (SEWall.GetTypeId() == WallId)
                            {
                                SE_Wall_Strg1 = SE_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                SE_Wall_Strg = SE_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                               WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(SE_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (SE_Wall_InstArea_List[num4]);
                                SE_Wall_Area = SE_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                SE_Wall_U = SE_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_SE").Set(SE_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_SE").Set(SE_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_SE").Set(SE_Wall_Area);
                        room.LookupParameter("WallUValueLst_SE").Set(SE_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// SouthEast Wall /////////////////////////////////////

                    ////////////////////////////////// SouthWest Wall /////////////////////////////////////
                    SW_Wall_Strg = "";
                    SW_Wall_Strg1 = "";
                    SW_Wall_Area = "";
                    SW_Wall_U = "";

                    num4 = 0;
                    foreach (Element SWWall in SW_Wall_List)
                    {
                        //////to get Area/U values//////
                        ElementType WallType = doc.GetElement(SWWall.GetTypeId()) as ElementType;
                        BoundingBoxXYZ rBB = room.get_BoundingBox(null);
                        double h = rBB.Max.Z - rBB.Min.Z;
                        //Parameter h = room.LookupParameter("Unbounded Height");
                        double WallHgt = (h * _footToMm) / 1000;
                        Parameter U = WallType.LookupParameter("Heat Transfer Coefficient (U)");
                        if (U == null)
                        {
                            continue;
                        }

                        double R = 1 / U.AsDouble();
                        R += 0.044 + 0.12;
                        double WallUValue = 1 / R;
                        //////to get Area/U values//////

                        num3 = 0;
                        foreach (ElementId WallId in WallTypeId_2)
                        {
                            if (SWWall.GetTypeId() == WallId)
                            {
                                SW_Wall_Strg1 = SW_Wall_Strg1 + "W" + ((num3 + 1).ToString()) + "\r\n";
                                SW_Wall_Strg = SW_Wall_Strg + WallTypeLst[num3].FamilyName + " " +
                                               WallTypeLst[num3].Name + "\r\n";
                                double A = (WallHgt * ((Convert.ToDouble(SW_Wall_Lgt_List[num4]) * _footToMm) / 1000)) -
                                           (SW_Wall_InstArea_List[num4]);
                                SW_Wall_Area = SW_Wall_Area + (A.ToString("0.00")) + "\r\n";
                                SW_Wall_U = SW_Wall_U + (WallUValue.ToString("0.0000")) + "\r\n";

                                num4 = num4 + 1;
                            }
                            num3 = num3 + 1;
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "ETTV_1"))
                    {
                        trans.Start();

                        room.LookupParameter("WallTypeLst_SW").Set(SW_Wall_Strg);
                        room.LookupParameter("WallTypeLst1_SW").Set(SW_Wall_Strg1);
                        room.LookupParameter("WallAreaLst_SW").Set(SW_Wall_Area);
                        room.LookupParameter("WallUValueLst_SW").Set(SW_Wall_U);

                        trans.Commit();
                    }
                    ////////////////////////////////// SouthWest Wall /////////////////////////////////////

                    // Selecting AC External Walls for checking purpose
                    uidoc.Selection.SetElementIds(ids_3);
                }
                //////////////////////////////////Original Code Ends Here////////////////////////////////////////////

            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            UpdateProgress(100);
            System.Threading.Thread.Sleep(500);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Close the progress dialog
            progressDialog.Close();

            ETTV_F2 F2 = new ETTV_F2(commandData);
            F2.ShowDialog();

            try
            {
                // Testing for List

                StringBuilder builder = new StringBuilder();

                //builder.Append(WndwTypeLst.Count).AppendLine();
                //builder.Append(WndwTypeLst[0].FamilyName + " " + WndwTypeLst[0].Name).AppendLine();
                //builder.Append(ExtWallLst.Count).AppendLine();
                //builder.Append(WallTypeId_1.Count).AppendLine();
                //builder.Append(WallTypeId_2.Count).AppendLine();
                //builder.Append(WallTypeLst.Count).AppendLine();
                //builder.Append(WallTypeLst[0].FamilyName + " " + WallTypeLst[0].Name).AppendLine();

                //builder.Append(WallTypeId_2[0]).AppendLine();
                //builder.Append(WallTypeId_2[1]).AppendLine();
                //builder.Append(WallTypeId_2[2]).AppendLine();
                //builder.Append(WallTypeId_2[3]).AppendLine();
                //builder.Append(NewWallLst1[0].GetTypeId()).AppendLine();


                //builder.Append(N_Wndw_Id_List.Count).AppendLine();
                //builder.Append(N_Wndw_Id_List[0]).AppendLine();
                //builder.Append(N_Wndw_Id_List[1]).AppendLine();
                //builder.Append(N_Wndw_Id_List[2]).AppendLine();
                //builder.Append(N_Wndw_Id_List[3]).AppendLine();

                builder.Append("DONE!").AppendLine();
                
                TaskDialog.Show("Test", builder.ToString());

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
        }
    }
}