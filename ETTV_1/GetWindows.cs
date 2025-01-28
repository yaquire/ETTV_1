using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Macros;
using Autodesk.Revit.UI;



namespace boundingBox
{
  [TransactionAttribute(TransactionMode.Manual)]
  public class GetWindows : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {

      //get the uidoc
      Document doc = commandData.Application.ActiveUIDocument.Document;
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      //create a filtered collector for rooms
      List<Room> rooms = new List<Room>();
      FilteredElementCollector roomCollector = new FilteredElementCollector(doc)
          .OfCategory(BuiltInCategory.OST_Rooms)
          .WhereElementIsNotElementType();
      
      List<FamilyInstance> windows = new List<FamilyInstance>();
      FilteredElementCollector collector = new FilteredElementCollector(doc);
      collector.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilyInstance));
      // Add windows to the list
      
      foreach (Element element in collector)
      {
        if (element is FamilyInstance window)
        {
          if(CheckDirection(element) != string.Empty) windows.Add(window);
        }
      }
      
      try
      {
        foreach (Element element in roomCollector)
        {
          Room room = element as Room;
          rooms.Add(room);
        }
        List<(ElementId ids, Curve curve, Room room)> wallId_Curve = IntWallStuff(rooms,doc);
        int i=1;
        foreach (FamilyInstance window in windows)
        {
          Curve windoCurve=getWindowCurve(doc,window);
          //WriteToFile($"{i}---window {window.Id} curve= 0:{windoCurve.GetEndPoint(0)} 1:{windoCurve.GetEndPoint(1)}");
          
          String room_Window= WindowInRoom(windoCurve,wallId_Curve,window,doc);
          
          i++;
        }

        WriteToFile("-----------");
        TaskDialog.Show("Success", "End of Get Window");
	      return Result.Succeeded;
      }
      catch(Exception e)
      {
        return Result.Failed;
      }
    }
    public static List<ElementId> Curved_Angle_WallAttachment(List<Element> window, List<Wall> wall)
    {
	    List<ElementId> WindowElementIds = new List<ElementId>();
	    List<ElementId> WallElementIds = new List<ElementId>();
	    foreach (Element ele in window)
	    {
		    var familyInstance = ele as FamilyInstance;
		    if (familyInstance != null && familyInstance.Host != null)
		    {
                 TaskDialog.Show("a", "Window as Fam");
                    // Retrieve the host
                 Element host = familyInstance.Host;

                 // Return the host's ElementId
                WindowElementIds.Add(host.Id);
		    }
	    }
      foreach (Element ele in wall)
      {
			WallElementIds.Add(ele.Id);
      }
		//if the id of the wall the window is attached to matches the selected wall, returns 1, if not 0

		//Uses LINQ to compare
		return (WallElementIds);
    }
    public static String WindowInRoom(Curve windowcurve,List<(ElementId ids, Curve curve, Room room)> WallItems,FamilyInstance window,Document doc)
    {
      TaskDialog.Show("a","In the Window Room function");
      int j=1;
      List<(ElementId windowId, Room room, double topDis, double botDis)> winID_RoomName_Room = new List<(ElementId windowId, Room room, double topDis, double botDis)>();
      //a wall needs to be less than 10,000 feet long
      double bestDist = 1000;
      double dist;
      ElementId WindowHostID = window.Host.Id;
      string returned = "";
      foreach (var tupple in WallItems)
      {
        
        if (tupple.ids == WindowHostID)
        {
          double topLength = GetDistance(tupple.curve.GetEndPoint(1),windowcurve.GetEndPoint(1));
          double botLength = GetDistance(tupple.curve.GetEndPoint(0),windowcurve.GetEndPoint(0));
          winID_RoomName_Room.Add((window.Id, tupple.room, topLength, botLength));
          dist = topLength + botLength;
          if (dist < bestDist)
            bestDist = dist;
        }
      }
      foreach (var tuple2 in winID_RoomName_Room)
      {
        if (bestDist == (tuple2.topDis + tuple2.botDis))
        {
          if(tuple2.room.LookupParameter("AC_Space").AsInteger()==1)
          {
            Element wndw = doc.GetElement(tuple2.windowId);
            Room room = tuple2.room;
            wndw.LookupParameter("ETTV_Room").Set(room.Name);
            WriteToFile($"{j} Room: {tuple2.room.Name} ~ Wind:{tuple2.windowId}");
            j++;
          }
        }
      }
      
      return "";

    }
    public static double GetDistance(XYZ wallP,XYZ winP)
    {
      double distance = 0;

      distance = Math.Sqrt( (wallP.X-winP.X)* (wallP.X - winP.X) + (wallP.Y - winP.Y)* (wallP.Y - winP.Y) + (wallP.Z - winP.Z)* (wallP.Z - winP.Z));
      return distance;
    }
    public static List<(ElementId ids, Curve curve, Room room)> IntWallStuff(List<Room> rooms,Document docs)
    {
      double midHeight;
      List<(ElementId ids, Curve curve, Room room)> ExteriorWalls = new List<(ElementId ids, Curve curve, Room room)>();
      foreach (Room room in rooms)
      {
        if (room != null)
        {
          // Get the boundary segments of the room
          IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
          BoundingBoxXYZ roomBB = room.get_BoundingBox(null);
          midHeight = (roomBB.Max.Z - roomBB.Min.Z)/2;
          if (boundaries != null)
          {
            foreach (IList<BoundarySegment> boundarySegmentList in boundaries)
            {
              foreach (BoundarySegment segment in boundarySegmentList)
              {
                // Get the curve of the boundary segment
                Curve curve = segment.GetCurve();

                // Try to retrieve the wall associated with the boundary segment
                ElementId elementId = segment.ElementId;
                Element boundaryElement = docs.GetElement(elementId);

                if ( (CheckDirection(boundaryElement)) != "")
                {
                  //WriteToFile($"ELEMENT ID: {elementId}, Curve: {curve.Length * 304.8}");
                  XYZ translationVector = new XYZ(0,0,midHeight);

                  // Create a translation transform
                  Transform translationTransform = Transform.CreateTranslation(translationVector);

                  // Create a new curve by applying the transformation
                  Curve newCurve = curve.CreateTransformed(translationTransform);

                  // Output the details of the original and new curve
                  
                  ExteriorWalls.Add((elementId,newCurve, room));
                }
              }
            }
          }
        }
      }
      return ExteriorWalls;
    }
    public static Curve getWindowCurve(Document doc, FamilyInstance window)
    {
      List<Curve> curves = new List<Curve>();
      Options options = new Options
      {
        ComputeReferences = true,
        IncludeNonVisibleObjects = false
      };

      {
        BoundingBoxXYZ boundingBox = window.get_BoundingBox(null);
        XYZ min = boundingBox.Min;
        XYZ max = boundingBox.Max;
        //
        //WriteToFile($"{window.Id} min:{min} | max:{max}");
        // Define the bottom edges as curves
        XYZ bottomRight = new XYZ(max.X,min.Y,min.Z);
        XYZ topLeft = new XYZ(min.X,max.Y,min.Z);
        Curve windowCur = Line.CreateBound(bottomRight,topLeft);
        return windowCur;
      }
      return null;
    }
    public static Curve Make1Futher(Curve curve)
    {
      Curve reversedCurve = curve;
      XYZ OGpoint1 = curve.GetEndPoint(1);
      XYZ OGpoint0 = curve.GetEndPoint(0);
      //makes sure the point 1 is always above 0 & to the left of 0
      if (OGpoint1.X >= OGpoint0.X || OGpoint1.Y <= OGpoint0.Y)
      {
        //WriteToFile("CHANGED");
        Curve newCurve = curve.CreateReversed();
        return newCurve;
      }
      else 
      {
        //WriteToFile("No change");
        return reversedCurve;
      }
    }
    public static void WriteToFile(string content)
    {
      string filePath = "C:\\Users\\yaqub\\Desktop\\FYP\\boundingBox\\test.txt";
      using (StreamWriter swrite = new StreamWriter(filePath, append: true))
      {
        swrite.WriteLine(content);
      }
    }
    public void GetWallCurves(Document doc)
    {
      // Collect all walls in the document
      FilteredElementCollector collector = new FilteredElementCollector(doc)
          .OfClass(typeof(Wall));

      foreach (Element element in collector)
      {
        Wall wall = element as Wall;
        if (wall != null)
        {
          if(CheckDirection(wall)!=string.Empty)
          {
            // Get the Location property of the wall
            Location location = wall.Location;

            // Check if the Location is a LocationCurve
            if (location is LocationCurve locationCurve)
            {
              // Get the curve representing the wall's base geometry
              Curve wallCurve1 = locationCurve.Curve;

              // Check the type of the curve (Line or Arc)
              string curveType = wallCurve1.GetType().Name;
              Curve wallCurve = Make1Futher(wallCurve1);
              // Output the curve information
              //WriteToFile($"Wall ID: {wall.Id} Curve L:{wallCurve.Length*304.8} Curve Type: {curveType} Start: {wallCurve.GetEndPoint(0)} End: {wallCurve.GetEndPoint(1)}");
            }
          }
        }
      }
    }
    public static string CheckDirection(Element window)
        {
            if (window.LookupParameter("North").AsInteger() == 1)
            {
                string re = "North";
                return re;
            }
            if (window.LookupParameter("East").AsInteger() == 1)
            {
                string re = "East";
                return re;
            }
            if (window.LookupParameter("South").AsInteger() == 1)
            {
                string re = "South";
                return re;
            }
            if (window.LookupParameter("West").AsInteger() == 1)
            {
                string re = "West";
                return re;
            }
            if (window.LookupParameter("SouthEast").AsInteger() == 1)
            {
                string re = "SouthEast";
                return re;
            }
            if (window.LookupParameter("SouthWest").AsInteger() == 1)
            {
                string re = "SouthWest";
                return re;
            }
            if (window.LookupParameter("NorthWest").AsInteger() == 1)
            {
                string re = "NorthWest";
                return re;
            }
            if (window.LookupParameter("NorthEast").AsInteger() == 1)
            {
                string re = "NorthEast";
                return re;
            }

            return string.Empty;
        }
  }
}
