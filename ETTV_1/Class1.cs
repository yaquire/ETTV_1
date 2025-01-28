﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ETTV_1
{
    public static class JtBoundingBoxXyzExtensionMethods
    {
        /// <summary>
        /// Expand the given bounding box to include 
        /// and contain the given point.
        /// </summary>
        public static void ExpandToContain(
          this BoundingBoxXYZ bb,
          XYZ p)
        {
            bb.Min = new XYZ(Math.Min(bb.Min.X, p.X),
              Math.Min(bb.Min.Y, p.Y),
              Math.Min(bb.Min.Z, p.Z));

            bb.Max = new XYZ(Math.Max(bb.Max.X, p.X),
              Math.Max(bb.Max.Y, p.Y),
              Math.Max(bb.Max.Z, p.Z));
        }

        /// <summary>
        /// Expand the given bounding box to include 
        /// and contain the given other one.
        /// </summary>
        public static void ExpandToContain(
          this BoundingBoxXYZ bb,
          BoundingBoxXYZ other)
        {
            bb.ExpandToContain(other.Min);
            bb.ExpandToContain(other.Max);
        }

    }
}
