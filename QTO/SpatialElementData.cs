using System;
using System.Globalization;
using Autodesk.Revit.DB;

namespace QTO
{
    internal class SpatialElementData
    {
        public string LocationType { get; init; } = "";
        public string PositionXFeet { get; init; } = "";
        public string PositionYFeet { get; init; } = "";
        public string PositionZFeet { get; init; } = "";
        public string StartXFeet { get; init; } = "";
        public string StartYFeet { get; init; } = "";
        public string StartZFeet { get; init; } = "";
        public string EndXFeet { get; init; } = "";
        public string EndYFeet { get; init; } = "";
        public string EndZFeet { get; init; } = "";
        public string RotationDegrees { get; init; } = "";
        public string BoundingBoxMinXFeet { get; init; } = "";
        public string BoundingBoxMinYFeet { get; init; } = "";
        public string BoundingBoxMinZFeet { get; init; } = "";
        public string BoundingBoxMaxXFeet { get; init; } = "";
        public string BoundingBoxMaxYFeet { get; init; } = "";
        public string BoundingBoxMaxZFeet { get; init; } = "";
        public string BoundingBoxCenterXFeet { get; init; } = "";
        public string BoundingBoxCenterYFeet { get; init; } = "";
        public string BoundingBoxCenterZFeet { get; init; } = "";

        public static SpatialElementData FromElement(Element elem)
        {
            BoundingBoxXYZ? boundingBox = null;
            try
            {
                boundingBox = elem.get_BoundingBox(null);
            }
            catch
            {
                boundingBox = null;
            }

            XYZ? bboxCenter = boundingBox == null
                ? null
                : new XYZ(
                    (boundingBox.Min.X + boundingBox.Max.X) / 2.0,
                    (boundingBox.Min.Y + boundingBox.Max.Y) / 2.0,
                    (boundingBox.Min.Z + boundingBox.Max.Z) / 2.0
                );

            string locationType = "";
            XYZ? position = null;
            XYZ? start = null;
            XYZ? end = null;
            string rotationDegrees = "";

            if (elem.Location is LocationPoint locationPoint)
            {
                locationType = "Point";
                position = locationPoint.Point;
                rotationDegrees = FormatDouble(locationPoint.Rotation * 180.0 / Math.PI);
            }
            else if (elem.Location is LocationCurve locationCurve)
            {
                locationType = "Curve";
                try
                {
                    Curve curve = locationCurve.Curve;
                    start = curve.GetEndPoint(0);
                    end = curve.GetEndPoint(1);
                    position = start != null && end != null
                        ? new XYZ(
                            (start.X + end.X) / 2.0,
                            (start.Y + end.Y) / 2.0,
                            (start.Z + end.Z) / 2.0
                        )
                        : bboxCenter;
                }
                catch
                {
                    position = bboxCenter;
                }
            }
            else if (bboxCenter != null)
            {
                locationType = "BoundingBox";
                position = bboxCenter;
            }

            return new SpatialElementData
            {
                LocationType = locationType,
                PositionXFeet = FormatCoordinate(position?.X),
                PositionYFeet = FormatCoordinate(position?.Y),
                PositionZFeet = FormatCoordinate(position?.Z),
                StartXFeet = FormatCoordinate(start?.X),
                StartYFeet = FormatCoordinate(start?.Y),
                StartZFeet = FormatCoordinate(start?.Z),
                EndXFeet = FormatCoordinate(end?.X),
                EndYFeet = FormatCoordinate(end?.Y),
                EndZFeet = FormatCoordinate(end?.Z),
                RotationDegrees = rotationDegrees,
                BoundingBoxMinXFeet = FormatCoordinate(boundingBox?.Min.X),
                BoundingBoxMinYFeet = FormatCoordinate(boundingBox?.Min.Y),
                BoundingBoxMinZFeet = FormatCoordinate(boundingBox?.Min.Z),
                BoundingBoxMaxXFeet = FormatCoordinate(boundingBox?.Max.X),
                BoundingBoxMaxYFeet = FormatCoordinate(boundingBox?.Max.Y),
                BoundingBoxMaxZFeet = FormatCoordinate(boundingBox?.Max.Z),
                BoundingBoxCenterXFeet = FormatCoordinate(bboxCenter?.X),
                BoundingBoxCenterYFeet = FormatCoordinate(bboxCenter?.Y),
                BoundingBoxCenterZFeet = FormatCoordinate(bboxCenter?.Z)
            };
        }

        private static string FormatCoordinate(double? value)
        {
            return value.HasValue ? FormatDouble(value.Value) : "";
        }

        private static string FormatDouble(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
