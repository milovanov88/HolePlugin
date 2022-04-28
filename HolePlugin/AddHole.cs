using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents
                             .OfType<Document>()
                             .Where(x => x.Title.Contains("ОВ"))
                             .FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbolRectangular = new FilteredElementCollector(arDoc)
            .OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_GenericModel)
            .OfType<FamilySymbol>()
            .Where(x => x.FamilyName.Equals("Отверстие Прямоугольное"))
            .FirstOrDefault();

            FamilySymbol familySymbolRound = new FilteredElementCollector(arDoc)
            .OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_GenericModel)
            .OfType<FamilySymbol>()
            .Where(x => x.FamilyName.Equals("Отверстие Круглое"))
            .FirstOrDefault();

            if (familySymbolRectangular == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстие Прямоугольное\"");
                return Result.Cancelled;
            }

            if (familySymbolRound == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстие Круглое\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
            .OfClass(typeof(Duct))
            .OfType<Duct>()
            .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
            .OfClass(typeof(Pipe))
            .OfType<Pipe>()
            .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
            .OfClass(typeof(View3D))
            .OfType<View3D>()
            .Where(x => !x.IsTemplate)
            .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не нaйден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transactionduct = new Transaction(arDoc);
            transactionduct.Start("Расстановка отверстий");
            if (!familySymbolRectangular.IsActive)
            {
                familySymbolRectangular.Activate();
            }
            foreach (Duct duct in ducts)
            {
                Line curve = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                .Where(x => x.Proximity <= curve.Length)
                .Distinct(new ReferenceWithContextElementEqualityComparer())
                .ToList();
                foreach (ReferenceWithContext intersection in intersections)
                {
                    double proximity = intersection.Proximity;
                    Reference reference = intersection.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + direction * proximity;
                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbolRectangular, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble());
                    height.Set(duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble());
                }
            }
            transactionduct.Commit();
            Transaction transactionpipe = new Transaction(arDoc);
            transactionpipe.Start("Расстановка отверстий");
            if (!familySymbolRound.IsActive)
            {
                familySymbolRound.Activate();
            }
            foreach (Pipe pipe in pipes)
            {
                Line curve = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                .Where(x => x.Proximity <= curve.Length)
                .Distinct(new ReferenceWithContextElementEqualityComparer())
                .ToList();
                foreach (ReferenceWithContext intersection in intersections)
                {
                    double proximity = intersection.Proximity;
                    Reference reference = intersection.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + direction * proximity;
                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbolRound, wall, level, StructuralType.NonStructural);
                    Parameter diameter = hole.LookupParameter("Диаметр");
                    diameter.Set(pipe.Diameter);
                }
            }
            transactionpipe.Commit();
            return Result.Succeeded;
        }
    }
    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();
            var yReference = y.GetReference();
            return xReference.LinkedElementId == yReference.LinkedElementId
                      && xReference.ElementId == yReference.ElementId;
        }
        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();
            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}



