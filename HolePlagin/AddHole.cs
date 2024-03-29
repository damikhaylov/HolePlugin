﻿using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;

namespace HolePlagin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>()
                .Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                message = "Не найден ОВ файл.";
                return Result.Failed;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                message = "Не найдено семейство \"Отверстие\".";
                return Result.Failed;
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
                .Where(x => ! x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null)
            {
                message = "Не найден 3D вид.";
                return Result.Failed;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)),
                FindReferenceTarget.Element, view3D);

            Transaction transaction0 = new Transaction(arDoc);
            transaction0.Start("Активация семейства отверстия");
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            transaction0.Commit();

            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");

            foreach (Duct duct in ducts)
            {
                Line line = (duct.Location as LocationCurve).Curve as Line;
                AddHoleByLine(arDoc, referenceIntersector, line, familySymbol, duct.Diameter, duct.Diameter);
            }

            foreach (Pipe pipe in pipes)
            {
                Line line = (pipe.Location as LocationCurve).Curve as Line;
                double dimension = UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);
                dimension = Math.Max(dimension, pipe.Diameter);
                AddHoleByLine(arDoc, referenceIntersector, line, familySymbol, dimension, dimension);
            }

            transaction.Commit();
            return Result.Succeeded;
        }

        private static void AddHoleByLine(Document document, ReferenceIntersector referenceIntersector, 
            Line line, FamilySymbol familySymbol, double width, double height)
        {
            XYZ point = line.GetEndPoint(0);
            XYZ direction = line.Direction;

            List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                .Where(x => x.Proximity <= line.Length)
                .Distinct(new ReferenceWithContextElementEqualityComparer())
                .ToList();
            foreach (ReferenceWithContext refer in intersections)
            {
                double proximity = refer.Proximity;
                Reference reference = refer.GetReference();
                Wall wall = document.GetElement(reference.ElementId) as Wall;
                Level level = document.GetElement(wall.LevelId) as Level;
                XYZ pointHole = point + (direction * proximity);

                FamilyInstance hole = document.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                Parameter holeWidth = hole.LookupParameter("Ширина");
                Parameter holeHeight = hole.LookupParameter("Высота");
                holeWidth.Set(width);
                holeHeight.Set(height);
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
}
