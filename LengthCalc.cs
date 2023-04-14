using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Controls;

namespace LengthCalc
{
    [Transaction(TransactionMode.ReadOnly)]
    public class LengthCalcCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get application and document objects
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Access the selected objects
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            // Filter and classify the selected objects
            var ductsPipesBySystemAndSize = new Dictionary<Tuple<string, string>, double>();

            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                string size = string.Empty;
                string systemType = string.Empty;
                double length = 0;

                if (elem is Duct duct)
                {
                    systemType = duct.DuctType.FamilyName;

                    if (systemType.IndexOf("round", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Parameter diameterParam = duct.LookupParameter("Diameter");
                        size = diameterParam.AsValueString();
                    }
                    else
                    {
                        Parameter widthParam = duct.LookupParameter("Width");
                        Parameter heightParam = duct.LookupParameter("Height");
                        double width = widthParam.AsDouble()*12;
                        double height = heightParam.AsDouble()*12;
                        size = $"{width} x {height}";
                    }

                    length = duct.LookupParameter("Length").AsDouble();
                }

                else if (elem is Pipe pipe)
                {
                    Parameter diameterParam = pipe.LookupParameter("Diameter");
                    size = diameterParam.AsValueString();
                    systemType = pipe.PipeType.FamilyName;
                    length = pipe.LookupParameter("Length").AsDouble();
                }
                else
                {
                    continue;
                }

                var key = new Tuple<string, string>(systemType, size);
                if (ductsPipesBySystemAndSize.ContainsKey(key))
                {
                    ductsPipesBySystemAndSize[key] += length;
                }
                else
                {
                    ductsPipesBySystemAndSize[key] = length;
                }
            }

            // Generate a table for the output
            //string tableHeader = "System Type\tSize\tTotal Length\n";
            //string tableData = string.Join("\n", ductsPipesBySystemAndSize.Select(
            //    kvp => $"{kvp.Key.Item1}\t{kvp.Key.Item2}\t{UnitUtils.ConvertFromInternalUnits(kvp.Value, UnitTypeId.Meters)} ft"));

            //string resultTable = tableHeader + tableData;

            //// Display the result
            //TaskDialog.Show("Duct and Pipe Lengths by System Type and Size", resultTable);

            var resultTableControl = new ResultTableControl();

            var resultList = ductsPipesBySystemAndSize.Select(
                kvp => new { SystemType = kvp.Key.Item1, Size = kvp.Key.Item2, TotalLength = Math.Round(UnitUtils.ConvertFromInternalUnits(kvp.Value, UnitTypeId.Feet), 2) }).ToList();


            resultTableControl.ResultDataGrid.ItemsSource = resultList;

            var window = new Window
            {
                Title = "Duct and Pipe Lengths by System Type and Size",
                Content = resultTableControl,
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            // Show the window as a dialog (modal)
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
