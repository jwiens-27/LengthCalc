using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace LengthCalc
{
    [Transaction(TransactionMode.Manual)]
    public class LengthCalcCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Get the selected pipe and duct elements
                List<Element> selectedElements = GetSelectedElements(uiDoc);

                if (selectedElements.Count == 0)
                {
                    TaskDialog.Show("Error", "No pipes or ducts are selected. Please select some pipes or ducts and try again.");
                    return Result.Cancelled;
                }

                // Calculate the total length of selected elements and separate it by system type
                var lengthsBySystem = CalculateElementLengthsBySystem(selectedElements);

                // Display the data in a TaskDialog
                DisplayElementLengths(lengthsBySystem.Item1, lengthsBySystem.Item2);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private List<Element> GetSelectedElements(UIDocument uiDoc)
        {
            ICollection<ElementId> selectedElementIds = uiDoc.Selection.GetElementIds();
            List<Element> selectedElements = new List<Element>();

            foreach (ElementId id in selectedElementIds)
            {
                Element elem = uiDoc.Document.GetElement(id);
                if (elem is Pipe || elem is Duct)
                {
                    selectedElements.Add(elem);
                }
            }

            return selectedElements;
        }

        private Tuple<Dictionary<string, double>, Dictionary<string, double>> CalculateElementLengthsBySystem(List<Element> elements)
        {
            Dictionary<string, double> pipeLengthsBySystem = new Dictionary<string, double>();
            Dictionary<string, double> ductLengthsBySystem = new Dictionary<string, double>();

            foreach (Element element in elements)
            {
                string systemType;
                double length;

                if (element is Pipe)
                {
                    Pipe pipe = element as Pipe;
                    systemType = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    length = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    if (pipeLengthsBySystem.ContainsKey(systemType))
                    {
                        pipeLengthsBySystem[systemType] += length;
                    }
                    else
                    {
                        pipeLengthsBySystem.Add(systemType, length);
                    }
                }
                else if (element is Duct)
                {
                    Duct duct = element as Duct;
                    systemType = duct.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString();
                    length = duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    if (ductLengthsBySystem.ContainsKey(systemType))
                    {
                        ductLengthsBySystem[systemType] += length;
                    }
                    else
                    {
                        ductLengthsBySystem.Add(systemType, length);
                    }
                }
            }

            return new Tuple<Dictionary<string, double>, Dictionary<string, double>>(pipeLengthsBySystem, ductLengthsBySystem);
        }

        private void DisplayElementLengths(Dictionary<string, double> pipeLengthsBySystem, Dictionary<string, double> ductLengthsBySystem)
        {
            // Prepare the TaskDialog
            TaskDialog dialog = new TaskDialog("Element Lengths by System")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                MainInstruction = "Total length of selected pipes and ducts by system:",
                AllowCancellation = false
            };

            // Add the data to the TaskDialog
            string data = string.Empty;

            if (pipeLengthsBySystem.Count > 0)
            {
                data += "Pipes:\n";
                double totalPipeLength = 0;

                foreach (var entry in pipeLengthsBySystem)
                {
                    // Display the length in feet and inches
                    double lengthFeet = entry.Value;
                    int feet = (int)Math.Floor(lengthFeet);
                    double inches = (lengthFeet - feet) * 12;

                    data += $"{entry.Key}: {feet} ft {inches:F2} in{Environment.NewLine}";
                    totalPipeLength += lengthFeet;
                }

                // Calculate the total length of pipe systems in feet rounded up to the nearest 5 feet
                int totalPipeLengthRounded = (int)Math.Ceiling(totalPipeLength / 5) * 5;
                data += $"{Environment.NewLine}Total Pipe Length (rounded up to nearest 5 ft): {totalPipeLengthRounded} ft{Environment.NewLine}";
            }

            if (ductLengthsBySystem.Count > 0)
            {
                if (data.Length > 0) data += "\n";
                data += "Ducts:\n";
                double totalDuctLength = 0;

                foreach (var entry in ductLengthsBySystem)
                {
                    // Display the length in feet and inches
                    double lengthFeet = entry.Value;
                    int feet = (int)Math.Floor(lengthFeet);
                    double inches = (lengthFeet - feet) * 12;

                    data += $"{entry.Key}: {feet} ft {inches:F2} in{Environment.NewLine}";
                    totalDuctLength += lengthFeet;
                }

                // Calculate the total length of duct systems in feet rounded up to the nearest 5 feet
                int totalDuctLengthRounded = (int)Math.Ceiling(totalDuctLength / 5) * 5;
                data += $"{Environment.NewLine}Total Duct Length (rounded up to nearest 5 ft): {totalDuctLengthRounded} ft{Environment.NewLine}";
            }

            dialog.MainContent = data;

            // Show the TaskDialog
            dialog.Show();
        }
    }
}
