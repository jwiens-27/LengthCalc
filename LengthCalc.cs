using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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
                // Get the selected pipe elements
                List<Element> selectedPipes = GetSelectedPipes(uiDoc);

                if (selectedPipes.Count == 0)
                {
                    TaskDialog.Show("Error", "No pipes are selected. Please select some pipes and try again.");
                    return Result.Cancelled;
                }

                // Calculate the total length of selected pipes and separate it by system type
                Dictionary<string, double> pipeLengthsBySystem = CalculatePipeLengthsBySystem(selectedPipes);

                // Display the data in a TaskDialog
                DisplayPipeLengths(pipeLengthsBySystem);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private List<Element> GetSelectedPipes(UIDocument uiDoc)
        {
            ICollection<ElementId> selectedElementIds = uiDoc.Selection.GetElementIds();
            List<Element> selectedPipes = new List<Element>();

            foreach (ElementId id in selectedElementIds)
            {
                Element elem = uiDoc.Document.GetElement(id);
                if (elem is Pipe)
                {
                    selectedPipes.Add(elem);
                }
            }

            return selectedPipes;
        }

        private Dictionary<string, double> CalculatePipeLengthsBySystem(List<Element> pipes)
        {
            Dictionary<string, double> pipeLengthsBySystem = new Dictionary<string, double>();

            foreach (Pipe pipe in pipes)
            {
                string systemType = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                double pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();

                if (pipeLengthsBySystem.ContainsKey(systemType))
                {
                    pipeLengthsBySystem[systemType] += pipeLength;
                }

                else
                {
                    pipeLengthsBySystem.Add(systemType, pipeLength);
                }
            }

            return pipeLengthsBySystem;
        }

        private void DisplayPipeLengths(Dictionary<string, double> pipeLengthsBySystem)
        {
            // Prepare the TaskDialog
            TaskDialog dialog = new TaskDialog("Pipe Lengths by System")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                MainInstruction = "Total length of selected pipes by system:",
                AllowCancellation = false
            };

            // Add the data to the TaskDialog
            string data = string.Empty;
            foreach (var entry in pipeLengthsBySystem)
            {
                // Display the length in feet and inches
                double lengthFeet = entry.Value;
                int feet = (int)Math.Floor(lengthFeet);
                double inches = (lengthFeet - feet) * 12;

                data += $"{entry.Key}: {feet} ft {inches:F2} in{Environment.NewLine}";
            }

            dialog.MainContent = data;

            // Show the TaskDialog
            dialog.Show();
        }
    }
}
