using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System.Runtime.InteropServices.WindowsRuntime;

namespace RAA_Module2_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            /*
           For this session's challenge, you will create an add-in to generate Revit elements from model lines. 
           The add-in will create different Revit elements based on the line style name. Running the add-in on the file 
           provided below will reveal a hidden message in the Level 1 floor plan view.

           The add-in should prompt you to select elements. It should then filter the elements for model curves. 
           Once you've filtered the elements, loop through them and create the following Revit elements based on the line's line style:

               A-GLAZ - Storefront wall
               A-WALL - Generic 8" wall
               M-DUCT - Default duct
               P-PIPE - Default pipe

           Click the folder icon on the top right-hand side of the screen to download the starter file for this challenge.

           Remember to set the view's Detail Level to "fine" to see the pipes. To start the challenge, create a new command 
           file in your solution called "Module02Challenge". When complete, enter the hidden message and 
           submit a link to your solution on GitHub using the "Challenge Submission" page. Good luck! 

           Bonus!

           As a bonus, create custom methods to get the various types by name (wall type, pipe type, duct type, and system type). Incorporate these methods into your command code. 
            */

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Module02 Command");

                UIDocument uidoc = uiapp.ActiveUIDocument;
                IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select Elements");

                List<CurveElement> modelCurves = new List<CurveElement>();

                foreach (Element element in pickList)
                {
                    if (element is CurveElement)
                    {
                        modelCurves.Add((CurveElement)element);
                    }
                }

                List<XYZ> startPoints = new List<XYZ>();
                List<XYZ> endPoints = new List<XYZ>();

                foreach (CurveElement currentCurve in modelCurves)
                {
                    Curve curve = currentCurve.GeometryCurve;
                    XYZ startPoint = curve.GetEndPoint(0);
                    XYZ endPoint = curve.GetEndPoint(1);
                    startPoints.Add(startPoint);
                    endPoints.Add(endPoint);

                    GraphicsStyle currentStyle = (GraphicsStyle)currentCurve.LineStyle;

                    {
                        switch (currentStyle.Name)
                        {
                            case "A-GLAZ":
                                Wall.Create(doc, curve, GetWallTypeByName(doc, "Storefront wall").Id, false);
                                break;
                            case "A-WALL":
                                Wall.Create(doc, curve, GetWallTypeByName(doc, "Generic 8 wall").Id, false);
                                break;
                            case "M-DUCT":
                                Duct.Create(doc, GetMEPSystemTypeByName(doc, "Supply Air").Id, GetDuctTypeByName(doc, "Default Duct").Id, GetLevelByName(doc, "Level 1").Id, startPoint, endPoint);
                                break;
                            case "P-PIPE":
                                Pipe.Create(doc, GetMEPSystemTypeByName(doc, "Other").Id, GetPipeTypeByName(doc, "Default Pipe").Id, GetLevelByName(doc, "Level 1").Id, startPoint, endPoint);
                                break;

                        }
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;
        }

        internal WallType GetWallTypeByName(Document doc,
                                            string wallName)
        {
            FilteredElementCollector wallTypes = new FilteredElementCollector(doc);
            wallTypes.OfClass(typeof(WallType));

            foreach (WallType wallType in wallTypes)
            {
                if (wallType.Name == wallName)
                {
                    return wallType;
                }
                else
                {
                    return null;
                }
            }
        }

        internal PipeType GetPipeTypeByName(Document doc, string pipeName)
        {
            FilteredElementCollector pipeTypes = new FilteredElementCollector(doc);
            pipeTypes.OfClass(typeof(PipeType));

            foreach (PipeType pipeType in pipeTypes)
            {
                if (pipeType.Name == pipeName)
                {
                    return pipeType;
                }
            }
        }

        internal DuctType GetDuctTypeByName(Document doc, string ductName)
        {
            FilteredElementCollector ductTypes = new FilteredElementCollector(doc);
            ductTypes.OfClass(typeof(DuctType));

            foreach (DuctType ductType in ductTypes)
            {
                if (ductType.Name == ductName)
                {
                    return ductType;
                }
                else
                {
                    return null;
                }
            }
        }

        internal MEPSystemType GetMEPSystemTypeByName(Document doc, string mepSystemName)
        {
            FilteredElementCollector mepSystems = new FilteredElementCollector(doc);
            mepSystems.OfClass(typeof(MEPSystemType));

            foreach (MEPSystemType mepSystem in mepSystems)
            {
                if (mepSystem.Name == mepSystemName)
                {
                    return mepSystem;
                }
            }

            return null;
        }

        internal Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector allLevels = new FilteredElementCollector(doc);
            allLevels.OfClass(typeof(Level));

            foreach (Level level in allLevels)
            {
                if (level.Name == levelName)
                {
                    return level;
                }
            }

            return null;
        }


        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";
            string? methodBase = MethodBase.GetCurrentMethod().DeclaringType?.FullName;

            if (methodBase == null)
            {
                throw new InvalidOperationException("MethodBase.GetCurrentMethod().DeclaringType?.FullName is null");
            }
            else
            {
                Common.ButtonDataClass myButtonData1 = new Common.ButtonDataClass(
                    buttonInternalName,
                    buttonTitle,
                    methodBase,
                    Properties.Resources.Blue_32,
                    Properties.Resources.Blue_16,
                    "This is a tooltip for Button 1");

                return myButtonData1.Data;
            }
        }
    }

}
