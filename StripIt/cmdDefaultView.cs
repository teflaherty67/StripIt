using System.Windows.Controls;

namespace StripIt
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDefaultView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get Revit app and doc objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // Step 1: Try to find the default 3D view named "{3D}"
            FilteredElementCollector collector = new FilteredElementCollector(curDoc);
            View3D default3DView = collector
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate && v.Name == "{3D}");

            // Step 2: If found, make it the active view and return
            if (default3DView != null)
            {
                uidoc.ActiveView = default3DView;
                return Result.Succeeded;
            }

            // Step 3: Get the 3D view family type
            ViewFamilyType viewFamilyType = new FilteredElementCollector(curDoc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.ThreeDimensional);

            if (viewFamilyType == null)
            {
                message = "No ViewFamilyType for 3D views found.";
                return Result.Failed;
            }

            // Step 4: Create a new isometric 3D view named "{3D}"
            using (Transaction tx = new Transaction(curDoc, "Create Default 3D View"))
            {
                tx.Start();
                default3DView = View3D.CreateIsometric(curDoc, viewFamilyType.Id);
                default3DView.Name = "3D";
                tx.Commit();
            }

            // Optionally activate the new view
            uidoc.ActiveView = default3DView;

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            clsButtonData myButtonData = new clsButtonData(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }
    }

}
