using System.Windows.Controls;

namespace StripIt
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDefaultView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            ViewFamilyType viewFamilyType = new FilteredElementCollector(curDoc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.ThreeDimensional);

            if (viewFamilyType == null)
                throw new InvalidOperationException("No 3D view family type found.");

            using (Transaction t = new Transaction(curDoc, "Create Default 3D View"))
            {
                t.Start("Create 3D view");

                // Create an isometric 3D view
                View3D view3D = View3D.CreateIsometric(curDoc, viewFamilyType.Id);

                // Set the name to "{3D}" (Revit's default 3D view name)
                view3D.Name = "{3D}";

                t.Commit();
            }

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
