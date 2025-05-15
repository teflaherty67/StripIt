using StripIt;
using System.Windows.Controls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace StripIt
{
    [Transaction(TransactionMode.Manual)]
    public class cmdStripIt : IExternalCommand
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









            string userName = uiapp.Application.Username;

            // 01. set the active view to Elevation A 3D view
            View newView;

            if (curDoc.IsWorkshared == true)
                newView = Utils.GetViewByName(curDoc, "A - " + userName);
            else
                newView = Utils.GetViewByName(curDoc, "A");

            uidoc.ActiveView = newView;

            // 02. create list of all sheets
            List<ViewSheet> allSheets = Utils.GetAllSheets(curDoc);

            // 03. create view lists

            // create list of all views
            List<View> allViews = Utils.GetAllViews(curDoc);

            // create list of views to keep
            List<View> viewsToKeep = Utils.GetAllViewsByCategoryContains(curDoc, "Presentation Views");

            // create list of ElementIds from viewsToKeep for fast lookup
            List<ElementId> viewsToKeepIds = new List<ElementId>(viewsToKeep.Select(view => view.Id));

            // create the filtered list using the ElementId for comparison
            List<View> viewsToDelete = allViews.Where(view => !viewsToKeepIds.Contains(view.Id)).ToList();

            // 04. create list of all schedules
            List<ViewSchedule> allSchedules = Utils.GetAllSchedules(curDoc);

            // 05. get Revit Command Id for Purge
            RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(PostableCommand.PurgeUnused);

            // 04. create & start transaction
            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Strip the file");

                // 02a. delete all sheets
                foreach(ViewSheet curSheet in allSheets)
                {
                    curDoc.Delete(curSheet.Id);
                }

                // 03a. loop through the views & delete them
                foreach (View deleteView in viewsToDelete)
                    {
                        try
                        {
                            // delete the view
                            curDoc.Delete(deleteView.Id);
                        }
                        catch (Exception)
                        {                            
                        }
                    }

                // 04a. loop through the schedules & delete them

                foreach (ViewSchedule deleteSched in allSchedules)
                    {
                        curDoc.Delete(deleteSched.Id);
                    }

                // run the Purge Unused command using PostCommand
                uiapp.PostCommand(commandId);

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