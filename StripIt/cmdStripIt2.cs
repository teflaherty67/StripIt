using System.Windows.Controls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace StripIt
{
    [Transaction(TransactionMode.Manual)]
    public class cmdStripIt2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

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

            // 04. create & start transaction
            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Strip the file");

                // 02a. delete all sheets
                foreach(ViewSheet curSheet in allSheets)
                {
                    curDoc.Delete(curSheet.Id);
                }

                // 03. create view lists

                // Assuming allViews and viewsToKeep are both lists of Revit View objects
                List<View> allViews = Utils.GetAllViews(curDoc);
                List<View> viewsToKeep = Utils.GetAllViewsByCategoryContains(curDoc, "Presentation Views");

                // Create a HashSet of ElementIds from viewsToKeep for fast lookup
                List<ElementId> viewsToKeepIds = new List<ElementId>(viewsToKeep.Select(view => view.Id));

                // Create the filtered list using the ElementId for comparison
                List<View> viewsToDelete = allViews.Where(view => !viewsToKeepIds.Contains(view.Id)).ToList();

                //List<View> allViews = Utils.GetAllViews(curDoc);

                //List<View> viewsToKeep = Utils.GetAllViewsByCategoryContains(curDoc, "Presentation Views");

                //List<View> viewsToDelete = allViews.Where(view => !viewsToKeep.Contains(view)).ToList();


                //int counter = 2;

                //while (counter > 0)
                //{
                //    // loop through the views
                //    foreach (View curView in listViews)
                //    {
                //        // check if view has dependent views
                //        if (curView.GetDependentViewIds().Count() == 0)
                //        {
                //            // add view to list of views to delete
                //            viewsToDelete.Add(curView);
                //        }
                //    }

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

                //    counter--;
                //}              


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


