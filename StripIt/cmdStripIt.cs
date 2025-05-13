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

                // 03. create list of views to delete
                List<View> viewsToDelete = new List<View>();

                // get all the views in the project by category
                List<View> listViews = new List<View>();
                List<View> listCat01 = Utils.GetAllViewsByCategoryContains(curDoc, "Floor Plans");
                List<View> listCat02 = Utils.GetAllViewsByCategoryContains(curDoc, "Elevations");
                List<View> listCat03 = Utils.GetAllViewsByCategoryContains(curDoc, "Roof Plans");
                List<View> listCat04 = Utils.GetAllViewsByCategoryContains(curDoc, "Sections");
                List<View> listCat05 = Utils.GetAllViewsByCategoryContains(curDoc, "Interior Elevations");
                List<View> listCat06 = Utils.GetAllViewsByCategoryContains(curDoc, "Electrical Plans");
                List<View> listCat07 = Utils.GetAllViewsByCategoryContains(curDoc, "Form/Foundation Plans");
                List<View> listCat08 = Utils.GetAllViewsByCategoryContains(curDoc, "Ceiling Framing Plans");
                List<View> listCat09 = Utils.GetAllViewsByCategoryContains(curDoc, "Roof Framing Plans");
                List<View> listCat14 = Utils.GetAllViewsByCategoryContains(curDoc, "Ceiling Views");
                List<View> listCat16 = Utils.GetAllViewsByCategoryAndViewTemplate(curDoc, "16:3D Views", "16-3D Frame");

                // combine the lists together
                listViews.AddRange(listCat01);
                listViews.AddRange(listCat02);
                listViews.AddRange(listCat03);
                listViews.AddRange(listCat04);
                listViews.AddRange(listCat05);
                listViews.AddRange(listCat06);
                listViews.AddRange(listCat07);
                listViews.AddRange(listCat08);
                listViews.AddRange(listCat09);
                listViews.AddRange(listCat14);
                listViews.AddRange(listCat16);

                int counter = 2;

                while (counter > 0)
                {
                    counter--;
                }

                // get all sheets in project
                FilteredElementCollector colSheets = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(ViewSheet));

                // loop through the views
                foreach (View curView in listViews)
                {
                    // check if view has dependent views
                    if (curView.GetDependentViewIds().Count() == 0)
                    {
                        // add view to list of views to delete
                        viewsToDelete.Add(curView);
                    }
                }

                foreach (View deleteView in viewsToDelete)
                {
                    // delete the view
                    curDoc.Delete(deleteView.Id);
                }

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
