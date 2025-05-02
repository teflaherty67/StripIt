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
            View curView;

            if (curDoc.IsWorkshared == true)
                curView = Utils.GetViewByName(curDoc, "A - " + userName);
            else
                curView = Utils.GetViewByName(curDoc, "A");

            uidoc.ActiveView = curView;

            // 02. get all the sheets
            FilteredElementCollector colSheets = new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType();

            using(Transaction t = new Transaction(curDoc))
            {
                t.Start("Strip the file");

                foreach(ViewSheet curSheet in colSheets)
                {
                    curDoc.Delete(curSheet.Id);
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
