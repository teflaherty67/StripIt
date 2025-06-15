using StripIt.Common;

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

            // Step 1: check for the default 3D view
            FilteredElementCollector collector = new FilteredElementCollector(curDoc);
            View3D default3DView = collector
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate && v.Name == "{3D}");

            // Step 2: if found, make it the active view
            if (default3DView != null)
            {
                uidoc.ActiveView = default3DView;
                return Result.Succeeded;
            }

            // Step 3: if not found, get the 3D view family type
            ViewFamilyType viewFamilyType = new FilteredElementCollector(curDoc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.ThreeDimensional);

            if (viewFamilyType == null)
            {
                message = "No ViewFamilyType for 3D views found.";
                return Result.Failed;
            }

            // Step 4: create new 3D view
            using (Transaction tx = new Transaction(curDoc))
            {
                tx.Start("Create Default 3D View");

                default3DView = View3D.CreateIsometric(curDoc, viewFamilyType.Id);
                default3DView.Name = "3D";

                tx.Commit();
            }

            // Step 5: set active view
            uidoc.ActiveView = default3DView;

            // 01. create list of all sheets
            List<ViewSheet> allSheets = Utils.GetAllSheets(curDoc);

            // 02. create view lists

            // create list of all views
            List<View> allViews = Utils.GetAllViews(curDoc);

            // create list of views to keep
            List<View> viewsToKeep = Utils.GetAllViewsByCategoryContains(curDoc, "Presentation Views");

            // create list of ElementIds from viewsToKeep for fast lookup
            List<ElementId> viewsToKeepIds = new List<ElementId>(viewsToKeep.Select(view => view.Id));

            // create the filtered list using the ElementId for comparison
            List<View> viewsToDelete = allViews.Where(view => !viewsToKeepIds.Contains(view.Id)).ToList();

            // 03. create list of all schedules
            List<ViewSchedule> allSchedules = Utils.GetAllSchedules(curDoc);

            // 04. get Revit Command Id for Purge
            RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(PostableCommand.PurgeUnused);

            // 05. create & start transaction
            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Strip the file");

                // 01a. delete all sheets
                foreach (ViewSheet curSheet in allSheets)
                {
                    curDoc.Delete(curSheet.Id);
                }

                // 02a. loop through the views & delete them
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

                // 03a. loop through the schedules & delete them

                foreach (ViewSchedule deleteSched in allSchedules)
                {
                    curDoc.Delete(deleteSched.Id);
                }

                // 06. purged unused elements programatically
                Utils.PurgeUnusedFamilySymbols(curDoc);
                Utils.PurgeUnusedViewTemplates(curDoc);
                Utils.PurgeUnusedFilters(curDoc);
                Utils.PurgeUnusedMaterials(curDoc);
                Utils.PurgeUnusedLinePatterns(curDoc);
                Utils.PurgeUnusedFillPatterns(curDoc);
                Utils.PurgeUnusedGroups(curDoc);
                Utils.PurgeUnusedTextStyles(curDoc);
                Utils.PurgeUnusedDimensionStyles(curDoc);
                Utils.PurgeUnusedLineStyles(curDoc);
                Utils.PurgeUnusedAnnotationSymbols(curDoc);
                //Utils.PurgeUnusedLoadedFamilies(curDoc);
                Utils.PurgeUnusedAppearanceAssets(curDoc);
                Utils.PurgeUnusedRenderingMaterials(curDoc);

                t.Commit();
            }

            // 07. save file to Luis' folder
            Utils.SaveToLuisFolder(curDoc);

            // 08. switch to Manage ribbon tab before running Purge command
            try
            {
                // Get the ribbon control
                Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;

                // Find and activate the Manage tab
                foreach (Autodesk.Windows.RibbonTab tab in ribbon.Tabs)
                {
                    if (tab.Id == "Manage" || tab.Title == "Manage")
                    {
                        tab.IsActive = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't fail the command
                System.Diagnostics.Debug.WriteLine($"Failed to switch to Manage ribbon: {ex.Message}");
            }

            // 06a. run the Purge Unused command using PostCommand
            uiapp.PostCommand(commandId);

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