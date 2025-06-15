namespace StripIt.Common
{
    internal static class Utils
    {
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        internal static View GetViewByName(Document curDoc, string viewName)
        {
            List<View> viewList = GetAllViews(curDoc);

            //loop through views in the collector
            foreach (View curView in viewList)
            {
                if (curView.Name == viewName && curView.IsTemplate == false)
                {
                    return curView;
                }
            }

            return null;
        }

        internal static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_views = new List<View>();
            foreach (View x in m_colviews.ToElements())
            {
                m_views.Add(x);
            }

            return m_views;
        }

        internal static List<ViewSheet> GetAllSheets(Document curDoc)
        {
            //get all sheets
            FilteredElementCollector m_colViews = new FilteredElementCollector(curDoc);
            m_colViews.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> m_sheets = new List<ViewSheet>();
            foreach (ViewSheet x in m_colViews.ToElements())
            {
                m_sheets.Add(x);
            }

            return m_sheets;
        }

        internal static List<View> GetAllViewsByCategoryContains(Document curDoc, string catName)
        {
            List<View> m_colViews = GetAllViewsByCategory(curDoc, catName);

            List<View> m_returnList = new List<View>();

            foreach (View curView in m_colViews)
            {
                string viewCat = GetParameterValueByName(curView, "Category");

                if (viewCat.Contains(catName))
                    m_returnList.Add(curView);
            }

            return m_returnList;
        }

        internal static List<View> GetAllViewsByCategoryAndViewTemplate(Document curDoc, string catName, string vtName)
        {
            List<View> m_colViews = GetAllViewsByCategory(curDoc, catName);

            List<View> m_returnList = new List<View>();

            foreach (View curView in m_colViews)
            {
                ElementId vtId = curView.ViewTemplateId;

                if (vtId != ElementId.InvalidElementId)
                {
                    View vt = curDoc.GetElement(vtId) as View;

                    if (vt.Name == vtName)
                        m_returnList.Add(curView);
                }
            }

            return m_returnList;
        }

        internal static List<View> GetAllViewsByCategory(Document curDoc, string catName)
        {
            List<View> m_colViews = GetAllViews(curDoc);

            List<View> m_returnList = new List<View>();

            foreach (View curView in m_colViews)
            {
                string viewCat = GetParameterValueByName(curView, "Category");
                if (!string.IsNullOrEmpty(viewCat))

                {
                    if (viewCat.Contains(catName))
                        m_returnList.Add(curView);
                }

            }

            return m_returnList;
        }

        internal static List<ViewSchedule> GetAllSchedules(Document curDoc)
        {
            List<ViewSchedule> m_schedList = new List<ViewSchedule>();

            FilteredElementCollector curCollector = new FilteredElementCollector(curDoc);
            curCollector.OfClass(typeof(ViewSchedule));
            curCollector.WhereElementIsNotElementType();

            //loop through views and check if schedule - if so then put into schedule list
            foreach (ViewSchedule curView in curCollector)
            {
                if (curView.ViewType == ViewType.Schedule)
                {
                    if (curView.IsTemplate == false)
                    {
                        if (curView.Name.Contains("<") && curView.Name.Contains(">"))
                            continue;
                        else
                            m_schedList.Add((ViewSchedule)curView);
                    }
                }
            }

            return m_schedList;
        }

        private static string GetParameterValueByName(Element elem, string paramName)
        {
            IList<Parameter> m_paramList = elem.GetParameters(paramName);

            if (m_paramList != null)
                try
                {
                    Parameter param = m_paramList[0];
                    string paramValue = param.AsValueString();
                    return paramValue;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    return null;
                }

            return "";
        }

        // REDUNDANCY ANALYSIS AND OPTIMIZED PURGE METHODS

        // REDUNDANT METHODS - REMOVE THESE:
        // - PurgeUnusedLoadedFamilies() - Same as PurgeUnusedFamilySymbols()
        // - PurgeUnusedRenderingMaterials() - Incomplete and overlaps with PurgeUnusedMaterials()

        // OPTIMIZED METHODS:

        internal static void PurgeUnusedFamilySymbols(Document curDoc)
        {
            List<FamilySymbol> symbols = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList<FamilySymbol>();

            foreach (var symbol in symbols)
            {
                if (!symbol.IsActive && !SymbolIsInUse(curDoc, symbol))

                {
                    try
                    {
                        curDoc.Delete(symbol.Id);
                    }
                    catch { }
                }
            }
        }

        private static bool SymbolIsInUse(Document curDoc, FamilySymbol symbol)
        {
            return new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Any(e => e.GetTypeId() == symbol.Id);
        }

        internal static void PurgeUnusedViewTemplates(Document curDoc)
        {
            List<View> templates = new FilteredElementCollector(curDoc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => v.IsTemplate)
            .ToList<View>();

            var usedTemplateIds = new FilteredElementCollector(curDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.ViewTemplateId != ElementId.InvalidElementId)
                .Select(v => v.ViewTemplateId)
                .ToHashSet();

            foreach (var template in templates)
            {
                if (!usedTemplateIds.Contains(template.Id))
                {
                    try
                    {
                        curDoc.Delete(template.Id);
                    }
                    catch { }
                }
            }
        }

        internal static void PurgeUnusedFilters(Document curDoc)
        {
            List<ParameterFilterElement> filters = null;
            HashSet<ElementId> usedFilterIds = null;

            try
            {
                // Collect all parameter filters with error handling
                try
                {
                    filters = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(ParameterFilterElement))
                        .Cast<ParameterFilterElement>()
                        .Where(f => f != null) // Filter out any null elements
                        .ToList<ParameterFilterElement>();
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Failed to collect parameter filters", ex);
                }

                // Collect all views with error handling
                IEnumerable<View> views = null;
                try
                {
                    views = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => v != null); // Filter out any null views
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Failed to collect views", ex);
                }

                // Build set of used filter IDs
                usedFilterIds = new HashSet<ElementId>();

                foreach (var view in views)
                {
                    if (view == null) continue;

                    try
                    {
                        // Get filters for this view
                        var viewFilters = view.GetFilters();
                        if (viewFilters != null)
                        {
                            foreach (var id in viewFilters)
                            {
                                if (id != null && id != ElementId.InvalidElementId)
                                {
                                    usedFilterIds.Add(id);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log but continue processing other views
                        // You might want to add logging here
                        System.Diagnostics.Debug.WriteLine($"Error getting filters for view {view.Name}: {ex.Message}");
                        continue;
                    }
                }

                // Delete unused filters
                if (filters != null && filters.Count > 0)
                {
                    foreach (var filter in filters)
                    {
                        if (filter == null || filter.Id == null || filter.Id == ElementId.InvalidElementId)
                            continue;

                        try
                        {
                            // Check if filter is not in use
                            if (!usedFilterIds.Contains(filter.Id))
                            {
                                // Additional check to ensure the element is still valid
                                if (curDoc.GetElement(filter.Id) != null)
                                {
                                    curDoc.Delete(filter.Id);
                                }
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            // Element might already be deleted or invalid
                            System.Diagnostics.Debug.WriteLine($"Cannot delete filter {filter.Id}: Invalid element");
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            // Element might be in use by something we didn't detect
                            System.Diagnostics.Debug.WriteLine($"Cannot delete filter {filter.Id}: Element is in use");
                        }
                        catch (Exception ex)
                        {
                            // Catch any other unexpected exceptions
                            System.Diagnostics.Debug.WriteLine($"Unexpected error deleting filter {filter.Id}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Re-throw with more context
                throw new InvalidOperationException("Failed to purge unused filters", ex);
            }
        }

        internal static void PurgeUnusedMaterials(Document curDoc)
        {
            var allMaterialIds = new FilteredElementCollector(curDoc)
            .OfClass(typeof(Material))
            .ToElementIds();

            var usedMaterialIds = new HashSet<ElementId>();

            var elementCollector = new FilteredElementCollector(curDoc)
                .WhereElementIsNotElementType();

            foreach (var elem in elementCollector)
            {
                var matIds = elem.GetMaterialIds(false);
                foreach (var id in matIds)
                    usedMaterialIds.Add(id);
            }

            foreach (var matId in allMaterialIds)
            {
                if (!usedMaterialIds.Contains(matId))
                {
                    try
                    {
                        curDoc.Delete(matId);
                    }
                    catch { }
                }
            }
        }

        internal static void PurgeUnusedLinePatterns(Document curDoc)
        {
            // Check if document is valid
            if (curDoc == null)
            {
                // Silent return - no exception thrown
                System.Diagnostics.Debug.WriteLine("PurgeUnusedLinePatterns: Document is null");
                return;
            }

            IEnumerable<LinePatternElement> allPatterns = null;
            HashSet<ElementId> usedPatternIds = null;

            try
            {
                // Collect all line patterns with error handling
                try
                {
                    allPatterns = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(LinePatternElement))
                        .Cast<LinePatternElement>()
                        .Where(lpe => lpe != null &&
                               lpe.Name != null &&
                               !lpe.Name.StartsWith("<")); // Filter out system patterns
                }
                catch (Exception ex)
                {
                    // Silent handling - just log and return
                    System.Diagnostics.Debug.WriteLine($"Failed to collect line patterns: {ex.Message}");
                    return;
                }

                // Initialize the set for used pattern IDs
                usedPatternIds = new HashSet<ElementId>();

                // Check document settings and categories
                if (curDoc.Settings == null || curDoc.Settings.Categories == null)
                {
                    // Silent handling - just log and return
                    System.Diagnostics.Debug.WriteLine("Document settings or categories are not accessible");
                    return;
                }

                // Get lines category with error handling
                Category linesCategory = null;
                try
                {
                    linesCategory = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                }
                catch (Exception ex)
                {
                    // Silent handling - just log and return
                    System.Diagnostics.Debug.WriteLine($"Failed to get Lines category: {ex.Message}");
                    return;
                }

                if (linesCategory == null)
                {
                    // Silent handling - just log and return
                    System.Diagnostics.Debug.WriteLine("Lines category not found in document");
                    return;
                }

                // Check if we have an active view for getting overrides
                if (curDoc.ActiveView == null)
                {
                    System.Diagnostics.Debug.WriteLine("Warning: No active view available for checking category overrides");
                }

                // Process subcategories
                if (linesCategory.SubCategories != null)
                {
                    foreach (Category subcat in linesCategory.SubCategories)
                    {
                        if (subcat == null || subcat.Id == null || subcat.Id == ElementId.InvalidElementId)
                            continue;

                        try
                        {
                            // Get graphics style
                            Element element = curDoc.GetElement(subcat.Id);
                            if (element == null) continue;

                            GraphicsStyle gs = element as GraphicsStyle;
                            if (gs != null && gs.Id != null && gs.Id != ElementId.InvalidElementId)
                            {
                                try
                                {
                                    // Only try to get overrides if we have an active view
                                    if (curDoc.ActiveView != null)
                                    {
                                        // Verify the graphics style has a valid category
                                        if (gs.GraphicsStyleCategory != null &&
                                            gs.GraphicsStyleCategory.Id != null &&
                                            gs.GraphicsStyleCategory.Id != ElementId.InvalidElementId)
                                        {
                                            OverrideGraphicSettings ogs = curDoc.ActiveView.GetCategoryOverrides(gs.GraphicsStyleCategory.Id);

                                            if (ogs != null)
                                            {
                                                // Check projection line pattern
                                                if (ogs.ProjectionLinePatternId != null &&
                                                    ogs.ProjectionLinePatternId != ElementId.InvalidElementId)
                                                {
                                                    usedPatternIds.Add(ogs.ProjectionLinePatternId);
                                                }

                                                // Also check cut line pattern if you want to be thorough
                                                if (ogs.CutLinePatternId != null &&
                                                    ogs.CutLinePatternId != ElementId.InvalidElementId)
                                                {
                                                    usedPatternIds.Add(ogs.CutLinePatternId);
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                                {
                                    // Category might not support overrides
                                    System.Diagnostics.Debug.WriteLine($"Cannot get overrides for category {gs.Name}");
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error processing graphics style {gs.Name}: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log but continue processing other subcategories
                            System.Diagnostics.Debug.WriteLine($"Error processing subcategory: {ex.Message}");
                            continue;
                        }
                    }
                }

                // Delete unused patterns
                if (allPatterns != null)
                {
                    foreach (var pattern in allPatterns)
                    {
                        if (pattern == null || pattern.Id == null || pattern.Id == ElementId.InvalidElementId)
                            continue;

                        try
                        {
                            // Check if pattern is not in use
                            if (!usedPatternIds.Contains(pattern.Id))
                            {
                                // Additional check to ensure the element is still valid
                                if (curDoc.GetElement(pattern.Id) != null)
                                {
                                    curDoc.Delete(pattern.Id);
                                    System.Diagnostics.Debug.WriteLine($"Deleted unused line pattern: {pattern.Name}");
                                }
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            // Element might already be deleted or invalid
                            System.Diagnostics.Debug.WriteLine($"Cannot delete line pattern {pattern.Name}: Invalid element");
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            // Element might be in use by something we didn't detect
                            System.Diagnostics.Debug.WriteLine($"Cannot delete line pattern {pattern.Name}: Element is in use");
                        }
                        catch (Exception ex)
                        {
                            // Catch any other unexpected exceptions
                            System.Diagnostics.Debug.WriteLine($"Unexpected error deleting line pattern {pattern.Name}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Silent handling - just log the error
                System.Diagnostics.Debug.WriteLine($"Failed to purge unused line patterns: {ex.Message}");
                // Method completes without throwing exception
            }
        }

        internal static void PurgeUnusedFillPatterns(Document curDoc)
        {
            IEnumerable<FillPatternElement> patterns = null;
            HashSet<ElementId> usedPatternIds = null;

            try
            {
                // Collect all fill patterns with error handling
                try
                {
                    patterns = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(FillPatternElement))
                        .Cast<FillPatternElement>()
                        .Where(p => p != null)
                        .ToList(); // Convert to List to avoid iterator issues
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Failed to collect fill patterns", ex);
                }

                // Initialize the set for used pattern IDs
                usedPatternIds = new HashSet<ElementId>();

                // Collect views with error handling
                IEnumerable<View> views = null;
                try
                {
                    views = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => v != null && !v.IsTemplate)
                        .ToList(); // Convert to List to avoid iterator issues
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Failed to collect views", ex);
                }

                // Process each view
                foreach (var view in views)
                {
                    if (view == null || view.Id == null || view.Id == ElementId.InvalidElementId)
                        continue;

                    try
                    {
                        // Check if view is valid for element collection
                        if (!view.IsValidObject)
                        {
                            System.Diagnostics.Debug.WriteLine($"Skipping invalid view: {view.Name}");
                            continue;
                        }

                        // Collect elements in the view
                        ICollection<Element> elements = null;
                        try
                        {
                            elements = new FilteredElementCollector(curDoc, view.Id)
                                .WhereElementIsNotElementType()
                                .ToElements();
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            // View might not support element collection
                            System.Diagnostics.Debug.WriteLine($"Cannot collect elements from view: {view.Name}");
                            continue;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error collecting elements from view {view.Name}: {ex.Message}");
                            continue;
                        }

                        if (elements == null || elements.Count == 0)
                            continue;

                        // Process each element in the view
                        foreach (var elem in elements)
                        {
                            if (elem == null || elem.Id == null || elem.Id == ElementId.InvalidElementId)
                                continue;

                            try
                            {
                                // Get element overrides
                                OverrideGraphicSettings ogs = view.GetElementOverrides(elem.Id);

                                if (ogs != null)
                                {
                                    // Check surface foreground pattern
                                    if (ogs.SurfaceForegroundPatternId != null &&
                                        ogs.SurfaceForegroundPatternId != ElementId.InvalidElementId)
                                    {
                                        usedPatternIds.Add(ogs.SurfaceForegroundPatternId);
                                    }

                                    // Check surface background pattern
                                    if (ogs.SurfaceBackgroundPatternId != null &&
                                        ogs.SurfaceBackgroundPatternId != ElementId.InvalidElementId)
                                    {
                                        usedPatternIds.Add(ogs.SurfaceBackgroundPatternId);
                                    }

                                    // Check cut foreground pattern
                                    if (ogs.CutForegroundPatternId != null &&
                                        ogs.CutForegroundPatternId != ElementId.InvalidElementId)
                                    {
                                        usedPatternIds.Add(ogs.CutForegroundPatternId);
                                    }

                                    // Check cut background pattern
                                    if (ogs.CutBackgroundPatternId != null &&
                                        ogs.CutBackgroundPatternId != ElementId.InvalidElementId)
                                    {
                                        usedPatternIds.Add(ogs.CutBackgroundPatternId);
                                    }
                                }
                            }
                            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                            {
                                // Element might not support overrides in this view
                                continue;
                            }
                            catch (Autodesk.Revit.Exceptions.ArgumentException)
                            {
                                // Element might be invalid or deleted
                                continue;
                            }
                            catch (Exception ex)
                            {
                                // Log but continue processing other elements
                                System.Diagnostics.Debug.WriteLine($"Error getting overrides for element {elem.Id}: {ex.Message}");
                                continue;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log but continue processing other views
                        System.Diagnostics.Debug.WriteLine($"Error processing view {view.Name}: {ex.Message}");
                        continue;
                    }
                }

                // Also check materials for fill patterns (additional safety)
                try
                {
                    var materials = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(Material))
                        .Cast<Material>()
                        .Where(m => m != null)
                        .ToList(); // Convert to List to avoid iterator issues

                    foreach (var material in materials)
                    {
                        try
                        {
                            if (material.SurfaceForegroundPatternId != null &&
                                material.SurfaceForegroundPatternId != ElementId.InvalidElementId)
                            {
                                usedPatternIds.Add(material.SurfaceForegroundPatternId);
                            }

                            if (material.SurfaceBackgroundPatternId != null &&
                                material.SurfaceBackgroundPatternId != ElementId.InvalidElementId)
                            {
                                usedPatternIds.Add(material.SurfaceBackgroundPatternId);
                            }

                            if (material.CutForegroundPatternId != null &&
                                material.CutForegroundPatternId != ElementId.InvalidElementId)
                            {
                                usedPatternIds.Add(material.CutForegroundPatternId);
                            }

                            if (material.CutBackgroundPatternId != null &&
                                material.CutBackgroundPatternId != ElementId.InvalidElementId)
                            {
                                usedPatternIds.Add(material.CutBackgroundPatternId);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error checking material patterns: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error collecting materials: {ex.Message}");
                }

                // Collect patterns to delete first (to avoid iterator modification issues)
                List<ElementId> patternsToDelete = new List<ElementId>();

                if (patterns != null)
                {
                    foreach (var pattern in patterns)
                    {
                        if (pattern == null || pattern.Id == null || pattern.Id == ElementId.InvalidElementId)
                            continue;

                        try
                        {
                            // Check if pattern is not in use
                            if (!usedPatternIds.Contains(pattern.Id))
                            {
                                // Additional check to ensure the element is still valid
                                if (curDoc.GetElement(pattern.Id) != null)
                                {
                                    patternsToDelete.Add(pattern.Id);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error checking pattern {pattern.Name}: {ex.Message}");
                        }
                    }
                }

                // Now delete the collected patterns
                foreach (var patternId in patternsToDelete)
                {
                    try
                    {
                        var pattern = curDoc.GetElement(patternId) as FillPatternElement;
                        if (pattern != null)
                        {
                            curDoc.Delete(patternId);
                            System.Diagnostics.Debug.WriteLine($"Deleted unused fill pattern: {pattern.Name}");
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException)
                    {
                        // Element might already be deleted or invalid
                        System.Diagnostics.Debug.WriteLine($"Cannot delete fill pattern {patternId}: Invalid element");
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // Element might be in use by something we didn't detect
                        System.Diagnostics.Debug.WriteLine($"Cannot delete fill pattern {patternId}: Element is in use");
                    }
                    catch (Exception ex)
                    {
                        // Catch any other unexpected exceptions
                        System.Diagnostics.Debug.WriteLine($"Unexpected error deleting fill pattern {patternId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Re-throw with more context
                throw new InvalidOperationException("Failed to purge unused fill patterns", ex);
            }
        }

        internal static void PurgeUnusedGroups(Document curDoc)
        {
            // Collect all groups and convert to List to avoid iterator issues
            var groups = new FilteredElementCollector(curDoc)
                .OfClass(typeof(GroupType))
                .Cast<GroupType>()
                .ToList();

            // Collect groups to delete first
            List<ElementId> groupsToDelete = new List<ElementId>();

            foreach (var group in groups)
            {
                if (group.Groups.Size == 0)
                {
                    groupsToDelete.Add(group.Id);
                }
            }

            // Now delete the collected groups
            foreach (var groupId in groupsToDelete)
            {
                try
                {
                    curDoc.Delete(groupId);
                }
                catch { }
            }
        }

        // Additional purge methods to add to your Utils class

        internal static void PurgeUnusedTextStyles(Document curDoc)
        {
            try
            {
                // Collect all text note types and convert to List
                var textStyles = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .ToList();

                // Collect used text style IDs
                HashSet<ElementId> usedStyleIds = new HashSet<ElementId>();

                // Check text notes for used styles
                var textNotes = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(TextNote))
                    .Cast<TextNote>()
                    .ToList();

                foreach (var textNote in textNotes)
                {
                    if (textNote?.GetTypeId() != null && textNote.GetTypeId() != ElementId.InvalidElementId)
                    {
                        usedStyleIds.Add(textNote.GetTypeId());
                    }
                }

                // Collect styles to delete
                List<ElementId> stylesToDelete = new List<ElementId>();

                foreach (var style in textStyles)
                {
                    if (style?.Id != null && !usedStyleIds.Contains(style.Id))
                    {
                        stylesToDelete.Add(style.Id);
                    }
                }

                // Delete unused styles
                foreach (var styleId in stylesToDelete)
                {
                    try
                    {
                        curDoc.Delete(styleId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not delete text style {styleId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging text styles: {ex.Message}");
            }
        }

        internal static void PurgeUnusedDimensionStyles(Document curDoc)
        {
            try
            {
                // Collect all dimension types
                var dimensionTypes = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(DimensionType))
                    .Cast<DimensionType>()
                    .ToList();

                // Collect used dimension type IDs
                HashSet<ElementId> usedTypeIds = new HashSet<ElementId>();

                // Check dimensions for used types
                var dimensions = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(Dimension))
                    .Cast<Dimension>()
                    .ToList();

                foreach (var dimension in dimensions)
                {
                    if (dimension?.GetTypeId() != null && dimension.GetTypeId() != ElementId.InvalidElementId)
                    {
                        usedTypeIds.Add(dimension.GetTypeId());
                    }
                }

                // Collect types to delete
                List<ElementId> typesToDelete = new List<ElementId>();

                foreach (var dimType in dimensionTypes)
                {
                    if (dimType?.Id != null && !usedTypeIds.Contains(dimType.Id))
                    {
                        typesToDelete.Add(dimType.Id);
                    }
                }

                // Delete unused types
                foreach (var typeId in typesToDelete)
                {
                    try
                    {
                        curDoc.Delete(typeId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not delete dimension type {typeId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging dimension styles: {ex.Message}");
            }
        }

        internal static void PurgeUnusedLineStyles(Document curDoc)
        {
            try
            {
                // Get line styles (subcategory of Lines category)
                Category lineCategory = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                if (lineCategory?.SubCategories == null) return;

                // Collect used line style IDs
                HashSet<ElementId> usedLineStyleIds = new HashSet<ElementId>();

                // Check detail lines for used styles
                var detailLines = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(DetailLine))
                    .Cast<DetailLine>()
                    .ToList();

                foreach (var line in detailLines)
                {
                    if (line?.LineStyle?.Id != null)
                    {
                        usedLineStyleIds.Add(line.LineStyle.Id);
                    }
                }

                // Check model lines for used styles
                var modelLines = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(ModelLine))
                    .Cast<ModelLine>()
                    .ToList();

                foreach (var line in modelLines)
                {
                    if (line?.LineStyle?.Id != null)
                    {
                        usedLineStyleIds.Add(line.LineStyle.Id);
                    }
                }

                // Collect line styles to delete
                List<ElementId> stylesToDelete = new List<ElementId>();

                foreach (Category subCat in lineCategory.SubCategories)
                {
                    if (subCat?.Id != null && !usedLineStyleIds.Contains(subCat.Id))
                    {
                        stylesToDelete.Add(subCat.Id);
                    }
                }

                // Delete unused line styles
                foreach (var styleId in stylesToDelete)
                {
                    try
                    {
                        curDoc.Delete(styleId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not delete line style {styleId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging line styles: {ex.Message}");
            }
        }

        internal static void PurgeUnusedAnnotationSymbols(Document curDoc)
        {
            try
            {
                // Collect all annotation symbol types
                var annotationTypes = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(AnnotationSymbolType))
                    .Cast<AnnotationSymbolType>()
                    .ToList();

                // Collect used annotation type IDs
                HashSet<ElementId> usedTypeIds = new HashSet<ElementId>();

                // Check annotation symbols for used types
                var annotations = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(AnnotationSymbol))
                    .Cast<AnnotationSymbol>()
                    .ToList();

                foreach (var annotation in annotations)
                {
                    if (annotation?.GetTypeId() != null && annotation.GetTypeId() != ElementId.InvalidElementId)
                    {
                        usedTypeIds.Add(annotation.GetTypeId());
                    }
                }

                // Collect types to delete
                List<ElementId> typesToDelete = new List<ElementId>();

                foreach (var annoType in annotationTypes)
                {
                    if (annoType?.Id != null && !usedTypeIds.Contains(annoType.Id))
                    {
                        typesToDelete.Add(annoType.Id);
                    }
                }

                // Delete unused types
                foreach (var typeId in typesToDelete)
                {
                    try
                    {
                        curDoc.Delete(typeId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not delete annotation type {typeId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging annotation symbols: {ex.Message}");
            }
        }

        internal static void PurgeUnusedLoadedFamilies(Document curDoc)
        {
            try
            {
                // Collect all family symbols (types)
                var familySymbols = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                // Collect used family symbol IDs
                HashSet<ElementId> usedSymbolIds = new HashSet<ElementId>();

                // Check family instances for used symbols
                var familyInstances = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                foreach (var instance in familyInstances)
                {
                    if (instance?.GetTypeId() != null && instance.GetTypeId() != ElementId.InvalidElementId)
                    {
                        usedSymbolIds.Add(instance.GetTypeId());
                    }
                }

                // Collect symbols to delete
                List<ElementId> symbolsToDelete = new List<ElementId>();

                foreach (var symbol in familySymbols)
                {
                    if (symbol?.Id != null && !usedSymbolIds.Contains(symbol.Id))
                    {
                        symbolsToDelete.Add(symbol.Id);
                    }
                }

                // Delete unused symbols
                foreach (var symbolId in symbolsToDelete)
                {
                    try
                    {
                        curDoc.Delete(symbolId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not delete family symbol {symbolId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging loaded families: {ex.Message}");
            }
        }

        internal static void PurgeUnusedAppearanceAssets(Document curDoc)
        {
            try
            {
                // Collect all appearance assets
                var appearanceAssets = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(AppearanceAssetElement))
                    .Cast<AppearanceAssetElement>()
                    .ToList();

                // Collect used appearance asset IDs
                HashSet<ElementId> usedAssetIds = new HashSet<ElementId>();

                // Check materials for used appearance assets
                var materials = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(Material))
                    .Cast<Material>()
                    .ToList();

                foreach (var material in materials)
                {
                    try
                    {
                        var appearanceAssetId = material.AppearanceAssetId;
                        if (appearanceAssetId != null && appearanceAssetId != ElementId.InvalidElementId)
                        {
                            usedAssetIds.Add(appearanceAssetId);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error checking material appearance asset: {ex.Message}");
                    }
                }

                // Collect assets to delete
                List<ElementId> assetsToDelete = new List<ElementId>();

                foreach (var asset in appearanceAssets)
                {
                    if (asset?.Id != null && !usedAssetIds.Contains(asset.Id))
                    {
                        assetsToDelete.Add(asset.Id);
                    }
                }

                // Delete unused assets
                foreach (var assetId in assetsToDelete)
                {
                    try
                    {
                        curDoc.Delete(assetId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not delete appearance asset {assetId}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging appearance assets: {ex.Message}");
            }
        }

        internal static void PurgeUnusedRenderingMaterials(Document curDoc)
        {
            try
            {
                // This method attempts to purge unused rendering materials/assets
                // Note: Some rendering materials might be harder to detect usage for

                var renderingElements = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(Material))
                    .Cast<Material>()
                    .Where(m => m != null)
                    .ToList();

                // Collect used material IDs from elements
                HashSet<ElementId> usedMaterialIds = new HashSet<ElementId>();

                // Check all elements for material usage
                var allElements = new FilteredElementCollector(curDoc)
                    .WhereElementIsNotElementType()
                    .ToList();

                foreach (var element in allElements)
                {
                    try
                    {
                        // Try to get material IDs from the element
                        var materialIds = element.GetMaterialIds(false);
                        foreach (var matId in materialIds)
                        {
                            if (matId != null && matId != ElementId.InvalidElementId)
                            {
                                usedMaterialIds.Add(matId);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Some elements might not support GetMaterialIds
                        continue;
                    }
                }

                // Also check element types for materials
                var elementTypes = new FilteredElementCollector(curDoc)
                    .WhereElementIsElementType()
                    .ToList();

                foreach (var elementType in elementTypes)
                {
                    try
                    {
                        var materialIds = elementType.GetMaterialIds(false);
                        foreach (var matId in materialIds)
                        {
                            if (matId != null && matId != ElementId.InvalidElementId)
                            {
                                usedMaterialIds.Add(matId);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                // This method focuses on identifying unused materials more conservatively
                // since materials can be referenced in many ways
                System.Diagnostics.Debug.WriteLine($"Found {usedMaterialIds.Count} materials in use");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error purging rendering materials: {ex.Message}");
            }
        }

        internal static bool SaveToLuisFolder(Document doc, string fileName = null)
        {
            try
            {
                // Target directory
                string targetDirectory = @"S:\Shared Folders\Lifestyle USA Design\LGI Homes\!Luis";

                // Generate filename if not provided
                if (string.IsNullOrEmpty(fileName))
                {
                    string originalName = Path.GetFileNameWithoutExtension(doc.Title);

                    // Extract plan name and region code (remove lot dimensions and anything after)
                    // Example: "Torres(R)-CTX(50-5-29'11)" should become "Torres(R)-CTX"

                    // Find the pattern: look for the last hyphen followed by letters, then a parenthesis
                    // This should identify the region code section
                    int regionCodeStart = -1;
                    for (int i = originalName.Length - 1; i >= 0; i--)
                    {
                        if (originalName[i] == '-')
                        {
                            // Check if this is followed by letters (region code)
                            int letterStart = i + 1;
                            int letterEnd = letterStart;
                            while (letterEnd < originalName.Length && char.IsLetter(originalName[letterEnd]))
                            {
                                letterEnd++;
                            }

                            if (letterEnd > letterStart && letterEnd < originalName.Length && originalName[letterEnd] == '(')
                            {
                                // Found region code pattern, keep everything up to the opening parenthesis after region
                                originalName = originalName.Substring(0, letterEnd);
                                break;
                            }
                        }
                    }

                    fileName = $"{originalName}.rvt";
                }

                // Ensure filename has .rvt extension
                if (!fileName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                {
                    fileName += ".rvt";
                }

                string fullPath = Path.Combine(targetDirectory, fileName);

                // Ensure the directory exists
                if (!Directory.Exists(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                // Configure save options
                SaveAsOptions saveAsOptions = new SaveAsOptions();
                saveAsOptions.OverwriteExistingFile = true;

                // Perform the Save As operation
                doc.SaveAs(fullPath, saveAsOptions);

                // Optional: Show success message
                TaskDialog successDialog = new TaskDialog("Save Complete");
                successDialog.MainContent = $"File successfully saved to:\n{fullPath}";
                successDialog.Show();

                return true;
            }
            catch (Exception ex)
            {
                // Show error message
                TaskDialog errorDialog = new TaskDialog("Save Error");
                errorDialog.MainContent = $"Failed to save file:\n{ex.Message}";
                errorDialog.Show();

                return false;
            }
        }
    }
}