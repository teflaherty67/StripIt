using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace StripIt
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

                if (viewCat == catName)
                    m_returnList.Add(curView);
            }

            return m_returnList;
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

        internal static List<ViewSchedule> GetAllSchedules(Document curDoc)
        {
            List<ViewSchedule> schedList = new List<ViewSchedule>();

            FilteredElementCollector curCollector = new FilteredElementCollector(curDoc);
            curCollector.OfClass(typeof(ViewSchedule));

            //loop through views and check if schedule - if so then put into schedule list
            foreach (ViewSchedule curView in curCollector)
            {
                if (curView.ViewType == ViewType.Schedule)
                {
                    schedList.Add((ViewSchedule)curView);
                }
            }

            return schedList;
        }

        internal static void PurgeUnusedFamilySymbols(Document curDoc)
        {
            var symbols = new FilteredElementCollector(curDoc)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>();

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
            var templates = new FilteredElementCollector(curDoc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => v.IsTemplate);

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
            var filters = new FilteredElementCollector(curDoc)
            .OfClass(typeof(ParameterFilterElement))
            .Cast<ParameterFilterElement>();

            var views = new FilteredElementCollector(curDoc)
                .OfClass(typeof(View))
                .Cast<View>();

            var usedFilterIds = new HashSet<ElementId>();

            foreach (var view in views)
            {
                foreach (var id in view.GetFilters())
                {
                    usedFilterIds.Add(id);
                }
            }

            foreach (var filter in filters)
            {
                if (!usedFilterIds.Contains(filter.Id))
                {
                    try
                    {
                        curDoc.Delete(filter.Id);
                    }
                    catch { }
                }
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
            var allPatterns = new FilteredElementCollector(curDoc)
            .OfClass(typeof(LinePatternElement))
            .Cast<LinePatternElement>()
            .Where(lpe => !lpe.Name.StartsWith("<"));

            var usedPatternIds = new HashSet<ElementId>();

            Category linesCategory = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            foreach (Category subcat in linesCategory.SubCategories)
            {
                GraphicsStyle gs = curDoc.GetElement(subcat.Id) as GraphicsStyle;
                if (gs != null)
                {
                    OverrideGraphicSettings ogs = curDoc.GetElement(gs.Id).GetType() == typeof(GraphicsStyle)
                        ? curDoc.ActiveView.GetCategoryOverrides(gs.GraphicsStyleCategory.Id)
                        : null;

                    if (ogs != null && ogs.ProjectionLinePatternId != ElementId.InvalidElementId)
                        usedPatternIds.Add(ogs.ProjectionLinePatternId);
                }
            }

            foreach (var pattern in allPatterns)
            {
                if (!usedPatternIds.Contains(pattern.Id))
                {
                    try
                    {
                        curDoc.Delete(pattern.Id);
                    }
                    catch { }
                }
            }
        }

        internal static void PurgeUnusedFillPatterns(Document curDoc)
        {
            var patterns = new FilteredElementCollector(curDoc)
        .OfClass(typeof(FillPatternElement))
        .Cast<FillPatternElement>();

            var usedPatternIds = new HashSet<ElementId>();

            // Heuristic approach: try to gather fill patterns from view overrides
            var views = new FilteredElementCollector(curDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate);

            foreach (var view in views)
            {
                var elements = new FilteredElementCollector(curDoc, view.Id)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var elem in elements)
                {
                    OverrideGraphicSettings ogs = view.GetElementOverrides(elem.Id);
                    if (ogs != null)
                    {
                        if (ogs.SurfaceForegroundPatternId != ElementId.InvalidElementId)
                            usedPatternIds.Add(ogs.SurfaceForegroundPatternId);

                        if (ogs.CutForegroundPatternId != ElementId.InvalidElementId)
                            usedPatternIds.Add(ogs.CutForegroundPatternId);
                    }
                }
            }

            foreach (var pattern in patterns)
            {
                if (!usedPatternIds.Contains(pattern.Id))
                {
                    try
                    {
                        curDoc.Delete(pattern.Id);
                    }
                    catch { }
                }
            }
        }

        internal static void PurgeUnusedGroups(Document curDoc)
        {
            var groups = new FilteredElementCollector(curDoc)
           .OfClass(typeof(GroupType))
           .Cast<GroupType>();

            foreach (var group in groups)
            {
                if (group.Groups.Size == 0)
                {
                    try
                    {
                        curDoc.Delete(group.Id);
                    }
                    catch { }
                }
            }
        }

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
            List<ParameterFilterElement> filters = new FilteredElementCollector(curDoc)
            .OfClass(typeof(ParameterFilterElement))
            .Cast<ParameterFilterElement>()
            .ToList<ParameterFilterElement>();

            var views = new FilteredElementCollector(curDoc)
                .OfClass(typeof(View))
                .Cast<View>();

            var usedFilterIds = new HashSet<ElementId>();

            foreach (var view in views)
            {
                foreach (var id in view.GetFilters())
                {
                    usedFilterIds.Add(id);
                }
            }

            foreach (var filter in filters)
            {
                if (!usedFilterIds.Contains(filter.Id))
                {
                    try
                    {
                        curDoc.Delete(filter.Id);
                    }
                    catch { }
                }
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
            var allPatterns = new FilteredElementCollector(curDoc)
            .OfClass(typeof(LinePatternElement))
            .Cast<LinePatternElement>()
            .Where(lpe => !lpe.Name.StartsWith("<"));

            var usedPatternIds = new HashSet<ElementId>();

            Category linesCategory = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            foreach (Category subcat in linesCategory.SubCategories)
            {
                GraphicsStyle gs = curDoc.GetElement(subcat.Id) as GraphicsStyle;
                if (gs != null)
                {
                    OverrideGraphicSettings ogs = curDoc.GetElement(gs.Id).GetType() == typeof(GraphicsStyle)
                        ? curDoc.ActiveView.GetCategoryOverrides(gs.GraphicsStyleCategory.Id)
                        : null;

                    if (ogs != null && ogs.ProjectionLinePatternId != ElementId.InvalidElementId)
                        usedPatternIds.Add(ogs.ProjectionLinePatternId);
                }
            }

            foreach (var pattern in allPatterns)
            {
                if (!usedPatternIds.Contains(pattern.Id))
                {
                    try
                    {
                        curDoc.Delete(pattern.Id);
                    }
                    catch { }
                }
            }
        }

        internal static void PurgeUnusedFillPatterns(Document curDoc)
        {
            var patterns = new FilteredElementCollector(curDoc)
        .OfClass(typeof(FillPatternElement))
        .Cast<FillPatternElement>();

            var usedPatternIds = new HashSet<ElementId>();

            // Heuristic approach: try to gather fill patterns from view overrides
            var views = new FilteredElementCollector(curDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate);

            foreach (var view in views)
            {
                var elements = new FilteredElementCollector(curDoc, view.Id)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var elem in elements)
                {
                    OverrideGraphicSettings ogs = view.GetElementOverrides(elem.Id);
                    if (ogs != null)
                    {
                        if (ogs.SurfaceForegroundPatternId != ElementId.InvalidElementId)
                            usedPatternIds.Add(ogs.SurfaceForegroundPatternId);

                        if (ogs.CutForegroundPatternId != ElementId.InvalidElementId)
                            usedPatternIds.Add(ogs.CutForegroundPatternId);
                    }
                }
            }

            foreach (var pattern in patterns)
            {
                if (!usedPatternIds.Contains(pattern.Id))
                {
                    try
                    {
                        curDoc.Delete(pattern.Id);
                    }
                    catch { }
                }
            }
        }

        internal static void PurgeUnusedGroups(Document curDoc)
        {
            var groups = new FilteredElementCollector(curDoc)
           .OfClass(typeof(GroupType))
           .Cast<GroupType>();

            foreach (var group in groups)
            {
                if (group.Groups.Size == 0)
                {
                    try
                    {
                        curDoc.Delete(group.Id);
                    }
                    catch { }
                }
            }
        }
    }
}
