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
                if (!symbol.IsActive && !symbol.IsInUse)
                {
                    try
                    {
                        curDoc.Delete(symbol.Id);
                    }
                    catch { }
                }
            }
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
            .Where(lpe => !lpe.IsBuiltInPattern);

            var usedPatternIds = new HashSet<ElementId>();

            Category linesCategory = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            foreach (Category subcat in linesCategory.SubCategories)
            {
                if (subcat != null && subcat.LinePatternId != ElementId.InvalidElementId)
                    usedPatternIds.Add(subcat.LinePatternId);
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
            var allPatterns = new FilteredElementCollector(curDoc)
            .OfClass(typeof(FillPatternElement))
            .Cast<FillPatternElement>();

            var usedPatternIds = new HashSet<ElementId>();

            foreach (Category cat in curDoc.Settings.Categories)
            {
                if (cat == null) continue;

                if (!cat.CutPatternId.Equals(ElementId.InvalidElementId))
                    usedPatternIds.Add(cat.CutPatternId);

                if (!cat.MaterialId.Equals(ElementId.InvalidElementId))
                {
                    var mat = curDoc.GetElement(cat.MaterialId) as Material;
                    if (mat != null && mat.SurfacePatternId != ElementId.InvalidElementId)
                        usedPatternIds.Add(mat.SurfacePatternId);
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
