using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class CategoryUtil {
        internal class Data {
            public string label;
            public Outlook.OlCategoryColor color;
        }

        private static void UpdateCategory(Outlook.Application app, string name, Outlook.OlCategoryColor color) {
            Outlook.Categories allCats = app.GetNamespace("MAPI").Categories;
            bool found = false;
            for (int i = 1; i <= allCats.Count; i++) {
                if (allCats[i].Name == name) {
                    allCats[i].Color = color;
                    found = true;
                    break;
                }
            }
            if (!found) {
                app.GetNamespace("MAPI").Categories.Add(name, color, null);
            }
        }

        internal static void UpdateCategories(Outlook.Application app) {
            foreach (CategoryUtil.Data data in Config.Categories) {
                UpdateCategory(app, data.label, data.color);
            }
        }

        internal static Outlook.OlCategoryColor GetColor(string color) {
            if (Enum.TryParse(color, out Outlook.OlCategoryColor colorEnum) == true) {
                return colorEnum;
            } else {
                return Outlook.OlCategoryColor.olCategoryColorNone;
            }
        }
    }
}
