using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class ConvertChineseUtil {
        internal enum ChineseType {
            SimplifiedChinese,
            TraditionalChinese
        }

        private const int LOCALE_SYSTEM_DEFAULT = 0x0800;
        private const int LCMAP_SIMPLIFIED_CHINESE = 0x02000000;
        private const int LCMAP_TRADITIONAL_CHINESE = 0x04000000;

        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int LCMapString(int Locale, int dwMapFlags, string lpSrcStr, int cchSrc, [Out] string lpDestStr, int cchDest);

        private static string ConvertToTraditionalChinese(string source) {
            String target = new String(' ', source.Length);
            int ret = LCMapString(LOCALE_SYSTEM_DEFAULT, LCMAP_TRADITIONAL_CHINESE, source, source.Length, target, source.Length);
            return target;
        }

        private static string ConvertToSimplifiedChinese(string source) {
            String target = new String(' ', source.Length);
            int ret = LCMapString(LOCALE_SYSTEM_DEFAULT, LCMAP_SIMPLIFIED_CHINESE, source, source.Length, target, source.Length);
            return target;
        }

        private static void ConvertEmailChinese(Outlook.MailItem mailItem, ChineseType chineseType) {
            Debug.Assert(mailItem != null);

            switch (chineseType) {
                case ChineseType.SimplifiedChinese: mailItem.HTMLBody = ConvertToSimplifiedChinese(mailItem.HTMLBody); break;
                case ChineseType.TraditionalChinese: mailItem.HTMLBody = ConvertToTraditionalChinese(mailItem.HTMLBody); break;
            }

            mailItem.Save();
        }

        private static bool IsDepartmentUsingSimplfieidChinese(string dep) {
            foreach (string ChineseDept in Config.SimplifiedChineseDepts) {
                if (dep.StartsWith(ChineseDept)) {
                    return true;
                }
            }
            return false;
        }

        private static bool IsDepartmentUsingTraditionalChinese(string dep) {
            foreach (string ChineseDept in Config.TraditionalChineseDepts) {
                if (dep.StartsWith(ChineseDept)) {
                    return true;
                }
            }
            return false;
        }

        private static bool IsAllDepartmentUsingSimplfiedChinese(List<string> deps) {
            foreach (string dep in deps) {
                if ((dep != null) && (IsDepartmentUsingSimplfieidChinese(dep) == false)) {
                    return false;
                }
            }
            return true;
        }

        private static bool IsAllDepartmentUsingTraditionalChinese(List<string> deps) {
            foreach (string dep in deps) {
                if ((dep != null) && (IsDepartmentUsingTraditionalChinese(dep) == false)) {
                    return false;
                }
            }
            return true;
        }

        private static List<string> GetAllToRecipientDepartment(Outlook.MailItem mailItem) {
            Debug.Assert(mailItem != null);

            List<string> result = new List<string>();
            foreach (Outlook.Recipient recip in mailItem.Recipients) {
                if (recip.Type == (int)Outlook.OlMailRecipientType.olTo) {
                    if (recip.AddressEntry.GetExchangeUser() != null) {
                        result.Add(recip.AddressEntry.GetExchangeUser().Department);
                    }
                }
            }
            return result;
        }

        internal static void ConvertEmailChineseAccordingToRecipientDept(Outlook.MailItem mailItem) {
            Debug.Assert(mailItem != null);

            List<string> deps = GetAllToRecipientDepartment(mailItem);
            if (IsAllDepartmentUsingSimplfiedChinese(deps) == true) {
                ConvertEmailChinese(mailItem, ChineseType.SimplifiedChinese);
            } else if (IsAllDepartmentUsingTraditionalChinese(deps) == true) {
                ConvertEmailChinese(mailItem, ChineseType.TraditionalChinese);
            } else {
                switch (Config.DefaultChinese) {
                    case ChineseType.SimplifiedChinese: ConvertEmailChinese(mailItem, ChineseType.SimplifiedChinese); break;
                    case ChineseType.TraditionalChinese: ConvertEmailChinese(mailItem, ChineseType.TraditionalChinese); break;
                }
            }
        }
    }
}
