using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace wei_outlook_add_in {
    class Config {
        internal static bool AutoBackupEmailFromMe = true;
        internal static string MyEmailAddress = null;
        internal static string EmailBackupPath = null;
        internal static string AttachmentBackupPath = null;

        internal static bool EnableAutoBcc = false;
        internal static string AutoBccEmailAddress = null;

        internal static int Zoom = -1;

        internal static List<CategoryUtil.Data> Categories = new List<CategoryUtil.Data>();

        internal static List<FixedReplyUtil.Data> FixedReplies = new List<FixedReplyUtil.Data>();

        internal static ConvertChineseUtil.ChineseType DefaultChinese;
        internal static List<string> TraditionalChineseDepts = new List<string>();
        internal static List<string> SimplifiedChineseDepts = new List<string>();

        internal static void ReadFromFile() {
            string userProfileFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string configPath = userProfileFolder + @"\wei-outlook-add-in\config.xml";
            if (File.Exists(configPath) == false) {
                throw new Exception("config.xml not exist");
            }
            XDocument xmlDocument = XDocument.Load(configPath);

            XElement AutoBackupEmailFromMeElement = xmlDocument.Root.Element("AutoBackupEmailFromMe");
            if (AutoBackupEmailFromMeElement != null) {
                if (AutoBackupEmailFromMeElement.Value == "false") {
                    AutoBackupEmailFromMe = false;
                }
            }
            if (AutoBackupEmailFromMe == true) {
                XElement MyEmailAddressElement = xmlDocument.Root.Element("MyEmailAddress");
                if (MyEmailAddressElement == null) {
                    throw new Exception("MyEmailAddress not set");
                }
                MyEmailAddress = MyEmailAddressElement.Value;
                if (MyEmailAddress == "") {
                    throw new Exception("MyEmailAddress not set");
                }
            }

            XElement EmailBackupPathElement = xmlDocument.Root.Element("EmailBackupPath");
            if (EmailBackupPathElement == null) {
                throw new Exception("EmailBackupPath not set");
            }
            EmailBackupPath = EmailBackupPathElement.Value;
            if (Directory.Exists(EmailBackupPath) == false) {
                Directory.CreateDirectory(EmailBackupPath);
            }

            XElement AttachmentBackupPathElement = xmlDocument.Root.Element("AttachmentBackupPath");
            if (AttachmentBackupPathElement == null) {
                throw new Exception("AttachmentBackupPath not set");
            }
            AttachmentBackupPath = AttachmentBackupPathElement.Value;
            if (Directory.Exists(AttachmentBackupPath) == false) {
                Directory.CreateDirectory(AttachmentBackupPath);
            }

            XElement EnableAutoBccElement = xmlDocument.Root.Element("EnableAutoBcc");
            if (EnableAutoBccElement != null) {
                if (EnableAutoBccElement.Value == "true") {
                    EnableAutoBcc = true;

                    XElement AutoBccEmailAddressElement = xmlDocument.Root.Element("AutoBccEmailAddress");
                    if (AutoBccEmailAddressElement == null) {
                        throw new Exception("AutoBccEmailAddress not set");
                    }
                    AutoBccEmailAddress = AutoBccEmailAddressElement.Value;
                    if (AutoBccEmailAddress == "") {
                        throw new Exception("AutoBccEmailAddress not set");
                    }
                }
            }

            if (xmlDocument.Root.Element("Zoom") == null ||
                Int32.TryParse(xmlDocument.Root.Element("Zoom").Value, out Zoom) == false) {
                Zoom = 150;
            }

            IEnumerable<XElement> CategoryElements = xmlDocument.Root.Elements("Category");
            foreach (XElement CategoryElement in CategoryElements) {
                XAttribute label = CategoryElement.Attribute("label");
                XAttribute color = CategoryElement.Attribute("color");
                CategoryUtil.Data data = new CategoryUtil.Data {
                    label = label.Value,
                    color = CategoryUtil.GetColor(color.Value)
                };
                Categories.Add(data);
            }

            IEnumerable<XElement> FixedReplyElements = xmlDocument.Root.Elements("FixedReply");
            foreach (XElement FixedReplyElement in FixedReplyElements) {
                XAttribute label = FixedReplyElement.Attribute("label");
                XAttribute text = FixedReplyElement.Attribute("text");
                FixedReplyUtil.Data data = new FixedReplyUtil.Data {
                    label = label.Value,
                    text = text.Value
                };
                FixedReplies.Add(data);
            }

            XElement DefaultChineseElement = xmlDocument.Root.Element("DefaultChinese");
            if (DefaultChineseElement == null) {
                throw new Exception("DefaultChinese not set");
            }
            string DefaultChineseStr = DefaultChineseElement.Value;
            switch (DefaultChineseStr) {
                case "Simplified": DefaultChinese = ConvertChineseUtil.ChineseType.SimplifiedChinese; break;
                case "Traditional": DefaultChinese = ConvertChineseUtil.ChineseType.TraditionalChinese; break;
                default: throw new Exception("DefaultChinese not set to Simplified or Traditional");
            }

            IEnumerable<XElement> SimplifiedChineseDeptsElements = xmlDocument.Root.Elements("SimplifiedChineseDept");
            foreach (XElement SimplifiedChineseDeptsElement in SimplifiedChineseDeptsElements) {
                XAttribute startsWith = SimplifiedChineseDeptsElement.Attribute("startsWith");
                SimplifiedChineseDepts.Add(startsWith.Value);
            }

            IEnumerable<XElement> TraditionalChineseDeptsElements = xmlDocument.Root.Elements("TraditionalChineseDept");
            foreach (XElement TraditionalChineseDeptsElement in TraditionalChineseDeptsElements) {
                XAttribute startsWith = TraditionalChineseDeptsElement.Attribute("startsWith");
                TraditionalChineseDepts.Add(startsWith.Value);
            }
        }
    }
}
