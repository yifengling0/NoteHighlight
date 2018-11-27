using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.OneNote;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using System.Drawing.Imaging;
using System.IO;
using NoteHighLightForm;
using System.Xml.Linq;
using System.Diagnostics;
using System.Reflection;
using System.Drawing;
using Helper;

namespace NoteHighLightAddins
{
    [GuidAttribute("4C6B0362-F139-417F-9661-3663C268B9E9"), ProgId("NoteHighLightAddins.NoteAddins")]
    public class NoteAddins : IDTExtensibility2, IRibbonExtensibility
    {
        private XNamespace ns;

        private Microsoft.Office.Interop.OneNote.Application onApp = new Microsoft.Office.Interop.OneNote.Application();

        #region IDTExtensibility2 成員

        public void OnAddInsUpdate(ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            if (onApp != null)
                onApp = null;
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            onApp = (Microsoft.Office.Interop.OneNote.Application)Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            onApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnStartupComplete(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        #endregion

        #region IRibbonExtensibility 成員

        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.Ribbon;
        }

        #endregion

        /// <summary>
        /// 彈出 HighLightCode 的表單
        /// Show HighLightCode Form
        /// </summary>
        public void ShowCodeForm(IRibbonControl control)
        {
            string outFileName = Guid.NewGuid().ToString();

            try
            {
                ProcessHelper processHelper = new ProcessHelper("NoteHighLightForm.exe",new string[]{control.Tag, outFileName});
                processHelper.IsWaitForInputIdle = true;
                processHelper.ProcessStart();
            }
            catch (Exception ex)
            {
                MessageBox.Show("執行NoteHighLightForm.exe發生錯誤，錯誤訊息為：" + ex.Message);
                return;
            }

            string fileName = Path.Combine(Path.GetTempPath(), outFileName + ".html");

            if (File.Exists(fileName))
                InsertHighLightCodeToCurrentSide(fileName);
        }

        /// <summary>
        /// 插入 HighLight Code 至滑鼠游標的位置
        /// Insert HighLight Code To Mouse Position  
        /// </summary>
        private void InsertHighLightCodeToCurrentSide(string fileName)
        {
            string htmlContent = File.ReadAllText(fileName, Encoding.UTF8);

            string notebookXml;
            onApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            ns = doc.Root.Name.Namespace;

            var pageNode = doc.Descendants(ns + "Page")
                              .Where(n => n.Attribute("isCurrentlyViewed") != null && n.Attribute("isCurrentlyViewed").Value == "true")
                              .FirstOrDefault();

            if (pageNode != null)
            {
                var existingPageId = pageNode.Attribute("ID").Value;

                string[] position = GetMousePointPosition(existingPageId);

                var page = InsertHighLightCode(htmlContent, position);
                page.Root.SetAttributeValue("ID", existingPageId);

                onApp.UpdatePageContent(page.ToString(), DateTime.MinValue);
            }
        }

        /// <summary>
        /// 取得滑鼠所在的點
        /// Get Mouse Point
        /// </summary>
        private string[] GetMousePointPosition(string pageID)
        {
            string pageXml;
            onApp.GetPageContent(pageID, out pageXml, PageInfo.piSelection);

            var node = XDocument.Parse(pageXml).Descendants(ns + "Outline")
                                               .Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "partial")
                                               .FirstOrDefault();
            if (node != null)
            {
                var attrPos = node.Descendants(ns + "Position").FirstOrDefault();
                if (attrPos != null)
                {
                    var x = attrPos.Attribute("x").Value;
                    var y = attrPos.Attribute("y").Value;
                    return new string[] { x, y };
                }
            }
            return null;
        }

        /// <summary>
        /// 產生 XML 插入至 OneNote
        /// Generate XML Insert To OneNote
        /// </summary>
        public XDocument InsertHighLightCode(string htmlContent, string[] position)
        {
            XElement children = new XElement(ns + "OEChildren");

            var arrayLine = htmlContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            foreach (var item in arrayLine)
            {
                children.Add(new XElement(ns + "OE",
                                new XElement(ns + "T",
                                    new XCData(item))));
            }

            XElement outline = new XElement(ns + "Outline");

            if (position != null && position.Length == 2)
            {
                XElement pos = new XElement(ns + "Position");
                pos.Add(new XAttribute("x", position[0]));
                pos.Add(new XAttribute("y", position[1]));
                outline.Add(pos);
            }
            outline.Add(children);

            XElement page = new XElement(ns + "Page");
            page.Add(outline);

            XDocument doc = new XDocument();
            doc.Add(page);

            return doc;
        }

        /// <summary>
        /// 取得資源檔的圖片
        /// Get Resources Picture
        /// </summary>
        public IStream GetImage(string imageName)
        {
            MemoryStream mem = new MemoryStream();

            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            var b = typeof(Properties.Resources).GetProperty(imageName.Substring(0, imageName.IndexOf('.')), flags).GetValue(null, null) as Bitmap;
            b.Save(mem, ImageFormat.Png);

            return new CCOMStreamWrapper(mem);
        }
    }
}
