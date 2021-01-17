using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Drawing.Imaging;
using Autodesk.AutoCAD.ApplicationServices;
using ACAD = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Customization;
using Windo = Autodesk.Windows;


namespace MyAuto
{
    public class Ribbon : IExtensionApplication
    {
        // All Cui files (main/partial/enterprise) have to be loaded into an object of class 
        // CustomizationSection
        // cs - main AutoCAD CUI file
        CustomizationSection cs;
        Editor ed = ACAD.DocumentManager.MdiActiveDocument.Editor;

        //This bool is used to determine when to save the cui
        //If running the callForAllChanges(), only want to call saveCui at the end
        bool bSaveCui = true;

        String TabName = "Talon_Codes";
        String PanelName = "Initialization";

        public void Initialize()
        {
            //Check Wether tab is present
            cs = new CustomizationSection((string)ACAD.GetSystemVariable("MENUNAME"));
            RibbonRoot ribbonRoot = cs.MenuGroup.RibbonRoot;
            RibbonTabSourceCollection tabs = ribbonRoot.RibbonTabSources;
            RibbonTabSource tab = ribbonRoot.FindTab(TabName + "_TabSourceID");
            if (tab == null) { DeleteTab((string)ACAD.GetSystemVariable("WSCURRENT"), TabName, PanelName); CreateRibbonTabAndPanel_Method(); }
            CreateStartupCommands();
        }

        public void Terminate() { }

        [CommandMethod("CreateRibbonTabAndPanel")] public void CreateRibbonTabAndPanel_Method()
        {       
            try
            {
                string curWorkspace = (string)ACAD.GetSystemVariable("WSCURRENT");
                CreateRibbonTabAndPanel(cs, curWorkspace, TabName, PanelName);
                RibbonPanelSource panelSrc = GetRibbonPanel(cs, PanelName);
                if (panelSrc == null) return;

                MacroGroup macGroup = cs.MenuGroup.MacroGroups[0];

                panelSrc.Items.Clear();
                RibbonRow row = new RibbonRow();
                panelSrc.Items.Add(row);

                RibbonCommandButton button1 = new RibbonCommandButton(row);
                button1.Text = "Load library";
                MenuMacro menuMac1 = macGroup.CreateMenuMacro("Load Dlls", "^C^CNETLOAD ", "Load Dlls in AutoCAD", "Refer C:/Autodesk/ For dll files",
                                    MacroType.Any, "ANAW_16x16.bmp", "ANAW_32x32.bmp", "Button1_Label_Id");
                button1.MacroID = menuMac1.ElementID;
                button1.ButtonStyle = RibbonButtonStyle.LargeWithText;
                button1.KeyTip = "Load Dlls in AutoCAD!!!";
                button1.TooltipTitle = "Load Dlls!!!";
                row.Items.Add(button1);

                //RibbonSeparator separator1 = new RibbonSeparator(row);
                //separator1.SeparatorStyle = RibbonSeparatorStyle.Line;
                //row.Items.Add(separator1);

                if (bSaveCui) { SaveCUI(curWorkspace); }
                ed.WriteMessage(Environment.NewLine + "Menu Created");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(Environment.NewLine + ex.Message);
            }
        }

        [CommandMethod("savecui")] public void SaveCUI(string curWorkspace)
        {
            // Save all Changes made to the CUI file in this session. 
            // If changes were made to the Main CUI file - save it
            // If changes were made to teh Partial CUI files need to save them too
            if (cs.IsModified)
                cs.Save();

            // Here we unload and reload the main CUI file so the changes to the CUI file could take effect immediately.
            ACAD.SetSystemVariable("FILEDIA", 0);

            string cuiMenuGroup = cs.MenuGroup.Name;
            ACAD.DocumentManager.MdiActiveDocument.SendStringToExecute("cuiunload " + cuiMenuGroup + " ", false, false, true);

            string flName = cs.CUIFileName;
            string flNameWithQuotes = "\"" + flName + "\"";
            ACAD.DocumentManager.MdiActiveDocument.SendStringToExecute("cuiload " + flNameWithQuotes + "\n", false, false, true);

            String currentWkspace = curWorkspace;
            String cmdString = "WSCURRENT " + currentWkspace + "\n";
            ACAD.DocumentManager.MdiActiveDocument.SendStringToExecute(cmdString, false, false, true);

            cmdString = "RIBBON ";
            ACAD.DocumentManager.MdiActiveDocument.SendStringToExecute(cmdString, false, false, true);

            cmdString = "MENUBAR " + "1" + "\n"; ;
            ACAD.DocumentManager.MdiActiveDocument.SendStringToExecute(cmdString, false, false, true);

            ACAD.DocumentManager.MdiActiveDocument.SendStringToExecute("filedia 1 ", false, false, true);

        }

        [CommandMethod("DELTAB")] public void DeleteTab(string strWrkSpaceName, string strRibTabSrcName, string strPanelTabSrcName)
        {
            WSRibbonRoot wrkSpaceRibbonRoot = null;
            WorkspaceRibbonTabCollection wsRibbonTabCollection = null;

            try
            {
                WorkspaceCollection WsCollect = cs.Workspaces;

                if ((WsCollect.Count <= 0))
                {
                    ed.WriteMessage("Failed to Get the WorkspaceCollection\n");
                    return;
                }

                if (-1 != cs.Workspaces.IndexOfWorkspaceName(strWrkSpaceName))
                {
                    int curWsIndex = cs.Workspaces.IndexOfWorkspaceName(strWrkSpaceName);
                    wrkSpaceRibbonRoot = cs.Workspaces[curWsIndex].WorkspaceRibbonRoot;
                }
                else
                {
                    ed.WriteMessage("Workspace '" + strWrkSpaceName + "' not found\n");
                    return;
                }

                //Get the Tabcollection in the Workspace
                wsRibbonTabCollection = wrkSpaceRibbonRoot.WorkspaceTabs;
                if (wsRibbonTabCollection.Count <= 0)
                {
                    ed.WriteMessage("Workspace does not contain any Tabs\n");
                    return;
                }

                RibbonRoot ribbonRoot = cs.MenuGroup.RibbonRoot;
                RibbonTabSourceCollection tabs = ribbonRoot.RibbonTabSources;

                // Change this to the name of the  // tab to be removed
                RibbonTabSource ribTabSrc = ribbonRoot.FindTab(strRibTabSrcName + "_TabSourceID");
                string strRibTabSrcTabId = ribTabSrc.ElementID.ToString();
                //Remove From workspace based on eleament ID
                foreach (WSRibbonTabSourceReference wsTabRef in wsRibbonTabCollection)
                {
                    if (wsTabRef.TabId.ToString() == strRibTabSrcTabId)
                    {
                        wsRibbonTabCollection.Remove(wsTabRef);
                        break;
                    }
                }

                RibbonTabSource tab = ribbonRoot.FindTab(strRibTabSrcName + "_TabSourceID");
                if (tab != null) tabs.Remove(tab);
                RibbonPanelSourceCollection panels = ribbonRoot.RibbonPanelSources;
                RibbonPanelSource panel = ribbonRoot.FindPanel(strPanelTabSrcName + "_PanelSourceID");
                if (panel != null) panels.Remove(panel);
                //Workspace Name as entry
                //SaveCUI(strWrkSpaceName);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }
        }

        public static void CreateRibbonTabAndPanel(CustomizationSection cs, string toWorkspace, string tabName, string panelName)
        {
            RibbonRoot root = cs.MenuGroup.RibbonRoot;
            RibbonPanelSourceCollection panels = root.RibbonPanelSources;

            //Create the ribbon panel source and add it to the ribbon panel source collection
            RibbonPanelSource panelSrc = new RibbonPanelSource(root);
            panelSrc.Text = panelSrc.Name = panelName;
            panelSrc.ElementID = panelSrc.Id = panelName + "_PanelSourceID";
            panels.Add(panelSrc);

            //Create the ribbon tab source and add it to the ribbon tab source collection
            RibbonTabSource tabSrc = new RibbonTabSource(root);
            tabSrc.Name = tabSrc.Text = tabName;
            tabSrc.ElementID = tabSrc.Id = tabName + "_TabSourceID";
            root.RibbonTabSources.Add(tabSrc);

            //Create the ribbon panel source reference and add it to the ribbon panel source reference collection
            RibbonPanelSourceReference ribPanelSourceRef = new RibbonPanelSourceReference(tabSrc);
            ribPanelSourceRef.PanelId = panelSrc.ElementID;
            tabSrc.Items.Add(ribPanelSourceRef);

            //Get the ribbon root of the workspace
            int curWsIndex = cs.Workspaces.IndexOfWorkspaceName(toWorkspace);
            WSRibbonRoot wsRibbonRoot = cs.Workspaces[curWsIndex].WorkspaceRibbonRoot;

            //Create the workspace ribbon tab source reference
            WSRibbonTabSourceReference tabSrcRef = WSRibbonTabSourceReference.Create(tabSrc);

            //Set the owner of the ribbon tab source reference and add it to the workspace ribbon tab collection
            tabSrcRef.SetParent(wsRibbonRoot);
            wsRibbonRoot.WorkspaceTabs.Add(tabSrcRef);
        }

        public static RibbonPanelSource GetRibbonPanel(CustomizationSection cs, string panelName)
        {
            RibbonRoot root = cs.MenuGroup.RibbonRoot;
            RibbonPanelSourceCollection panels = root.RibbonPanelSources;

            foreach (RibbonPanelSource panelsrc in panels)
            {
                if (panelsrc.Name == panelName)
                {
                    return panelsrc;
                }
            }
            return null;
        }

        /* using Autodesk.Windows; dll  [CommandMethod("testmyRibbon", CommandFlags.Transparent)] */
        [CommandMethod("STARTUP_BUTTONS")] public void CreateStartupCommands()
        {
            Windo.RibbonControl ribbon = Windo.ComponentManager.Ribbon;
            bool check = false;
            if (ribbon != null)
            {   //Find Ribbon Tab to add panel and button
                Windo.RibbonTab rtab = ribbon.FindTab("ACAD."+TabName + "_TabSourceID");
                if (rtab != null)
                {
                    Windo.RibbonPanelSource rps = new Windo.RibbonPanelSource();
                    rps.Title = "Block Commands";
                    Windo.RibbonPanel Rpanel = AddPanel(rps);
                    AddButton(rps, "Block-Import", Properties.Resources.Import_Excel, "Insert_Block_usingexcel");
                    AddButton(rps, "Block-Export Excel", Properties.Resources.Export_Excel, "ExportBlockDetails_Excel");
                    AddButton(rps, "Block-Export Txt", Properties.Resources.Export_Text, "ExportBlockDetails_Txt");
                    AddButton(rps, "Plot PDF", Properties.Resources.Export_PDF, "PlotToPdf");

                    //Add to Ribbon Tab
                    rtab.Panels.Add(Rpanel);
                }
                else { ed.WriteMessage(Environment.NewLine + "Ribbon Tab not found!!"); }
            }
            else { ed.WriteMessage(Environment.NewLine + "Ribbon Control not found!!"); }
        }

        static Windo.RibbonPanel AddPanel(Windo.RibbonPanelSource ribbonPanelSource)
        {
            Windo.RibbonPanel rp = new Windo.RibbonPanel();
            rp.Source = ribbonPanelSource;
            return rp;
        }

        static void AddButton(Windo.RibbonPanelSource rps, String rbNAME, Bitmap image, String Command)
        {
            Windo.RibbonButton rb = new Windo.RibbonButton();
            
            rb.Name = rbNAME;
            rb.ShowText = true;
            rb.Text = rbNAME;
            rb.ShowImage = true;
            rb.Image = Images.getBitmap(image);
            rb.LargeImage = Images.getBitmap(image);
            rb.Size = Windo.RibbonItemSize.Large;
            rb.CommandParameter = Command;
            rb.CommandHandler = new AusSignCommandHandler();
            rb.Orientation = System.Windows.Controls.Orientation.Vertical;
            rb.Width = 500;

            //Add the Button to the panel
            rps.Items.Add(rb);
            //assign the Command Item to the DialgLauncher which auto - enables
            // the little button at the lower right of a Panel
            //rps.DialogLauncher = rci;
        }

    }

    public class AusSignCommandHandler : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            Document doc = ACAD.DocumentManager.MdiActiveDocument;
            String Sent = (parameter as Windo.RibbonButton).CommandParameter.ToString() + " ";
            //if ((parameter as Windo.RibbonButton).CommandParameter.Equals("Insert_Block_usingexcel"))
            doc.SendStringToExecute(Sent, false, false, true);
        }
    }

    public class Images
    {
        public static BitmapImage getBitmap(Bitmap image)
        {
            MemoryStream stream = new MemoryStream();
            image.Save(stream, ImageFormat.Png);
            BitmapImage bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.StreamSource = stream;
            bmp.EndInit();

            return bmp;
        }
    }
}
