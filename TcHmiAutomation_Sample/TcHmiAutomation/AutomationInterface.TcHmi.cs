using System.Windows;
using TCatSysManagerLib;
using TcHmiAutomation;
using TcHmiAutomation.Publish;


namespace TcHmiAutomationInterface
{
    public partial class AutomationInterface
    {
        ITcHmiAutomation tcHmiAi;
        ITcHmiProject tcHmiPrj;

        string prjName = "MyFirstPrj";

        string tcHmiTemplateName = "TwinCAT HMI Project";
        string tcHmiFrameworkTemplate = "TwinCAT HMI Framework Project";
        string tcHmiServerExtension = "ServerExtensionCSharp";
        string tcHmiServerExtensionReflection = "ServerExtensionCSharpReflection";      

        public void TriggerTest()
        {
            // Create environment
            CreateVsInstance(Vs2017);
            DeleteAndCreateDirectory(baseFolder, solutionFolder);
            CreateSolution(baseFolder, solutionFolder, solutionName);

            // get reference to automation interface entry point
            tcHmiAi = dte.GetObject("Beckhoff.TcHmi.1.12");

            if (tcHmiAi == null)
                return;

            // create a new HMI project
            tcHmiPrj = tcHmiAi.CreateHmiProject(tcHmiTemplateName, baseFolder + solutionFolder + @"\" + prjName + @"\", prjName);

            if (tcHmiPrj == null)
                return;

            // Part 1: General handling

            // get reference to Desktop.view
            ITcHmiItem desktop = tcHmiPrj.LookupChild("Desktop.view");

            // get infos using ITcHmiItem
            string desktopName = desktop.Name;            

            // add a new folder
            ITcHmiItem folder1 = tcHmiPrj.AddFolder("Contents");

            // add new content files
            ITcHmiItem content1 = folder1.AddItem("MyContent1", HmiTemplates.Content);
            ITcHmiItem content2 = folder1.AddItem("MyContent2", HmiTemplates.Content);

            // get infos using ITcHmiItem
            string pathContent1 = content1.PathName;
            string parentName = content1.Parent.Name;

            // cast to file interface
            ITcHmiFile content1File = content1 as ITcHmiFile;

            // create a new button on content 1
            ITcHmiControl ctrlButton = content1File.AddControl("MyContent1", "MyButton", "TcHmi.Controls.Beckhoff.TcHmiButton");

            ctrlButton.ChangeAttributes(new ITcHmiControlAttribute[]
            {
                tcHmiPrj.GetControlAttributeInstance("data-tchmi-left", "100"),
                tcHmiPrj.GetControlAttributeInstance("data-tchmi-top", "200")
            });

            // build project and deploy in browser (F5 build)
            tcHmiPrj.ShowInBrowser(true);

            // Part 2: PLC and control handling

            // create sample plc
            CreateSamplePlc();

            // get reference to server interface
            ITcHmiServer server = tcHmiPrj.GetServerInterface();

            // map a symbol
            server.MapSymbol("PLC1.MAIN.bStart", "PLC1::MAIN::bStart", "ADS");
              
            // refresh symbols
            server.RefreshSymbols();
            // wait until refresh 
            System.Threading.Thread.Sleep(10);

            // stop subscription mode
            // neccessary if you add a lot of controls with links to server variables
            tcHmiPrj.ToggleSubscriptionMode(false);

            // create a new button on content 1
            ITcHmiControl ctrlButton2 = content1File.AddControl("MyContent1", "MyButton2", "TcHmi.Controls.Beckhoff.TcHmiButton");

            ctrlButton2.ChangeAttributes(new ITcHmiControlAttribute[]
            {
                tcHmiPrj.GetControlAttributeInstance("data-tchmi-left", "200"),
                tcHmiPrj.GetControlAttributeInstance("data-tchmi-top", "300"),
                tcHmiPrj.GetControlAttributeInstance("data-tchmi-state-symbol", "%s%PLC1.MAIN.fbManager.fbTest1.bStart%/s%")
            });

            // activate subscription mode again
            tcHmiPrj.ToggleSubscriptionMode(true);

            // Part 3: Historize handling

            // install historize extension via NuGet
            tcHmiPrj.AddNuGetPackage(@"C:\TwinCAT\Functions\TE2000-HMI-Engineering\Infrastructure\Packages", "Beckhoff.TwinCAT.HMI.SqliteHistorize", "12.744.2");

            // get list of all mapped symbols
            ITcHmiMappedSymbol[] mappedSymbols = server.GetMappedSymbols(true);           
            
            ITcHmiMappedSymbol myNewSymbol = null;

            // search in list of mapped symbols
            foreach (ITcHmiMappedSymbol sym in mappedSymbols)
            {
                if (sym.MappedName == "PLC1.MAIN.bStart")
                {
                    myNewSymbol = sym;
                    break;
                }
            }

            if (myNewSymbol != null)
            {
                // create new historize settings
                ITcHmiHistorizeSettings histo = tcHmiPrj.GetHistorizeSettingsInstance();               

                // define settings
                histo.Interval = "PT1S";
                histo.RowLimit = 10000;
                histo.MaxEntries = 10000;

                // settings must be part of a recoding setting
                ITcHmiRecordingSettings recordingsDefault = tcHmiPrj.GetRecordingSettingsInstance();

                // recording settings can be different in different configurations
                recordingsDefault.PublishConfiguration = "default";
                recordingsDefault.Recording = true;

                // remote profile
                ITcHmiRecordingSettings recordingsRemote = tcHmiPrj.GetRecordingSettingsInstance();

                recordingsRemote.PublishConfiguration = "remote";
                recordingsRemote.Recording = true;

                histo.RecordingSettings = new[] { recordingsDefault, recordingsRemote };

                // appli historize settings 
                myNewSymbol.ApplyHistorizeSettings(histo);

                // update symbols (compare to refresh in config window)
                server.RefreshSymbols();
            }

            // Part 4: Internal Symbol handling

            // get a new internal symbol instance
            ITcHmiInternalSymbol myFirstInternalSymbol = tcHmiPrj.GetInternalSymbolInstance();

            // define smybol information
            myFirstInternalSymbol.Name = "MyFirstInternalSymbol";
            myFirstInternalSymbol.Datatype = "tchmi:general#/definitions/Number";
            myFirstInternalSymbol.DefaultValue = 123;

            // add internal symbol
            tcHmiPrj.AddInternalSymbol(myFirstInternalSymbol);

            // remove it by name
            tcHmiPrj.RemoveInternalSymbol("MyFirstInternalSymbol");

            // Part 5: Localization handling

            // add a new localization
            ITcHmiLocalizationItem locEnGB = tcHmiPrj.AddLocalization(@"Localization\en-GB", "en-GB");

            // add a new localization entry with default text for all languages
            tcHmiPrj.AddLocalizationEntry("MyOwnKey", "Default Text");

            // change text for specific language
            locEnGB.AddEdit("MyOwnKey", "This is an english text");

            // Part 6: Penetration of HMI

            Create200Buttons();

            // Save all
            SaveProjectAndSolution();     
        }


        #region TcHmi Project Handling       

        public bool CreateNewHmiProject(string prjName, out ITcHmiProject prj)
        {
            prj = tcHmiAi.CreateHmiProject(tcHmiTemplateName, baseFolder + solutionFolder + @"\" + prjName + @"\", prjName);

            if (prj == null)
            {
                MessageBox.Show("Unable to create new HMI project");
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool OpenHmiProject(string prjName, out ITcHmiProject prj)
        {
            if (dte.Solution == null)
            {
                prj = null;
                return false;
            }

            dte.Solution.Open(baseFolder + solutionFolder + solutionName);

            prj = tcHmiAi.GetHmiProject(prjName);

            if (prj == null)
                return false;


            return true;
        }

        #endregion
                
        #region Test Cases

        public bool Create200Buttons()
        {
            tcHmiPrj.ToggleSubscriptionMode(false);

            tcHmiPrj.AddFolder("TestTim");

            tcHmiPrj.AddFolder(@"TestTim\Test");

            ITcHmiFile myContent = tcHmiPrj.AddContent(@"TestTim\Test\NewContent");

            ITcHmiControl itself = myContent.GetControl("NewContent");

            itself.ChangeAttributes(new ITcHmiControlAttribute[]
            {
                    tcHmiPrj.GetControlAttributeInstance("data-tchmi-width","1920"),
                    tcHmiPrj.GetControlAttributeInstance("data-tchmi-height","1080")
            });

            System.Threading.Thread.Sleep(3000);

            int left = 0;
            int top = 0;

            for (var i = 0; i < 200; i++)
            {
                ITcHmiControl ctrl = myContent.AddControl("NewContent", "MyButton" + i, "TcHmi.Controls.Beckhoff.TcHmiButton");

                ctrl.ChangeAttributes(new ITcHmiControlAttribute[]
                    {
                         tcHmiPrj.GetControlAttributeInstance("data-tchmi-left", left.ToString()),
                         tcHmiPrj.GetControlAttributeInstance("data-tchmi-top", top.ToString())
                    });

                left = left + 100;

                if (left >= 1900)
                {
                    left = 0;
                    top = top + 50;
                }
            }

            tcHmiPrj.ToggleSubscriptionMode(true);

            return true;
        }

        public void CreateSamplePlc()
        {
            CreateNewTcProject();

            ITcSmTreeItem plc = sysManager.LookupTreeItem("TIPC");           
            ITcSmTreeItem newProject = plc.CreateChild("PLC1", 0, "", "Standard PLC Template");
            ITcSmTreeItem main = sysManager.LookupTreeItem("TIPC^PLC1^PLC1 Project^POUs^MAIN");
            ITcPlcDeclaration mainDecl = main as ITcPlcDeclaration;
            ITcPlcImplementation mainImpl = main as ITcPlcImplementation;
          
            string strMainDeclNew = @"PROGRAM MAIN
VAR
    nCount : INT;
    bStart   : BOOL;
END_VAR";

            string strMainImplNew = @"IF bStart THEN
    nCount := nCount + 1;    
END_IF

IF nCount >= 100 THEN
    nCount := 0;
END_IF";

            mainDecl.DeclarationText = strMainDeclNew;
            mainImpl.ImplementationText = strMainImplNew;

            // activate configuration
            sysManager.ActivateConfiguration();

            // start TwinCAT in run mode
            sysManager.StartRestartTwinCAT();

            System.Threading.Thread.Sleep(4000);

        }
    

        #endregion
    }
}
