using System;
using System.Collections;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using EnvDTE;
using EnvDTE80;
using TCatSysManagerLib;

namespace TcHmiAutomationInterface
{
    public partial class AutomationInterface
    {
        // member variables
        DTE dte                     = null;
        Solution solution           = null;
        dynamic envDteProject       = null;
        ITcSysManager sysManager    = null;

        string Vs2017               = "VisualStudio.DTE.15.0";
        string TcXaeShell           = "TcXaeShell.DTE.15.0";
        string baseFolder           = @"C:\temp\1.12\Ai\Test\";
        string solutionFolder       = "AiSuite";
        string solutionName         = @"\AiSuite.sln";
        string tcTemplatePath       = @"C:\TwinCAT\3.1\Components\Base\PrjTemplate\TwinCAT Project.tsproj";
        string templateDir          = @"C:\temp\Training\Templates\";
        string projectName          = "AiSuite";
        public const string vsProjectKindMisc = "{66A2671D-8FB5-11D2-AA7E-00C04F688DDE}";

        public AutomationInterface()
        {
            // register COM message filter
            MessageFilter.Register();
        }

        ~AutomationInterface()
        {
            // revoke COM message filter
            MessageFilter.Revoke();
        }

        #region VS Handling        

        public bool CreateVsInstance(string progId)
        {                 
            Type t = System.Type.GetTypeFromProgID(progId);
            dte = (DTE)Activator.CreateInstance(t);

            if (dte == null)
            {
                MessageBox.Show("Creation of dte failed");
                return false;
            }

            dte.SuppressUI = true;
            dte.MainWindow.Visible = true;

            return true;
        }

        public bool DeleteAndCreateDirectory(string baseFolder, string solutionFolder)
        {
            try
            {
                DirectoryHelper.DeleteDirectory(baseFolder);
                Directory.CreateDirectory(baseFolder + solutionFolder);

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Delete or create directory failed " + e.Message);
                return false;
            }            
        }

        public bool CreateSolution(string baseFolder, string solutionFolder, string solutionName)
        {
            if (dte == null)
                return false;

            solution = dte.Solution;

            try
            {
                solution.Create(baseFolder, solutionFolder);
                solution.SaveAs(baseFolder + solutionFolder + solutionName);

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Create solution failed " + e.Message);

                return false;
            }            
        }
                
        public void AttachToVsInstance()
        {
            /* ------------------------------------------------------
            * Get reference to an existing vs instance with a helper class
            * ------------------------------------------------------ */
            DTE _dte = null;
            try
            {
                Hashtable dteInstances = AttachVsInstanceHelper.GetIDEInstances(false, "VisualStudio.DTE.15.0");
                IDictionaryEnumerator hashtableEnumerator = dteInstances.GetEnumerator();

                /* ------------------------------------------------------
                * Search all vs instances for a specific solution
                * ------------------------------------------------------ */
                while (hashtableEnumerator.MoveNext())
                {
                    DTE dteTemp = (DTE)hashtableEnumerator.Value;
                    if (dteTemp.Solution.FullName == baseFolder + solutionFolder + solutionName)
                    {
                        Console.WriteLine("Found solution in list of all open DTE objects. " + dteTemp.Name);
                        _dte = dteTemp;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            // solution was found in an existing vs instance
            if (_dte != null)
            {
                dte = _dte;
                solution = dte.Solution;               
            }
            else
            {
                MessageBox.Show("Attach to VS instance failed");
            }
        }

        #endregion
        

        #region Solution Handling

        private void SetProjectReference(string fileExt)
        {
            /* ------------------------------------------------------
            * Set the project reference for an existing project in a solution and cast to ITcSysManager interface
            * ------------------------------------------------------ */
            if (solution != null)
            {
                foreach (var pp in solution.Projects)
                {
                    Project prj = pp as Project;

                    if (prj == null)
                        continue;

                    // get first hit
                    // or search by name etc.

                    string ext = Path.GetExtension(prj.FileName);
                    if (String.Compare(ext, fileExt, false) == 0)
                    {
                        envDteProject = prj;

                        sysManager = envDteProject.Object as ITcSysManager;

                        if (sysManager != null)
                        {
                            Console.WriteLine("Project found");
                        }                        
                    }                       
                }
            }
        }

        public void SaveProjectAndSolution()
        {
            /* ------------------------------------------------------
            * Save project and solution if the references are valid
            * ------------------------------------------------------ */

            if (solution != null)
            {
                if (solution.Projects != null)
                {
                    for (var i = 1; i <= solution.Projects.Count; i++)
                    {
                        // skip virtual projects and folders
                        if (solution.Projects.Item(i).Kind == vsProjectKindMisc)
                            continue;

                        solution.Projects.Item(i).Save();
                    }
                }    
                
                solution.SaveAs(baseFolder + solutionFolder + solutionName);
            }           
        }
       
        public void CloseSolution()
        {            
            if (solution != null)
            {
                solution.Close(true);
            }                
        }

        public void ExitVisualStudio()
        {
            if (dte != null)
            {
                // quit dte
                dte.Quit();
                dte = null;
            }
        }

        public void BuildSolution()
        {
            if (dte.Solution == null)
                return;

            dte.Solution.SolutionBuild.Build(true);
        }
           
        public bool CompilerErrors()
        {
            if (dte.Solution.SolutionBuild.LastBuildInfo != 0)
                return true;

            DTE2 dte2 = dte as DTE2;
            ErrorItems errors = dte2.ToolWindows.ErrorList.ErrorItems;

            if (errors.Count != 0)
                return true;           

            return false;
        }

        #endregion              
           

    }


}
