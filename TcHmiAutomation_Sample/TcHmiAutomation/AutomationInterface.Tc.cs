using System;
using System.Collections;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using EnvDTE;
using EnvDTE80;
using TCatSysManagerLib;
using TcHmiAutomation;

namespace TcHmiAutomationInterface
{
    public partial class AutomationInterface
    {
        #region TcProject Handling
        public void CreateNewTcProject()
        {
            /* ------------------------------------------------------
            * Create new TwinCAT project, based on TwinCAT Project file (delivered with TwinCAT XAE)
            * ------------------------------------------------------ */
            envDteProject = solution.AddFromTemplate(tcTemplatePath, baseFolder + solutionFolder + @"\" + projectName + @"\", projectName);

            /* ------------------------------------------------------
             * Cast TwinCAT project object to ITcSysManager interface --> Automation Interface starts here
             * ------------------------------------------------------ */
            if (envDteProject == null)
            {
                MessageBox.Show("Create new TC project failed!");
            }

            sysManager = envDteProject.Object as ITcSysManager;
        }

        public void OpenExistingTcProject()
        {
            /* ------------------------------------------------------
             * Open existing solution
             * ------------------------------------------------------ */
            Type t = System.Type.GetTypeFromProgID("VisualStudio.DTE.15.0");
            dte = (DTE)Activator.CreateInstance(t);
            dte.SuppressUI = false;
            dte.MainWindow.Visible = true;

            solution = dte.Solution;
            solution.Open(baseFolder + solutionFolder + solutionName);

            SetProjectReference(".tsproj");
        }

        #endregion      

    }
}
