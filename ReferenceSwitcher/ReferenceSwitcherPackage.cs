using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using ReferenceSwitcher;

namespace Microsoft.VSPackage1
{
    static class GuidList
    {
        public const string guidVSPackage1PkgString = "39bf4245-64d8-4180-9c00-b8c809c29c4a";
    };

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string)]
    [Guid(GuidList.guidVSPackage1PkgString)]
    public sealed class VSPackage1Package : Package
    {
        private SolutionEvents solutionEvents;
        private ReferenceHelper referenceHelper;

        public VSPackage1Package()
        {
        }

        protected override void Initialize()
        {
            base.Initialize();
            referenceHelper = new ReferenceHelper(this.AskUserToProceed);

            var envDTE = GetService(typeof (EnvDTE.DTE)) as EnvDTE80.DTE2;
            solutionEvents = envDTE.Events.SolutionEvents;
            solutionEvents.ProjectAdded += referenceHelper.SolutionEvents_ProjectAdded;
            solutionEvents.ProjectRemoved += referenceHelper.SolutionEvents_ProjectRemoved;
        }

        private bool AskUserToProceed(string title, string text)
        {
            var uiShell = (IVsUIShell) GetService(typeof (SVsUIShell));
            // Show a Message Box to prove we were here
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       title,
                       text,
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_YESNO,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,        // false
                       out result));

            return result == 6;
        }
    }
}
