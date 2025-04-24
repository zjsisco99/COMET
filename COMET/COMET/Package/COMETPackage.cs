using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace COMET
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(COMETPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(ToolWindow1))]
    public sealed class COMETPackage : AsyncPackage
    {
        /// <summary>
        /// COMETPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "075c6195-3d8f-4e9f-91b6-5352e43540a7";
        public static DTE2 DTEInstance { get; private set; }
        public static IVsFontAndColorStorage FontAndColorStorage { get; private set; }
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
       /*METHOD  Created By: AP Date: 03/18/2025 Events Handled:  Event Raised:  Exception Thrown:  Exception Caught:  Purpose:  Revise History: */
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            DTEInstance = await GetServiceAsync(typeof(DTE)) as DTE2;
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            FontAndColorStorage = await GetServiceAsync(typeof(SVsFontAndColorStorage)) as IVsFontAndColorStorage;
            await ToolWindow1Command.InitializeAsync(this);
            // Create a CommentMethodCommand Object
            // That is the file that holds the COMET commenting logic
            await CommentMethodCommand.InitializeAsync(this);

        }

        #endregion
        public static ToolWindow1Control ToolWindowControlInstance { get; set; }



    }




}
