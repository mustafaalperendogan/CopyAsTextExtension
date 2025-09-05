using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace CopyAsTextExtension
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
    [Guid(CopyAsTextExtensionPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class CopyAsTextExtensionPackage : AsyncPackage
    {
        /// <summary>
        /// FileToTextCopierPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "b9c8d5e2-3a4f-4b8c-9d1e-2f5a8b7c9e3f";

        #region Package Members

        /// <summary>
        /// Async paket başlatma
        /// </summary>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <param name="progress">Progress reporter</param>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // Paket yüklendiğinde, UI thread'e geçip komutları başlat
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await CopyFileAsTextCommand.InitializeAsync(this);
        }

        #endregion
    }
}
