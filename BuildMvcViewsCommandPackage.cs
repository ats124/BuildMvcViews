using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace BuildMvcViews
{
    /// <summary>
    ///     This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         The minimum requirement for a class to be considered a valid package for Visual Studio
    ///         is to implement the IVsPackage interface and register itself with the shell.
    ///         This package uses the helper classes defined inside the Managed Package Framework (MPF)
    ///         to do it: it derives from the Package class that provides the implementation of the
    ///         IVsPackage interface and uses the registration attributes defined in the framework to
    ///         register itself and its components with the shell. These attributes tell the pkgdef creation
    ///         utility what data to put into .pkgdef file.
    ///     </para>
    ///     <para>
    ///         To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...
    ///         &gt; in .vsixmanifest file.
    ///     </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class BuildMvcViewsCommandPackage : AsyncPackage
    {
        /// <summary>
        ///     BuildMvcViewsCommandPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "052835ec-a5af-4d90-85f2-a840b5468a6e";

        private const string MsBuildXmlNamespace = "http://schemas.microsoft.com/developer/msbuild/2003";

        public void BuildCurrentProjectWithBuildMvcViews()
        {
            var dte = (DTE2)GetService(typeof(DTE));
            var project = dte.SelectedItems.OfType<SelectedItem>().FirstOrDefault()?.Project;
            if (project == null) return;

            if (!PrepareUserProjectXml(project, out var oldValue)) return;

            var solutionBuild = dte.Solution.SolutionBuild;
            solutionBuild.BuildProject(solutionBuild.ActiveConfiguration.Name, project.UniqueName, true);

            RestoreUserProjectXml(project, oldValue);
        }

        private bool PrepareUserProjectXml(Project project, out string oldValue)
        {
            oldValue = null;

            var userProjectFile = project.FullName + ".user";
            if (!File.Exists(userProjectFile))
            {
                MessageBox.Show(".userファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            var userProjectXml = new XmlDocument();
            try
            {
                userProjectXml.Load(userProjectFile);
            }
            catch (Exception)
            {
                MessageBox.Show($"{Path.GetFileName(userProjectFile)}ファイルを読み込ません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            var xmlns = new XmlNamespaceManager(userProjectXml.NameTable);
            xmlns.AddNamespace("ns", MsBuildXmlNamespace);

            var projectNode = userProjectXml.SelectSingleNode("/ns:Project", xmlns);
            if (projectNode == null)
            {
                MessageBox.Show($"Projectノードが見つかりません。{Path.GetFileName(userProjectFile)}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            var propertyGroupNode = projectNode.SelectSingleNode("./ns:PropertyGroup", xmlns);
            if (propertyGroupNode == null)
            {
                propertyGroupNode = userProjectXml.CreateElement("PropertyGroup", MsBuildXmlNamespace);
                projectNode.AppendChild(propertyGroupNode);
            }

            var mvcBuildViewsNode = propertyGroupNode.SelectSingleNode("./ns:MvcBuildViews", xmlns);
            if (mvcBuildViewsNode == null)
            {
                mvcBuildViewsNode = userProjectXml.CreateElement("MvcBuildViews", MsBuildXmlNamespace);
                propertyGroupNode.AppendChild(mvcBuildViewsNode);
            }

            oldValue = mvcBuildViewsNode.InnerText;
            mvcBuildViewsNode.InnerText = "true";
            userProjectXml.Save(userProjectFile);

            return true;
        }

        private void RestoreUserProjectXml(Project project, string oldValue)
        {
            var userProjectFile = project.FullName + ".user";
            if (!File.Exists(userProjectFile)) return;
            var userProjectXml = new XmlDocument();
            try
            {
                userProjectXml.Load(userProjectFile);
            }
            catch (Exception)
            {
                return;
            }

            var xmlns = new XmlNamespaceManager(userProjectXml.NameTable);
            xmlns.AddNamespace("ns", MsBuildXmlNamespace);

            var mvcBuildViewsNode = userProjectXml.SelectSingleNode("/ns:Project/ns:PropertyGroup/ns:MvcBuildViews", xmlns);
            if (mvcBuildViewsNode == null) return;

            mvcBuildViewsNode.InnerText = oldValue;

            userProjectXml.Save(userProjectFile);
        }

        #region Package Members

        /// <summary>
        ///     Initialization of the package; this method is called right after the package is sited, so this is the place
        ///     where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">
        ///     A cancellation token to monitor for initialization cancellation, which can occur when
        ///     VS is shutting down.
        /// </param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>
        ///     A task representing the async work of package initialization, or an already completed task if there is none.
        ///     Do not return null from this method.
        /// </returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await BuildMvcViewsCommand.InitializeAsync(this);
        }

        #endregion
    }
}