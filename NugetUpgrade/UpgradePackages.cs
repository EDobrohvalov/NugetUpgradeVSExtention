using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Task = System.Threading.Tasks.Task;

namespace NugetUpgrade
{

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class UpgradePackages
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("1cbe85f2-1989-4d41-a09b-4d42b7068d92");
        private readonly AsyncPackage package;
        private static bool _isProcessing;
        public static DTE2 _dte;


        public static UpgradePackages Instance;


        private UpgradePackages(AsyncPackage package, OleMenuCommandService commandService, DTE2 dte)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
            _dte = dte;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            DTE2 dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            Instance = new UpgradePackages(package, commandService, dte);           
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            UpgradePackagesConfig();
        }


        private void UpgradePackagesConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _isProcessing = true;

            var projects = ProjectHelpers.GetAllProjects().Where(c => File.Exists(c.GetFullPath() + @"\packages.config")).ToArray();

            if (!projects.Any())
            {
                _dte.StatusBar.Text = "Please select a package.config file to nuke from orbit.";
                _isProcessing = false;
                return; 
            }

            //var projectFolder = ProjectHelpers.GetRootFolder(ProjectHelpers.GetActiveProject());
            int count = projects.Count();

            //RWM: Don't mess with these.
            XNamespace defaultNs = "http://schemas.microsoft.com/developer/msbuild/2003";
            var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

            try
            {
                string text = count == 1 ? " file" : " files";
                _dte.StatusBar.Progress(true, $"Fixing {count} config {text}...", AmountCompleted: 1, Total: count + 1);

                //for (int i = 0; i < count; i++)

                Parallel.For(0, count, options, i =>
                {
                    var packageReferences = new XElement(defaultNs + "ItemGroup");
                    var packagesConfigItem = projects.ElementAt(i);
                    var packagesConfigPath = packagesConfigItem.GetFullPath() + @"\packages.config";
                    var projectPath = packagesConfigItem.FileName;

                    //RWM: Start by backing up the files.
                    File.Copy(packagesConfigPath, $"{packagesConfigPath}.bak", true);
                    File.Copy(projectPath, $"{projectPath}.bak", true);

                    // Logger.Instance.LogAsync($"Backup created for {packagesConfigPath}.").Wait();

                    //RWM: Load the files.
                    var project = XDocument.Load(projectPath);
                    var packagesConfig = XDocument.Load(packagesConfigPath);

                    //RWM: Get references to the stuff we're gonna get rid of.
                    var oldReferences = project.Root.Descendants().Where(c => c.Name.LocalName == "Reference");
                    var errors = project.Root.Descendants().Where(c => c.Name.LocalName == "Error");
                    var targets = project.Root.Descendants().Where(c => c.Name.LocalName == "Import");

                    foreach (var row in packagesConfig.Root.Elements().ToList())
                    {
                        //RWM: Create the new PackageReference.
                        packageReferences.Add(new XElement(defaultNs + "PackageReference",
                            new XAttribute("Include", row.Attribute("id").Value),
                            new XAttribute("Version", row.Attribute("version").Value)));

                        //RWM: Remove the old Standard Reference.
                        oldReferences.Where(c => c.Attribute("Include").Value.Split(new Char[] { ',' })[0].ToLower() == row.Attribute("id").Value.ToLower()).ToList()
                            .ForEach(c => c.Remove());
                        //RWM: Remove any remaining Standard References where the PackageId is in the HintPath.
                        oldReferences.Where(c => c.Descendants().Any(d => d.Value.Contains(row.Attribute("id").Value))).ToList()
                            .ForEach(c => c.Remove());
                        //RWM: Remove any Error conditions for missing Package Targets.
                        errors.Where(c => c.Attribute("Condition") != null)
                            .Where(c => c.Attribute("Condition").Value.Contains(row.Attribute("id").Value)).ToList()
                            .ForEach(c => c.Remove());
                        //RWM: Remove any Package Targets.
                        targets.Where(c => c.Attribute("Project").Value.Contains(row.Attribute("id").Value)).ToList()
                            .ForEach(c => c.Remove());
                    }

                    //RWM: Fix up the project file by adding PackageReferences, removing packages.config, and pulling NuGet-added Targets.
                    project.Root.Elements().First(c => c.Name.LocalName == "ItemGroup").AddBeforeSelf(packageReferences);
                    var packageConfigReference = project.Root.Descendants().FirstOrDefault(c => c.Name.LocalName == "None" && c.Attribute("Include").Value == "packages.config");
                    if (packageConfigReference != null)
                    {
                        packageConfigReference.Remove();
                    }

                    var nugetBuildImports = project.Root.Descendants().FirstOrDefault(c => c.Name.LocalName == "Target" && c.Attribute("Name").Value == "EnsureNuGetPackageBuildImports");
                    if (nugetBuildImports != null && nugetBuildImports.Descendants().Count(c => c.Name.LocalName == "Error") == 0)
                    {
                        nugetBuildImports.Remove();
                    }

                    //RWM: Upgrade the ToolsVersion so it can't be opened in VS2015 anymore.
                    project.Root.Attribute("ToolsVersion").Value = "15.0";

                    //RWM: Save the project and delete Packages.config.
                    ProjectHelpers.CheckFileOutOfSourceControl(projectPath);
                    ProjectHelpers.CheckFileOutOfSourceControl(packagesConfigPath);
                    project.Save(projectPath, SaveOptions.None);
                    File.Delete(packagesConfigPath);

                });
            }
            catch (AggregateException agEx)
            {
                _dte.StatusBar.Progress(false);

                _dte.StatusBar.Text = "Operation failed. Please see Output Window for details.";
                _isProcessing = false;

            }
            finally
            {
                _dte.StatusBar.Progress(false);
                _dte.StatusBar.Text = "Operation finished. Please see Output Window for details.";
                _isProcessing = false;
            }

        }
    }
}
