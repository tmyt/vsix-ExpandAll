using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace ExpandAll
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ExpandAll
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("f85837e9-9618-4b08-b7ed-289d6f4d1a2a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        private readonly DTE _DTE;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpandAll"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ExpandAll(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _DTE = package.GetService<DTE, DTE>();
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ExpandAll Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider => _package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ExpandAll's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ExpandAll(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var projects = new List<Project>();
            // Expand Project Roots
            var root = ((DTE2)_DTE).ToolWindows.SolutionExplorer.UIHierarchyItems;
            Walker.Walk(root, (UIHierarchyItem item) => item.UIHierarchyItems, item =>
            {
                item.UIHierarchyItems.Expanded = true;
                if (!(item.Object is Project proj) || proj.Kind == "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}") return true;
                projects.Add(proj);
                return false;
            });
            // Expand Project Items
            foreach (Project project in projects)
            {
                Walker.Walk(project.ProjectItems, (ProjectItem item) => item.ProjectItems, item =>
                {
                    if (item.Kind != "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") return false;
                    item.ExpandView();
                    return true;
                });
            }
        }
    }

    static class Walker
    {
        public static void Walk<T, U>(T source, Func<U, T> getChildren, Func<U, bool> walker)
        where T : IEnumerable
        {
            var queue = new Queue<T>();
            queue.Enqueue(source);
            while (queue.Count > 0)
            {
                var s = queue.Dequeue();
                if (s == null) continue;
                foreach (U i in s)
                {
                    if (!walker(i)) continue;
                    queue.Enqueue(getChildren(i));
                }
            }
        }
    }
}
