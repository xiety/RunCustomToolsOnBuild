using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSLangProj;

namespace Extensions
{
	[PackageRegistration(UseManagedResourcesOnly = true)]
	[Guid("a5dd5f4a-cf8c-4845-a8dd-b34a5e514ecc")]
	[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
	[ProvideObject(typeof(RunCustomToolsOnBuildPackage))]
	[ComVisible(true)]
	public class RunCustomToolsOnBuildPackage : Package
	{
		private DTE _dte;
		private BuildEvents _buildEvents;

		protected override void Initialize()
		{
			base.Initialize();

			_dte = GetService(typeof(SDTE)) as DTE;

			_buildEvents = _dte.Events.BuildEvents;
			_buildEvents.OnBuildBegin += OnBuildBegin;
		}

		private void OnBuildBegin(vsBuildScope scope, vsBuildAction action)
		{
			foreach (Project project in _dte.Solution.Projects)
			{
				Generate(project.ProjectItems);
			}
		}

		private void Generate(ProjectItems items)
		{
			if (items == null) return;

			foreach (ProjectItem item in items)
			{
				Generate(item);
			}
		}

		private void Generate(ProjectItem item)
		{
			switch (item.Kind)
			{
				case EnvDTE.Constants.vsProjectItemKindPhysicalFile:
					RunCustomTool(item);
					break;
				case EnvDTE.Constants.vsProjectItemKindPhysicalFolder:
					Generate(item.ProjectItems);
					break;
				case EnvDTE.Constants.vsProjectItemKindSolutionItems:
					Generate(item.SubProject?.ProjectItems);
					break;
			}
		}

		private void RunCustomTool(ProjectItem item)
		{
			var vsProjectItem = item.Object as VSProjectItem;

			if (vsProjectItem != null)
			{
				var property = item.Properties.Item("CustomTool");

				if (!String.IsNullOrEmpty(property.Value as string))
				{
					try
					{
						OutputMessage($"Running custom tool on file {item.Name}");

						vsProjectItem.RunCustomTool();
					}
					catch (Exception ex)
					{
						OutputMessage($"Exception while running custom tool on file {item.Name}: {ex.Message}");
					}
				}
			}
		}

		private void OutputMessage(string message)
		{
			var window = _dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
			var outputWindow = (OutputWindow)window.Object;
			outputWindow.ActivePane.Activate();
			outputWindow.ActivePane.OutputString(message + Environment.NewLine);
		}

		protected override int QueryClose(out bool canClose)
		{
			var result = base.QueryClose(out canClose);

			if (!canClose)
				return result;

			if (_buildEvents != null)
			{
				_buildEvents.OnBuildBegin -= OnBuildBegin;
				_buildEvents = null;
			}

			return result;
		}
	}
}
