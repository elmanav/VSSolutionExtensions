using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.TemplateWizard;

namespace CustomWizardExtensions
{
	public class PreprocessWizard : IWizard
	{
		#region data

		private DTE _visualStudio;
		private DTE2 _visualStudio2;
		private string _solutionName;

		/// <summary>
		/// Путь к папке с проектом.
		/// </summary>
		private string _destinationPath;

		/// <summary>
		/// Путь к папке с исходниками.
		/// </summary>
		private string _sourcePath;
		private string _templateSourcePath;
		private bool _isSolution;
		private string _safeSolutionName;

		#endregion

		#region IWizard Members

		/// <summary>
		/// Runs custom wizard logic at the beginning of a template wizard run.
		/// </summary>
		/// <param name="automationObject">The automation object being used by the template wizard.</param><param name="replacementsDictionary">The list of standard parameters to be replaced.</param><param name="runKind">A <see cref="T:Microsoft.VisualStudio.TemplateWizard.WizardRunKind"/> indicating the type of wizard run.</param><param name="customParams">The custom parameters with which to perform parameter replacement in the project.</param>
		public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary,
			WizardRunKind runKind, object[] customParams)
		{
			_visualStudio = (DTE) automationObject;
			_visualStudio2 = automationObject as DTE2;

			WriteToLog($"{_visualStudio2 != null}");

			var builder = new StringBuilder();

			// TO DEBUG
			// replacement dictionary
			foreach (var pair in replacementsDictionary)
			{
				builder.AppendLine($"{pair.Key} - {pair.Value}");
			}
			builder.AppendLine(new string('-', 5));

			builder.AppendLine($"Wizard run kind: {runKind}");
			builder.AppendLine(new string('-', 5));
			// Custom params
			builder.AppendLine("Custom params");
			foreach (var customParam in customParams)
			{
				builder.AppendLine($"{customParam}");
			}
			// --END TO DEBUG --


			_templateSourcePath = customParams[0].ToString();

			//if (replacementsDictionary.ContainsKey("$destinationdirectory$"))
			//{
			//	_destinationDirectory = replacementsDictionary["$destinationdirectory$"];
			//}

			builder.AppendLine($"MultiProject: {runKind == WizardRunKind.AsMultiProject}");
			if (runKind == WizardRunKind.AsMultiProject)
			{
				
				_isSolution = true;
				_solutionName = replacementsDictionary["$projectname$"];
				_safeSolutionName = replacementsDictionary["$safeprojectname$"];
			}

			_destinationPath = replacementsDictionary["$destinationdirectory$"];
			
			MessageBox.Show(builder.ToString(), "Path caption", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		/// <summary>
		/// Runs custom wizard logic when a project has finished generating.
		/// </summary>
		/// <param name="project">The project that finished generating.</param>
		public void ProjectFinishedGenerating(Project project)
		{
		}

		/// <summary>
		/// Runs custom wizard logic when a project item has finished generating.
		/// </summary>
		/// <param name="projectItem">The project item that finished generating.</param>
		public void ProjectItemFinishedGenerating(ProjectItem projectItem)
		{
		}

		/// <summary>
		/// Indicates whether the specified project item should be added to the project.
		/// </summary>
		/// <returns>
		/// true if the project item should be added to the project; otherwise, false.
		/// </returns>
		/// <param name="filePath">The path to the project item.</param>
		public bool ShouldAddProjectItem(string filePath)
		{
			return true;
		}

		/// <summary>
		/// Runs custom wizard logic before opening an item in the template.
		/// </summary>
		/// <param name="projectItem">The project item that will be opened.</param>
		public void BeforeOpeningFile(ProjectItem projectItem)
		{
		}

		/// <summary>
		/// Runs custom wizard logic when the wizard has completed all tasks.
		/// </summary>
		void IWizard.RunFinished()
		{
			if (_isSolution)
			{
				try
				{
					CopySolutionItemsToDestinationFolder(); // копируем необходимые файлы
					MoveSolutionToSourceSubdirectory("src"); // переносим .sln в подкаталог \src
					WriteToLog("Solution finished");
				}
				catch (Exception e)
				{
					MessageBox.Show(e.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		#endregion

		private string GetSolutionFilePath(DTE vs)
		{
			return (string)vs.Solution.Properties.Item("Path").Value;
		}

		/// <summary>
		/// Copy the solution items (files, folders) to destination folder.
		/// </summary>
		private void CopySolutionItemsToDestinationFolder()
		{
			string sourceFromTemplateDirectory = Path.Combine(Path.GetDirectoryName(_templateSourcePath), "solution_items");
			CopyFolder(sourceFromTemplateDirectory, _destinationPath);
		}

		private void CopyFolder(string sourceDirectory, string destinationDirectory)
		{
			if (Directory.Exists(sourceDirectory))
			{
				string[] files = Directory.GetFiles(sourceDirectory);
				foreach (string file in files)
				{
					string name = Path.GetFileName(file);
					string dest = Path.Combine(destinationDirectory, name);
					File.Copy(file, dest);
					WriteToLog($"Copy file: {dest}");
				}
			}
		}

		/// <summary>
		/// Moves the solution to source subdirectory.
		/// </summary>
		/// <param name="subdirectory">The subdirectory.</param>
		private void MoveSolutionToSourceSubdirectory(string subdirectory)
		{
			string currentSolutionFilePath = GetSolutionFilePath(_visualStudio);
			string currentSolutionDirectory = Path.GetDirectoryName(currentSolutionFilePath);
			_sourcePath = Path.Combine(currentSolutionDirectory, subdirectory);
			CreateDirectory(_sourcePath);
			SaveSolutionToNewDirectoryAndDelete(_sourcePath);
		}

		private void SaveSolutionToNewDirectoryAndDelete(string destinationSolutionDirectory)
		{
			string sourceSolutionFilePath = GetSolutionFilePath(_visualStudio);
			string newFilename = Path.Combine(destinationSolutionDirectory, Path.GetFileName(sourceSolutionFilePath));
			_visualStudio.Solution.SaveAs(newFilename);

			File.Delete(sourceSolutionFilePath);
			string sourceSolutionDirectory = Path.GetDirectoryName(sourceSolutionFilePath);
			string suoFile = Path.GetFileNameWithoutExtension(sourceSolutionFilePath) + ".suo";
			string suoFilePath = Path.Combine(sourceSolutionDirectory, suoFile);
			File.Delete(suoFilePath);
		}

		private void CreateDirectory(string destinationSolutionDirectory)
		{
			if (!Directory.Exists(destinationSolutionDirectory))
			{
				Directory.CreateDirectory(destinationSolutionDirectory);
			}
		}
		
		private void WriteToLog(string text)
		{
			if (_visualStudio2 == null)
				MessageBox.Show(text);
			else
			{
				try
				{
					OutputWindow ow = _visualStudio2.ToolWindows.OutputWindow;
					var pane = ow.OutputWindowPanes.Item("General");
					
					pane.OutputString(text + Environment.NewLine);
				}
				catch (Exception e)
				{
					MessageBox.Show(e.ToString());
				}
			}
		}
	}

}