using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;

namespace ManagedShell.WindowsTasks
{
	public class Tasks : IDisposable
	{
		private readonly TasksService _tasksService;
		private ICollectionView groupedWindows;

		public ICollectionView GroupedWindows => groupedWindows;

		public Tasks(TasksService tasksService)
		{
			_tasksService = tasksService;
			// prepare collections
			groupedWindows = CollectionViewSource.GetDefaultView(_tasksService.Windows);
			groupedWindows.GroupDescriptions.Add(new PropertyGroupDescription(nameof(ApplicationWindow.Category)));
			groupedWindows.SortDescriptions.Add(new SortDescription(nameof(ApplicationWindow.Index),
				ListSortDirection.Ascending));
			groupedWindows.CollectionChanged += groupedWindows_Changed;
			groupedWindows.Filter = groupedWindows_Filter;

			if (groupedWindows is ICollectionViewLiveShaping taskbarItemsView)
			{
				taskbarItemsView.IsLiveFiltering = true;
				taskbarItemsView.LiveFilteringProperties.Add(nameof(ApplicationWindow.ShowInTaskbar));
				taskbarItemsView.IsLiveGrouping = true;
				taskbarItemsView.LiveGroupingProperties.Add(nameof(ApplicationWindow.Category));
				taskbarItemsView.IsLiveSorting = true;
				taskbarItemsView.LiveSortingProperties.Add(nameof(ApplicationWindow.Index));
			}
		}

		public void Initialize(ITaskCategoryProvider taskCategoryProvider)
		{
			if (!_tasksService.IsInitialized)
			{
				_tasksService.SetTaskCategoryProvider(taskCategoryProvider);
				Initialize();
			}
		}

		public void Move(IntPtr oldApplicationWindowHandle, IntPtr newApplicationWindowHandle)
		{
			var oldApplicationWindow =
				_tasksService.Windows.FirstOrDefault(e => e.Handle == oldApplicationWindowHandle);
			var newApplicationWindow =
				_tasksService.Windows.FirstOrDefault(e => e.Handle == newApplicationWindowHandle);
			if (oldApplicationWindow == null || newApplicationWindow == null)
			{
				return;
			}

			if (oldApplicationWindow == newApplicationWindow)
			{
				return;
			}

			if (oldApplicationWindow.Category != newApplicationWindow.Category)
			{
				return;
			}

			var tmp = oldApplicationWindow.Index;
			oldApplicationWindow.Index = newApplicationWindow.Index;
			newApplicationWindow.Index = tmp;
		}

		public void Initialize()
		{
			_tasksService.Initialize();
		}

		private void groupedWindows_Changed(object sender, NotifyCollectionChangedEventArgs e)
		{
			// yup, do nothing. helps prevent a NRE
		}

		private bool groupedWindows_Filter(object item)
		{
			return item is ApplicationWindow window && window.ShowInTaskbar;
		}

		public void Dispose()
		{
			_tasksService.Dispose();
		}
	}
}