# ManagedShell
A library for creating Windows shell replacements using .NET, written in C#.

## Features
- Tasks service that provides taskbar functionality
- Tray service that provides notification area functionality
- AppBar WPF window class and helper methods
- Several helper classes for common shell functions
- Implements `INotifyPropertyChanged` and `ObservableCollection` to support binding with WPF

## Basic Usage
1. Add a reference to ManagedShell in your project.
2. Instantiate a `ManagedShell.ShellManager` object, optionally passing a `ManagedShell.ShellConfig` struct with custom configuration parameters.
3. Call the `ShellManager.Dispose()` method when your application's shutdown process begins, so that unmanaged resources can be freed, and to signal to AppBars that they may now close.

## Example implementations
- [RetroBar](https://github.com/dremin/RetroBar)