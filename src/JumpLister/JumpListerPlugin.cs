﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using BepInEx;
using BepInEx.Unity.IL2CPP;
using HarmonyLib;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;
using UnityEngine;

#pragma warning disable CA1416

namespace JumpLister;

[BepInPlugin(GUID, PluginName, Version)]
public class JumpListerPlugin : BasePlugin
{
    public const string GUID = "marco.jumplister";
    public const string PluginName = "JumpLister";
    public const string Version = "1.0.0";

    public override void Load()
    {
        if (!TaskbarManager.IsPlatformSupported)
        {
            Log.LogWarning("JumpLists are not supported on this platform. Windows 7 or later is required.");
            return;
        }

        var enabled = Config.Bind("Jump List", "Enable", true, "Add useful items to the game's Jump List (the list of actions you get when right-clicking on the game in your Windows Taskbar).\nThese actions will appear even if the game is turned off if you pin it to your taskbar. To remove them, turn off this setting.");
        var watch = Config.Bind("Jump List", "Show Recently Saved files", true, "Include a list of 10 most recently saved screenshots, cards and scenes (since game start) at the top of the game's Jump List.\nMight slightly affect performance.");

        if (enabled.Value) SetUpJumpList();
        enabled.SettingChanged += (sender, args) =>
        {
            if (enabled.Value)
            {
                SetUpJumpList();
                if (watch.Value) SetUpWatcher();
            }
            else if (_jumpList != null)
            {
                _jumpList.ClearAllUserTasks();
                _jumpList.Refresh();
                _watcher?.Dispose();
            }
        };

        if (watch.Value) SetUpWatcher();
        watch.SettingChanged += (sender, args) =>
        {
            if (watch.Value && enabled.Value)
            {
                SetUpWatcher();
            }
            else if (_watcher != null)
            {
                _watcher.Dispose();
                if (_recentsCategoryInnerList != null && _recentsCategoryInnerList.Count > 0)
                {
                    _recentsCategoryInnerList.Clear();
                    _Recents.Clear();
                    _jumpList.Refresh();
                }
            }
        };
    }

    public override bool Unload()
    {
        _watcher?.Dispose();
        return true;
    }

    #region JumpList

    private JumpList _jumpList;
    private ICollection<IJumpListItem> _recentsCategoryInnerList;

    // TODO?
    // - option to restart without mods? kill current instance and disable doorstop for a single start
    // - option to open studio if installed, or to open the game if inside studio
    private void SetUpJumpList()
    {
        _jumpList = JumpList.CreateJumpListForIndividualWindow(TaskbarManager.Instance.ApplicationId, Process.GetCurrentProcess().MainWindowHandle);

        // Built in recents doesn't work for some reason
        //JumpList.AddToRecent(Path.GetFullPath(Path.Combine(Paths.GameRootPath, "UserData/chara/female/HC_F_000.png")));
        _jumpList.KnownCategoryToDisplay = JumpListKnownCategoryType.Neither;

        // Create the category even if recents are disabled since it won't show up if it's empty. 
        var recentsCategory = new JumpListCustomCategory("Recently saved (show in explorer)");
        _recentsCategoryInnerList = (ICollection<IJumpListItem>)recentsCategory.GetType().GetProperty("JumpListItems", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)?.GetValue(recentsCategory) ?? throw new MemberNotFoundException("JumpListItems not found");

        var bepinexCategory = new JumpListCustomCategory("BepInEx");
        bepinexCategory.AddDirectoryOpen("Config folder", "BepInEx/config");
        bepinexCategory.AddDirectoryOpen("Plugins folder", "BepInEx/plugins");
        bepinexCategory.AddFileOpen("Open LogOutput.log", "BepInEx/LogOutput.log", true);
        bepinexCategory.AddFileOpen("Open ErrorLog.log", "BepInEx/ErrorLog.log", true);
        //bepinexCategory.AddFileShowInExplorer("test6", "UserData/chara/female/HC_F_000.png", true, new IconReference(Paths.ExecutablePath, 0));

        var userDataCategory = new JumpListCustomCategory("UserData folders");
        userDataCategory.AddDirectoryOpen("Screenshots", "UserData/cap");
        userDataCategory.AddDirectoryOpen("Characters", "UserData/chara");
        userDataCategory.AddDirectoryOpen("Coordinates", "UserData/coordinate");
        userDataCategory.AddDirectoryOpen("Studio scenes", "UserData/craft/scene", true);

        _jumpList.AddCustomCategories(recentsCategory, bepinexCategory, userDataCategory);
        _jumpList.Refresh();
    }

    #endregion

    #region Recently opened

    private static readonly List<KeyValuePair<string, JumpListLink>> _Recents = new();
    private FileSystemWatcher _watcher;
    private string _watchPath;

    private void SetUpWatcher()
    {
        _watcher?.Dispose();

        _watchPath = Path.GetFullPath(Path.Combine(Paths.GameRootPath, "UserData"));
        //Console.WriteLine("watch in " + _watchPath);

        _watcher = new FileSystemWatcher(_watchPath, "*.png");
        _watcher.IncludeSubdirectories = true;
        _watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName;
        _watcher.Changed += WatcherEvent; // Fires for both new files and overwrites
        _watcher.EnableRaisingEvents = true;
        _watcher.SynchronizingObject = null; // Run on threadpool
    }

    private void WatcherEvent(object sender, FileSystemEventArgs e)
    {
        try
        {
            // Do not catch file changes done outside of the game
            if (!Application.isFocused)
                return;

            if (_recentsCategoryInnerList == null)
                return;

            var fullPath = Path.GetFullPath(e.FullPath);

            var fi = new FileInfo(fullPath);
            var parentName = fi.Directory?.Name ?? "???";

            _recentsCategoryInnerList.Clear();

            _Recents.RemoveAll(x => string.Equals(x.Key, fullPath, StringComparison.OrdinalIgnoreCase));

            var link = new JumpListLink("explorer.exe", $"{parentName}/{fi.Name}");
            link.Arguments = $"/select,\"{fullPath}\"";
            link.WorkingDirectory = fi.DirectoryName;
            link.IconReference = new IconReference("shell32.dll", 22);

            _Recents.Insert(0, new KeyValuePair<string, JumpListLink>(fullPath, link));

            if (_Recents.Count > 10)
                _Recents.RemoveRange(10, _Recents.Count - 10);

            foreach (var recent in _Recents)
                _recentsCategoryInnerList.Add(recent.Value);

            // This is the only expensive part, takes around 120ms
            _jumpList.Refresh();
        }
        catch (Exception exception)
        {
            Log.LogError(exception);
        }
    }

    #endregion
}
