using System;
using BepInEx;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using BepInEx.Unity.IL2CPP;
using HarmonyLib;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;
using UnityEngine;

#pragma warning disable CA1416

namespace JumpLister
{
    [BepInPlugin(GUID, PluginName, Version)]
    public class JumpListerPlugin : BasePlugin
    {
        public const string GUID = "com.jumplister";
        public const string PluginName = "JumpLister";
        public const string Version = "1.0.0";

        private JumpList _jumpList;
        private ICollection<IJumpListItem> _recentsCategoryInnerList;

        public override void Load()
        {
            if (!TaskbarManager.IsPlatformSupported)
            {
                Log.LogWarning("JumpLists are not supported on this platform. Windows 7 or later is required.");
                return;
            }

            _jumpList = JumpList.CreateJumpListForIndividualWindow(TaskbarManager.Instance.ApplicationId, Process.GetCurrentProcess().MainWindowHandle);
            //todo screenshots jumpList.KnownCategoryToDisplay = JumpListKnownCategoryType.Recent;


            // <summary>
            // TODO:
            // - restart without mods? kill current instance and disable doorstop for a single start
            // - populate recents with taken screenshots, probably use filewatcher? or hook file save?
            // - - or populate with recently saves characters, coordinates, scenes, on click open in explorer and focus on the file
            // </summary>


            // Built in recents doesn't work for some reason
            //JumpList.AddToRecent(Path.GetFullPath(Path.Combine(Paths.GameRootPath, "UserData/chara/female/HC_F_000.png")));
            _jumpList.KnownCategoryToDisplay = JumpListKnownCategoryType.Neither;

            var recentsCategory = new JumpListCustomCategory("Recently saved (show in explorer)");
            _recentsCategoryInnerList = (ICollection<IJumpListItem>)recentsCategory.GetType().GetProperty("JumpListItems", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)?.GetValue(recentsCategory) ?? throw new MemberNotFoundException("JumpListItems not found");

            var bepinexCategory = new JumpListCustomCategory("BepInEx");
            bepinexCategory.AddDirectoryOpen("Config folder", "BepInEx/config");
            bepinexCategory.AddDirectoryOpen("Plugins folder", "BepInEx/plugins");
            bepinexCategory.AddFileOpen("Open LogOutput", "BepInEx/LogOutput.log", true);
            bepinexCategory.AddFileOpen("Open ErrorLog", "BepInEx/ErrorLog.log", true);
            //bepinexCategory.AddFileShowInExplorer("test6", "UserData/chara/female/HC_F_000.png", true, new IconReference(Paths.ExecutablePath, 0));

            var userDataCategory = new JumpListCustomCategory("UserData folders");
            userDataCategory.AddDirectoryOpen("Screenshots", "UserData/cap");
            userDataCategory.AddDirectoryOpen("Characters", "UserData/chara");
            userDataCategory.AddDirectoryOpen("Coordinates", "UserData/coordinate");
            userDataCategory.AddDirectoryOpen("Studio scenes", "UserData/craft/scene", true);

            _jumpList.AddCustomCategories(recentsCategory, bepinexCategory, userDataCategory);

            _watchPath = Path.GetFullPath(Path.Combine(Paths.GameRootPath, "UserData"));
            //Console.WriteLine("watch in " + _watchPath);
            _watcher = new FileSystemWatcher(_watchPath, "*.png");
            _watcher.IncludeSubdirectories = true;
            _watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName;
            _watcher.Changed += WatcherEvent; // Fires for both new files and overwrites
            _watcher.EnableRaisingEvents = true;
            _watcher.SynchronizingObject = null; // Run on threadpool
            _jumpList.Refresh();
        }

        public override bool Unload()
        {
            _watcher.Dispose();
            return true;
        }

        private static readonly List<KeyValuePair<string, JumpListLink>> _Recents = new();
        private FileSystemWatcher _watcher;
        private string _watchPath;

        private void WatcherEvent(object sender, FileSystemEventArgs e)
        {
            try
            {
                // Do not catch file changes done outside of the game
                if (!Application.isFocused)
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
    }

    internal static class JumpListExtensions
    {
        public static void AddFileShowInExplorer(this JumpListCustomCategory category, string displayName, string relativePath, bool ifExists = false, IconReference icon = default)
        {
            var fullPath = Path.Combine(Paths.GameRootPath, relativePath);

            if (ifExists && !File.Exists(fullPath))
                return;

            // Required or it won't open correctly
            fullPath = Path.GetFullPath(fullPath);

            AddCommand(category, displayName, "explorer.exe", $"/select,\"{fullPath}\"", null, icon);
        }

        public static void AddFileOpen(this JumpListCustomCategory category, string displayName, string relativePath, bool ifExists = false, IconReference icon = default)
        {
            var fullPath = Path.Combine(Paths.GameRootPath, relativePath);

            if (ifExists && !File.Exists(fullPath))
                return;

            fullPath = Path.GetFullPath(fullPath);

            AddCommand(category, displayName, fullPath, "", Path.GetDirectoryName(fullPath), icon);
        }

        public static void AddDirectoryOpen(this JumpListCustomCategory category, string displayName, string relativePath, bool ifExists = false, IconReference? icon = null)
        {
            var fullPath = Path.Combine(Paths.GameRootPath, relativePath);

            if (ifExists && !Directory.Exists(fullPath))
                return;

            // Required or it won't open correctly
            fullPath = Path.GetFullPath(fullPath);
            AddCommand(category, displayName, "explorer.exe", $"\"{fullPath}\"", fullPath, icon ?? new IconReference("shell32.dll", 3));
        }

        public static void AddCommand(this JumpListCustomCategory category, string displayName, string command, string arguments, string workDir = null, IconReference icon = default)
        {
            var capFolder = new JumpListLink(command, displayName);
            capFolder.Arguments = arguments;
            capFolder.WorkingDirectory = workDir ?? Path.GetDirectoryName(Path.GetFullPath(command));
            capFolder.IconReference = icon;

            category.AddJumpListItems(capFolder);
        }
    }
}
