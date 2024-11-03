using BepInEx;
using System.Collections;
using System.Diagnostics;
using System.IO;
using BepInEx.Unity.IL2CPP;
using BepInEx.Unity.IL2CPP.Utils.Collections;
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

        public override void Load()
        {
            // Do not run below Windows 7
            if (!TaskbarManager.IsPlatformSupported)
            {
                Log.LogWarning("JumpLists are not supported on this platform. Windows 7 or later is required.");
                return;
            }

            AddComponent<Delay>();
        }

        public class Delay : MonoBehaviour
        {
            /// <summary>
            /// TODO:
            /// - open log file
            /// - open bepinex\config and plugins folders
            /// - restart without mods? kill current instance and disable doorstop for a single start
            /// - populate recents with taken screenshots, probably use filewatcher? or hook file save?
            /// - - or populate with recently saves characters, coordinates, scenes, on click open in explorer and focus on the file
            /// </summary>


            void Start()
            {
                StartCoroutine(Co().WrapToIl2Cpp());
            }

            IEnumerator Co()
            {
                yield return null;
                yield return null;

                JumpList jumpList = JumpList.CreateJumpListForIndividualWindow(TaskbarManager.Instance.ApplicationId, Process.GetCurrentProcess().MainWindowHandle);
                //todo screenshots jumpList.KnownCategoryToDisplay = JumpListKnownCategoryType.Recent;

                // Add tasks for opening folders - UserData\cap, UserData\studio\scene, UserData\coordinate, UserData\chara
                var folderCat = new JumpListCustomCategory("UserData folders");
                AddPath("UserData/cap", "Screenshots");
                AddPath("UserData/chara", "Characters");
                AddPath("UserData/coordinate", "Coordinates");
                AddPath("UserData/studio/scene", "Studio scenes");
                jumpList.AddCustomCategories(folderCat);
                jumpList.Refresh();

                void AddPath(string relativePath, string folderName)
                {
                    var fullPath = Path.Combine(Paths.GameRootPath, relativePath);

                    if (!Directory.Exists(fullPath))
                        return;

                    // Required or it won't open correctly
                    fullPath = Path.GetFullPath(fullPath);

                    var capFolder = new JumpListLink("explorer.exe", folderName);
                    capFolder.IconReference = new IconReference("shell32.dll", 3);
                    capFolder.Arguments = $"\"{fullPath}\"";
                    //capFolder.Arguments = "";
                    //capFolder.WorkingDirectory = fullPath;
                    folderCat.AddJumpListItems(capFolder);
                }
            }
        }
    }
}
