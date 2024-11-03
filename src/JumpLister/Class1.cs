using BepInEx;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using BepInEx.Logging;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;
using UnityEngine;

namespace JumpLister
{
    [BepInPlugin(GUID, PluginName, Version)]
    public class JumpListerPlugin : BaseUnityPlugin
    {
        public const string GUID = "com.jumplister";
        public const string PluginName = "JumpLister";
        public const string Version = "1.0.0";

        internal static new ManualLogSource Logger;

        IEnumerator Start()
        {
            Logger = base.Logger;

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

                var capFolder = new JumpListLink(fullPath, folderName);
                capFolder.IconReference = new IconReference(fullPath, 0);
                capFolder.Arguments = "";
                capFolder.WorkingDirectory = fullPath;
                folderCat.AddJumpListItems(capFolder);
            }
        }
    }
}
