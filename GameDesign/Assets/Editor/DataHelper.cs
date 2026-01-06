/*
These helpers package-up some common operations for writing Unity editor scripts.
They are provided under an MIT licence:

Copyright (c) 2026 Douglas Gregory

Permission is hereby granted, free of charge, to any person obtaining a copy of this 
software and associated documentation files (the "Software"), to deal in the Software 
without restriction, including without limitation the rights to use, copy, modify, 
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to 
permit persons to whom the Software is furnished to do so, subject to the following 
conditions:

The above copyright notice and this permission notice shall be included in all copies
or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR 
PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, 
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
*/

using System.Collections.Generic;
using UnityEngine;
using UnityEditor;

public static class DataHelper
{
    /// <summary>
    /// Marks any changes made to this object so they're correctly saved
    /// to disc the next time the scene / project is saved, and can be undone.
    /// </summary>
    /// <param name="target">The object being modified.</param>
    /// <param name="undoLabel">Optional: label to display in the Undo menu.</param>
    public static void MarkChangesForSaving(Object target, string undoLabel = "Data Import")
    {
        Undo.RecordObject(target, undoLabel);
        EditorUtility.SetDirty(target);
    }


    /// <summary>
    /// Searches the project for all assets of a given type and returns them in a dictionary, keyed by their file names.
    /// </summary>
    /// <typeparam name="T">The type of asset to find.</typeparam>
    /// <returns>A dictionary mapping asset names to asset references.</returns>
    public static Dictionary<string, T> GetAllAssetsOfType<T>() where T : Object {
        string[] guids = AssetDatabase.FindAssets($"t:{typeof(T).Name}");
        var collection = new Dictionary<string, T>(guids.Length);

        foreach (var guid in guids)
        {
            string path = AssetDatabase.GUIDToAssetPath(guid);
            T asset = AssetDatabase.LoadAssetAtPath<T>(path);
            collection.Add(asset.name, asset);
        }

        return collection;
    }

    /// <summary>
    /// Searches for an existing asset name in the dictionary and returns it, or if not found, creates a new asset by that name in the provided folder.
    /// </summary>
    /// <typeparam name="T">The type of asset to find / create.</typeparam>
    /// <param name="name">The name of the asset to find / create.</param>
    /// <param name="existing">A dictionary of all existing asset names and references - use GetAllAssetsOfType to build this.</param>
    /// <param name="folder">A path to put any newly-created assets.</param>
    /// <returns>A reference to the found / created asset.</returns>
    public static T GetOrCreateAsset<T>(string name, Dictionary<string, T> existing, string folder) where T : ScriptableObject {
        if (!existing.TryGetValue(name, out T asset)) {
            
            asset = ScriptableObject.CreateInstance<T>();

            if (folder.StartsWith("Assets/")) folder = folder.Substring(7);
            if (folder.EndsWith('/')) folder = folder.Substring(0, folder.Length - 1);

            string folderPath = $"Assets/{folder}";
            if (!AssetDatabase.IsValidFolder(folderPath)) {
                Debug.LogWarning($"Folder path '{folderPath}' does not exist - creating it now.");
                var folders = folder.Split('/');
                folderPath = "Assets";
                foreach(var newFolder in folders) {
                    AssetDatabase.CreateFolder(folderPath, newFolder);
                    folderPath = $"{folderPath}/{newFolder}";
                }                
            }
            AssetDatabase.CreateAsset(asset, $"{folderPath}/{name}.asset");
            existing.Add(name, asset);
        }

        MarkChangesForSaving(asset);
        return asset;
    }
}
