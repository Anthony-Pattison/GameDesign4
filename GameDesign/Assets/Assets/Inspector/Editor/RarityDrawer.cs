using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;

[CustomPropertyDrawer(typeof(Rarity))]
public class RarityDrawer : PropertyDrawer {

    static Color32[] _rarityColors = new[] {
        new Color32(64, 64, 64, 255),      // Unset
        new Color32(255, 255, 255, 255),// Common
        new Color32(0, 200, 64, 255),   // Uncommon
        new Color32(90, 160, 255, 255),   // Rare
        new Color32(200, 64, 255, 255),  // Epic
        new Color32(255, 180, 0, 255),  // Legendary
    };

    public override void OnGUI(Rect position, SerializedProperty property, GUIContent label) {

        var cache = GUI.color;
        GUI.color = _rarityColors[Mathf.Clamp(property.intValue, 0, _rarityColors.Length -1)];

        property.intValue = (int)(Rarity)EditorGUI.EnumPopup(position, label, (Rarity)property.intValue);

        GUI.color = cache;
    }
}
