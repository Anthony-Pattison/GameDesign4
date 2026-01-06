using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;

[CustomPropertyDrawer(typeof(RecipeCollection.Recipe))]
public class RecipeDrawer : PropertyDrawer {
    
    public override float GetPropertyHeight(SerializedProperty property, GUIContent label)
    {
        if (property.isExpanded) return EditorGUI.GetPropertyHeight(property);
        return EditorGUIUtility.singleLineHeight;
    }

    public override void OnGUI(Rect position, SerializedProperty property, GUIContent label) {
        string name;
        {
            string progress = string.Empty;
            var child = property.FindPropertyRelative("ingredients");
            int count = child.arraySize;
            for (int i = 0; i < count; i++) {
                var entry = child.GetArrayElementAtIndex(i);
                if (entry.objectReferenceValue != null) {
                    string next = ((ScriptableObject)entry.objectReferenceValue).name;
                    if (progress.Length == 0) progress = next;
                    else progress = $"{progress} + {next}";
                }
            }

            child = property.FindPropertyRelative("product");
            string product = ((ScriptableObject)child.objectReferenceValue)?.name;        
            name = $"({progress}) => {product}";
        }

        EditorGUI.PropertyField(position, property, new GUIContent(name), true);
    }
}
