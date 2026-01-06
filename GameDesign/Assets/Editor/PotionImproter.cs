using Codice.CM.Common.Selectors;
using OfficeOpenXml;
using System.Collections.Generic;
using Unity.Android.Gradle.Manifest;
using UnityEngine;
using static RecipeCollection;

[CreateAssetMenu(fileName = "PotionImproter", menuName = "Import/PotionImproter")]
public class PotionImproter : ScriptableObject
{
    [Tooltip("Path form assets to the asset spreadsheet")]
    public string ExelFilePath = "Editor/PotionCrafting.xlsx";
    [Tooltip("Asset to store the loaded recipies.")]
    public RecipeCollection recipes;

    [ContextMenu("Import")]
    public void Import()
    {
        var excel = new ExcelImporter(ExelFilePath);

        var item = DataHelper.GetAllAssetsOfType<InventoryItem>();

        ImprtIngredients(excel, "Ingredients", item);
        ImprtIngredients(excel, "Potions", item);

        ImprtRecipes(excel, item);
        Debug.Log("Importing completed");
    }

    void ImprtIngredients(ExcelImporter excel, string category, Dictionary<string, InventoryItem> items)
    {
        if (!excel.TryGetTable(category, out var table))
        {
            Debug.Log($"Could not find {category} table.");
            return;
        }

        for (int row = 1; row <= table.RowCount; row++)
        {
            string name = table.GetValue<string>(row, "Name");
            if (string.IsNullOrWhiteSpace(name)) continue;
            
            var item = DataHelper.GetOrCreateAsset(name, items, "Ingredients");
            if (table.HasColumn("Uses"))
                item.Uses = table.GetValue<int>(row, "Uses");
            if (table.HasColumn("Type"))
                item.Type = table.GetValue<string>(row, "Type");
            if (string.IsNullOrWhiteSpace(item.displayName))
                item.displayName = name;
            
            if (table.TryGetEnum<Rarity>(row, "Rarity", out var rarity))
                item.rarity = rarity;

            item.cost = table.GetValue<int>(row, "Cost");

        }
    }
    void ImprtRecipes(ExcelImporter excel, Dictionary<string, InventoryItem> items)
    {
        if (recipes == null)
        {
            Debug.Log($"Could not find {recipes} table.");
            return;
        }
        if (!excel.TryGetTable("Recipes", out var table))
        {
            Debug.Log($"Could not find'Recipes' table.");
            return;
        }

        DataHelper.MarkChangesForSaving(recipes);
        recipes.Clear();
        for (int row = 1; row <= table.RowCount; row++)
        {
            recipes.TryAddRecipe(
                items,
                table.GetValue<string>(row, "Potion"),
                table.GetValue<string>(row, "Item 1"),
                table.GetValue<string>(row, "Item 2"),
                table.GetValue<string>(row, "Item 3")
                );
        }
    }
}
