using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[CreateAssetMenu(fileName = "Recipes.asset", menuName = "Data/Recipe Collection")]
public class RecipeCollection : ScriptableObject
{
    [System.Serializable]
    public struct Recipe {
        public InventoryItem[] ingredients;
        public InventoryItem product;

        public (InventoryItem, InventoryItem, InventoryItem) GetIngredientsTuple() {
            (InventoryItem, InventoryItem, InventoryItem) tuple;
            tuple.Item1 = ingredients.Length > 0 ? ingredients[0] : null;
            tuple.Item2 = ingredients.Length > 1 ? ingredients[1] : null;
            tuple.Item3 = ingredients.Length > 2 ? ingredients[2] : null;
            return tuple;
        }
    }

    static List<InventoryItem> _tempItems = new();

    [SerializeField]
    List<Recipe> _recipes = new();

    Dictionary<(InventoryItem, InventoryItem, InventoryItem), InventoryItem> _lookup = new();

    [SerializeField] InventoryItem defaultProduct;

    public bool TryAddRecipe(Dictionary<string, InventoryItem> items, string productName, params string[] ingredientNames) {
        _tempItems.Clear();
        foreach (var name in ingredientNames) {
            if (!string.IsNullOrWhiteSpace(name) && items.TryGetValue(name, out var item)) {
                _tempItems.Add(item);
            }
        }

        if (_tempItems.Count == 0) return false;

        _tempItems.Sort((a, b) => a.name.CompareTo(b.name));

        if (!items.TryGetValue(productName, out var product)) return false;

        var recipe = new Recipe() {
            ingredients = _tempItems.ToArray(),
            product = product
        };
        var tuple = recipe.GetIngredientsTuple();

        if (_lookup.ContainsKey(tuple)) {
            Debug.LogError($"Duplicate recipe detected: '{tuple.Item1?.name}' + '{tuple.Item2?.name}' + '{tuple.Item3?.name}' maps to both {_lookup[tuple].name} and {product.name}.\nOnly the first mapping will be kept.");
        } else {
            _recipes.Add(recipe);
            _lookup.Add(tuple, product);
        }

        return true;
    }

    public void Clear() {
        _recipes.Clear();
        _lookup.Clear();
    }

    void OnEnable() {
        foreach(var recipe in _recipes) {
            _lookup.Add(recipe.GetIngredientsTuple(), recipe.product);
        }
    }
}
