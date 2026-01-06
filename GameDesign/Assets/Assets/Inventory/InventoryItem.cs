using UnityEngine;

[CreateAssetMenu(fileName = "NewItem.asset", menuName = "Data/Inventory Item")]
public class InventoryItem : ScriptableObject
{
    public string displayName;

    [TexturePreview(5)]
    public Sprite sprite;

    public Rarity rarity;

    public int cost;

    public int Uses;

    public string Type;
}
